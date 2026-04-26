/**
 * Excel Custom Function: getMarketData(regionId, location1, [location2], ...)
 *
 * regionId: EVE region_id for ESI /markets/{region_id}/types
 * locations: one or more system/station ids passed as separate scalar arguments
 *
 * Returns a spilled table with one row per type_id.
 *
 * When multiple locations are supplied:
 *   - Volume / orderCount fields are SUMMED across locations.
 *   - All other (price-like) fields are WEIGHTED-AVERAGED across locations,
 *     weighted by each location's sell volume for that type.
 *     If total sell volume is 0 for a type, weights fall back to equal.
 *
 * @customfunction
 * @param regionId EVE region_id used to fetch the active type list from ESI.
 * @param locations One or more system/station ids (pass each as a separate argument).
 * @returns Spilled table with one header row and one row per typeId.
 */
async function getMarketData(regionId: number, ...locations: number[]): Promise<any[][]> {
  try {
    const locIds = normalizeLocations(locations);

    if (!regionId || isNaN(regionId)) {
      return [["ERROR", "regionId must be a number"]];
    }
    if (locIds.length === 0) {
      return [["ERROR", "Provide at least one system/station id (location)"]];
    }

    // 1) Get type IDs from ESI once (cached in-memory)
    const typeIds = await getEsiMarketTypeIds(regionId);

    // 2) Batch fuzzwork calls
    const CHUNK_SIZE = 100;
    const DELAY_MS = 50;
    const MAX_CONCURRENCY = 6;

    const header: string[] = [
      "typeId",
      "buyWeightedAverage",
      "buyMax",
      "buyMin",
      "buyStddev",
      "buyMedian",
      "buyVolume",
      "buyOrderCount",
      "buyPercentile",
      "sellWeightedAverage",
      "sellMax",
      "sellMin",
      "sellStddev",
      "sellMedian",
      "sellVolume",
      "sellOrderCount",
      "sellPercentile",
    ];

    // metric definitions in the same order as header[1..]
    // mode: "wavg" -> sell-volume weighted average; "sum" -> plain sum
    const metricFields: Array<{ side: "buy" | "sell"; key: string; mode: "wavg" | "sum" }> = [
      { side: "buy", key: "weightedAverage", mode: "wavg" },
      { side: "buy", key: "max", mode: "wavg" },
      { side: "buy", key: "min", mode: "wavg" },
      { side: "buy", key: "stddev", mode: "wavg" },
      { side: "buy", key: "median", mode: "wavg" },
      { side: "buy", key: "volume", mode: "sum" },
      { side: "buy", key: "orderCount", mode: "sum" },
      { side: "buy", key: "percentile", mode: "wavg" },
      { side: "sell", key: "weightedAverage", mode: "wavg" },
      { side: "sell", key: "max", mode: "wavg" },
      { side: "sell", key: "min", mode: "wavg" },
      { side: "sell", key: "stddev", mode: "wavg" },
      { side: "sell", key: "median", mode: "wavg" },
      { side: "sell", key: "volume", mode: "sum" },
      { side: "sell", key: "orderCount", mode: "sum" },
      { side: "sell", key: "percentile", mode: "wavg" },
    ];

    // Per-typeId accumulator: typeId -> locId -> { buy, sell }
    const perType: Map<number, Map<number, { buy: any; sell: any }>> = new Map();
    const errorRows: (string | number)[][] = [];

    const chunks = chunkArray(typeIds, CHUNK_SIZE);
    const tasks: Array<() => Promise<void>> = [];

    for (const loc of locIds) {
      for (const chunk of chunks) {
        tasks.push(async () => {
          try {
            const data = await fetchFuzzworkAggregates(loc, chunk);
            for (const typeIdStr of Object.keys(data)) {
              const rec = data[typeIdStr];
              if (!rec || !rec.buy || !rec.sell) continue;

              const tid = toInt(typeIdStr);
              let byLoc = perType.get(tid);
              if (!byLoc) {
                byLoc = new Map();
                perType.set(tid, byLoc);
              }
              byLoc.set(loc, { buy: rec.buy, sell: rec.sell });
            }
          } catch (e: any) {
            const errRow: (string | number)[] = new Array(header.length).fill(0);
            errRow[0] = `ERROR loc ${loc}: ${e ?.message ?? String(e)}`;
            errorRows.push(errRow);
          }
          await sleep(DELAY_MS);
        });
      }
    }

    await runWithConcurrency(tasks, MAX_CONCURRENCY);

    // 3) Combine across locations
    const rows: (string | number)[][] = [header];
    const sortedTypeIds = Array.from(perType.keys()).sort((a, b) => a - b);

    for (const tid of sortedTypeIds) {
      const byLoc = perType.get(tid)!;
      const entries = Array.from(byLoc.values());

      // weights = each location's sell volume for this type
      const weights = entries.map((e) => toNum(e.sell ?.volume));
      const totalW = weights.reduce((a, b) => a + b, 0);
      const useEqual = totalW <= 0;
      const effWeights = useEqual ? new Array(entries.length).fill(1 / entries.length) : weights.map((w) => w / totalW);

      const row: (string | number)[] = [tid];
      for (const f of metricFields) {
        let acc = 0;
        if (f.mode === "sum") {
          for (const e of entries) {
            acc += toNum((e as any)[f.side] ?.[f.key]);
          }
        } else {
          // "wavg"
          for (let i = 0; i < entries.length; i++) {
            const side = (entries[i] as any)[f.side];
            acc += toNum(side ?.[f.key]) * effWeights[i];
          }
        }
        row.push(acc);
      }
      rows.push(row);
    }

    for (const er of errorRows) rows.push(er);

    return rows;
  } catch (e: any) {
    return [["ERROR", String(e ?.message ?? e)]];
  }
}

/**
 * Excel Custom Function: getBPCMarket(regionId)
 *
 * Scans every public ItemExchange contract in the given region and returns
 * the lowest contract-price-PER-RUN for each Blueprint Copy (BPC) typeId.
 *
 * A contract qualifies when every included item is a BPC (`is_blueprint_copy
 * === true`) of the same `type_id` with `runs > 0`. This admits both single-
 * BPC contracts and bundles of multiple BPCs of the same blueprint, while
 * rejecting contracts that mix in skillbooks, modules, or other types.
 *
 * For a qualifying contract:
 *   totalRuns    = Σ (item.runs × item.quantity) over included items
 *   pricePerRun  = contract.price / totalRuns
 *
 * The output keeps the minimum pricePerRun seen per typeId.
 *
 * Endpoints:
 *   GET /contracts/public/{region_id}            (paginated; X-Pages header)
 *   GET /contracts/public/items/{contract_id}    (per contract, paginated)
 *
 * Returns a spilled table:
 *   header: ["typeId", "price"]    (price column = ISK per run)
 *   rows:   one per BPC typeId, sorted ascending by typeId.
 *
 * @customfunction
 * @param regionId EVE region_id whose public contracts will be scanned.
 * @returns Spilled table of [typeId, lowest price-per-run] for each BPC.
 */
async function getBPCMarket(regionId: number): Promise<any[][]> {
  try {
    if (!regionId || isNaN(regionId)) {
      return [["ERROR", "regionId must be a number"]];
    }

    const header: (string | number)[] = ["typeId", "price"];

    // 1) Pull all public contracts in the region.
    //    Include item_exchange (price > 0) AND auctions where the seller
    //    has set a buyout (we can actually pay that price right now).
    //    Auction `price` is the starting bid, not a guaranteed sale price,
    //    so plain auctions without buyout are skipped.
    const contracts = await fetchAllPublicContracts(regionId);
    const sellable = contracts
      .map((c: any) => {
        if (!c) return null;
        if (c.type === "item_exchange") {
          const p = toNum(c.price);
          return p > 0 ? { c, effPrice: p } : null;
        }
        if (c.type === "auction") {
          const b = toNum(c.buyout);
          return b > 0 ? { c, effPrice: b } : null;
        }
        return null;
      })
      .filter((x: any) => x !== null) as Array<{ c: any; effPrice: number }>;

    if (sellable.length === 0) {
      return [header];
    }

    // 2) Fetch items for each contract and reduce to lowest price-per-run.
    //    ESI has no per-request rate limit — only an error-rate limit
    //    (~100 errors per 60s window via X-Esi-Error-Limit-Remain). The
    //    fetchContractItems helper auto-pauses when that budget gets low,
    //    so we don't need a per-call sleep here. Concurrency is the only
    //    thing keeping us polite.
    const ITEMS_CONCURRENCY = 20;

    // typeId -> lowest price-per-run seen
    const bestByType: Map<number, number> = new Map();
    let errorCount = 0;

    const tasks: Array<() => Promise<void>> = sellable.map((s: { c: any; effPrice: number }) => async () => {
      const c = s.c;
      const effPrice = s.effPrice;
      try {
        const items = await fetchContractItems(c.contract_id);
        const included = items.filter((it: any) => it && it.is_included === true);
        if (included.length === 0) return;

        // All included items must be blueprint copies of the SAME typeId
        // with runs > 0. This admits bundles (e.g. a seller dropped 5 BPCs
        // of the same blueprint into one contract) while still rejecting
        // contracts that mix in skillbooks, modules, or other types.
        const firstTid = toInt(included[0].type_id);
        if (!firstTid) return;

        let totalRuns = 0;
        let allValid = true;
        for (const it of included) {
          const isBPC = it.is_blueprint_copy === true;
          const itRuns = toInt(it.runs);
          const itQty = toInt(it.quantity);
          const itTid = toInt(it.type_id);
          // BPCs do not stack, so a real BPC item is always quantity=1.
          // Anything else is suspicious and we refuse to guess.
          if (!isBPC || itRuns <= 0 || itQty !== 1 || itTid !== firstTid) {
            allValid = false;
            break;
          }
          totalRuns += itRuns;
        }
        if (!allValid || totalRuns <= 0) return;

        const pricePerRun = effPrice / totalRuns;
        if (!Number.isFinite(pricePerRun) || pricePerRun <= 0) return;

        const cur = bestByType.get(firstTid);
        if (cur === undefined || pricePerRun < cur) {
          bestByType.set(firstTid, pricePerRun);
        }
      } catch {
        // Per-contract failures are common (expired contracts, transient
        // ESI hiccups). Count them silently — we never want to spill
        // non-numeric error rows into the output, because downstream Power
        // Query coerces typeId to Int64 and the whole table fails.
        errorCount++;
      }
    });

    await runWithConcurrency(tasks, ITEMS_CONCURRENCY);

    // 3) Build spilled output (numeric only — see comment above)
    const rows: (string | number)[][] = [header];
    const sortedTypeIds = Array.from(bestByType.keys()).sort((a, b) => a - b);
    for (const tid of sortedTypeIds) {
      rows.push([tid, bestByType.get(tid)!]);
    }

    return rows;
  } catch (e: any) {
    return [["ERROR", String(e ?.message ?? e)]];
  }
}

/**
 * Excel Custom Function: getBPCDebug(regionId, typeId)
 *
 * Diagnostic helper. Scans every public contract in `regionId` and returns
 * a spilled table of EVERY contract whose included items touch `typeId`,
 * regardless of contract type or whether `getBPCMarket` would have used it.
 * Use this when `getBPCMarket` reports a price that doesn't match what you
 * see in-game — the rows here show exactly what ESI told us, so you can
 * confirm whether the cheap contract is missing entirely (ESI lag), is an
 * auction (we currently exclude auctions), is a bundle with extra items
 * (rejected by `getBPCMarket`'s same-type rule), etc.
 *
 * Columns:
 *   contractId, contractType, price, buyout, dateExpired,
 *   includedCount, sameTypeBpcOnly, totalRuns, pricePerRun, notes
 *
 * @customfunction
 * @param regionId EVE region_id whose public contracts will be scanned.
 * @param typeId Blueprint typeId to investigate.
 * @returns Spilled diagnostic table.
 */
async function getBPCDebug(regionId: number, typeId: number): Promise<any[][]> {
  try {
    if (!regionId || isNaN(regionId)) return [["ERROR", "regionId must be a number"]];
    if (!typeId || isNaN(typeId)) return [["ERROR", "typeId must be a number"]];

    const header: (string | number)[] = [
      "contractId", "contractType", "price", "buyout", "dateExpired",
      "includedCount", "sameTypeBpcOnly", "totalRuns", "pricePerRun", "notes"
    ];

    const contracts = await fetchAllPublicContracts(regionId);
    if (contracts.length === 0) return [header];

    const ITEMS_CONCURRENCY = 20;
    const out: (string | number)[][] = [];
    let scannedItems = 0;
    let itemsErrors = 0;

    const tasks: Array<() => Promise<void>> = contracts.map((c: any) => async () => {
      try {
        const items = await fetchContractItems(c.contract_id);
        scannedItems++;
        if (!items || items.length === 0) return;

        const included = items.filter((it: any) => it && it.is_included === true);
        const touches = items.some((it: any) => it && toInt(it.type_id) === typeId);
        if (!touches) return;

        // Build a concise per-item summary
        const itemNotes = items.map((it: any) =>
          `${it.is_included ? "+" : "-"}` +
          `tid=${toInt(it.type_id)}` +
          `,qty=${toInt(it.quantity)}` +
          `,runs=${toInt(it.runs)}` +
          `,bpc=${it.is_blueprint_copy === true ? "1" : "0"}`
        ).join(" | ");

        // Replicate getBPCMarket's same-type-BPC-only rule
        let sameTypeBpcOnly = "yes";
        let totalRuns = 0;
        if (included.length === 0) {
          sameTypeBpcOnly = "no(noIncluded)";
        } else {
          for (const it of included) {
            const isBPC = it.is_blueprint_copy === true;
            const itTid = toInt(it.type_id);
            const itRuns = toInt(it.runs);
            const itQty = toInt(it.quantity);
            if (!isBPC) { sameTypeBpcOnly = "no(notBPC)"; totalRuns = 0; break; }
            if (itTid !== typeId) { sameTypeBpcOnly = `no(otherType=${itTid})`; totalRuns = 0; break; }
            if (itRuns <= 0)  { sameTypeBpcOnly = "no(runs<=0)"; totalRuns = 0; break; }
            if (itQty <= 0)   { sameTypeBpcOnly = "no(qty<=0)"; totalRuns = 0; break; }
            totalRuns += itRuns * itQty;
          }
        }

        const price = toNum(c.price);
        const buyout = toNum(c.buyout);
        const ppr = totalRuns > 0 ? price / totalRuns : 0;

        out.push([
          toInt(c.contract_id),
          String(c.type ?? ""),
          price,
          buyout,
          String(c.date_expired ?? ""),
          included.length,
          sameTypeBpcOnly,
          totalRuns,
          ppr,
          itemNotes
        ]);
      } catch (e: any) {
        itemsErrors++;
      }
    });

    await runWithConcurrency(tasks, ITEMS_CONCURRENCY);

    // Sort by pricePerRun ascending (0 = unknown sinks to bottom)
    out.sort((a, b) => {
      const pa = toNum(a[8]) || Number.POSITIVE_INFINITY;
      const pb = toNum(b[8]) || Number.POSITIVE_INFINITY;
      return pa - pb;
    });

    const summary: (string | number)[] = [
      0, "_summary",
      contracts.length, scannedItems, "",
      0, `itemsErrors=${itemsErrors}`,
      0, 0,
      `Scanned ${contracts.length} contracts; ${out.length} touched typeId ${typeId}`
    ];

    return [header, summary, ...out];
  } catch (e: any) {
    return [["ERROR", String(e ?.message ?? e)]];
  }
}

/* -------------------- Helpers -------------------- */

function normalizeLocations(input: any): number[] {
  const flat: number[] = [];

  const visit = (v: any) => {
    if (v == null) return;
    if (Array.isArray(v)) {
      for (const x of v) visit(x);
      return;
    }
    if (typeof v === "number") {
      if (Number.isFinite(v)) flat.push(v);
      return;
    }
    if (typeof v === "string") {
      for (const part of v.split(/[,;\s]+/).filter(Boolean)) {
        const n = parseInt(part, 10);
        if (Number.isFinite(n)) flat.push(n);
      }
      return;
    }
  };

  visit(input);
  return Array.from(new Set(flat));
}

function chunkArray<T>(arr: T[], size: number): T[][] {
  const out: T[][] = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

function sleep(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function toNum(v: any): number {
  const n = typeof v === "number" ? v : parseFloat(String(v));
  return Number.isFinite(n) ? n : 0;
}

function toInt(v: any): number {
  const n = typeof v === "number" ? v : parseInt(String(v), 10);
  return Number.isFinite(n) ? n : 0;
}

/* -------------------- ESI: error-limit guard --------------------
 * ESI doesn't rate-limit by request count. Instead, every endpoint
 * shares a per-IP error budget tracked in two response headers:
 *   X-Esi-Error-Limit-Remain : errors left in the current window
 *   X-Esi-Error-Limit-Reset  : seconds until the window resets
 * Burning the budget to 0 yields a 420 ban for the rest of the window.
 *
 * Strategy: every successful and failed ESI response feeds these headers
 * into a shared latch. If Remain ever drops to <= 2, the next ESI call
 * blocks for `Reset` seconds before firing — across ALL concurrent
 * workers — so we recover gracefully instead of nuking the budget.
 *
 * Successful 200s do NOT cost any budget; only error responses do.
 */
let esiPauseUntilMs = 0;

async function esiThrottleWait(): Promise<void> {
  const now = Date.now();
  if (now < esiPauseUntilMs) {
    await sleep(esiPauseUntilMs - now);
  }
}

function esiObserveHeaders(resp: Response): void {
  const remainHdr = resp.headers.get("X-Esi-Error-Limit-Remain")
                 ?? resp.headers.get("x-esi-error-limit-remain");
  const resetHdr  = resp.headers.get("X-Esi-Error-Limit-Reset")
                 ?? resp.headers.get("x-esi-error-limit-reset");
  if (remainHdr === null || resetHdr === null) return; // CORS hid them

  const remain = toInt(remainHdr);
  const reset  = toInt(resetHdr);
  if (remain <= 2 && reset > 0) {
    // +500ms safety margin to make sure we wake AFTER the window flips.
    const newPauseUntil = Date.now() + reset * 1000 + 500;
    if (newPauseUntil > esiPauseUntilMs) {
      esiPauseUntilMs = newPauseUntil;
    }
  }
}

/** Wrapper around fetch() that respects the shared ESI error-budget latch. */
async function esiFetch(url: string, init?: RequestInit): Promise<Response> {
  await esiThrottleWait();
  const resp = await fetch(url, init);
  esiObserveHeaders(resp);
  return resp;
}

/* -------------------- ESI: market type IDs -------------------- */

type CacheEntry = { ts: number; value: number[] };
const esiTypeCache: Record<string, CacheEntry> = {};
const ESI_CACHE_TTL_MS = 10 * 60 * 1000;

async function getEsiMarketTypeIds(regionId: number): Promise<number[]> {
  const key = String(regionId);
  const now = Date.now();

  const cached = esiTypeCache[key];
  if (cached && now - cached.ts < ESI_CACHE_TTL_MS) return cached.value;

  const base = `https://esi.evetech.net/latest/markets/${regionId}/types/?datasource=tranquility&page=`;
  const PAGE_SIZE = 1000;
  const MAX_PAGES = 50;

  const all: number[] = [];
  for (let p = 1; p <= MAX_PAGES; p++) {
    const resp = await esiFetch(base + String(p), {
      method: "GET",
      headers: { Accept: "application/json" },
    });

    if (resp.status === 404) break;
    if (!resp.ok) {
      throw new Error(`ESI error (page ${p}): ${resp.status} ${resp.statusText}`);
    }

    const data = (await resp.json()) as number[];
    if (!Array.isArray(data) || data.length === 0) break;

    all.push(...data);

    if (data.length < PAGE_SIZE) break;
  }

  const unique = Array.from(new Set(all));
  esiTypeCache[key] = { ts: now, value: unique };
  return unique;
}

/* -------------------- ESI: public contracts -------------------- */

const ESI_COMPAT_DATE = "2025-12-16";

/**
 * GET /contracts/public/{region_id}, following X-Pages.
 */
async function fetchAllPublicContracts(regionId: number): Promise<any[]> {
  const base = `https://esi.evetech.net/contracts/public/${regionId}`;
  return await esiGetAllPages(base);
}

/**
 * GET /contracts/public/items/{contract_id}, following X-Pages.
 * 204 = no items (returns []).
 */
async function fetchContractItems(contractId: number): Promise<any[]> {
  const base = `https://esi.evetech.net/contracts/public/items/${contractId}`;
  return await esiGetAllPages(base);
}

/**
 * Generic paginated GET against an ESI endpoint that returns a JSON array
 * and exposes total page count via the X-Pages response header.
 */

async function esiGetAllPages(baseUrl: string): Promise<any[]> {
  // Send compat date as a query param to avoid CORS preflight on custom header.
  const headers: Record<string, string> = {
    Accept: "application/json",
  };

  const withCompat = (url: string, page: number) => {
    const sep = url.includes("?") ? "&" : "?";
    return `${url}${sep}page=${page}&compatibility_date=${ESI_COMPAT_DATE}`;
  };

  // Some ESI endpoints occasionally return 200 with an empty body or
  // malformed JSON (especially items for stale/expired contracts). Treat
  // any parse failure on a successful response as an empty page rather
  // than letting it bubble up and abort the whole scan.
  const safeJsonArray = async (resp: Response): Promise<any[]> => {
    const text = await resp.text();
    if (!text || text.trim().length === 0) return [];
    try {
      const parsed = JSON.parse(text);
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  };

  // IMPORTANT: We can't trust the X-Pages header. Office Scripts runs in a
  // CORS context, and unless the server sends
  //   Access-Control-Expose-Headers: X-Pages
  // resp.headers.get("X-Pages") returns null and we'd silently stop after
  // page 1 — which for The Forge's contracts endpoint cuts us off at the
  // first 1000 contracts (and all the cheap stuff is buried deeper).
  //
  // Strategy: walk pages sequentially starting at 1. Stop when ESI returns
  // 204/404 OR an empty array OR a short page (< typical 1000-row page).
  // We still try X-Pages first as a fast-path optimization for non-CORS
  // hosts, but only use it when it's a sane positive number.

  const PAGE_SIZE_HINT = 1000;
  const HARD_PAGE_CAP = 200; // safety net: 200k contracts per region max

  // ---- Page 1 ----
  const first = await esiFetch(withCompat(baseUrl, 1), { method: "GET", headers });
  if (first.status === 204 || first.status === 404) return [];
  if (!first.ok) throw new Error(`ESI error: ${first.status} ${first.statusText}`);

  const firstBody = await safeJsonArray(first);
  const all: any[] = firstBody.slice();

  if (firstBody.length === 0) return all;

  const pagesHdr = first.headers.get("X-Pages") ?? first.headers.get("x-pages");
  const headerPages = toInt(pagesHdr);
  // Treat a missing/zero header as "unknown" rather than "1 page".
  const knownTotal = headerPages > 0 ? headerPages : null;

  // If we know the total, fetch the rest in parallel.
  if (knownTotal !== null) {
    if (knownTotal <= 1) return all;
    const PAGE_CONCURRENCY = 20;
    const pageTasks: Array<() => Promise<void>> = [];
    for (let p = 2; p <= knownTotal; p++) {
      const pageNum = p;
      pageTasks.push(async () => {
        const resp = await esiFetch(withCompat(baseUrl, pageNum), { method: "GET", headers });
        if (resp.status === 204 || resp.status === 404) return;
        if (!resp.ok) {
          throw new Error(`ESI error (page ${pageNum}): ${resp.status} ${resp.statusText}`);
        }
        const data = await safeJsonArray(resp);
        if (data.length > 0) all.push(...data);
      });
    }
    await runWithConcurrency(pageTasks, PAGE_CONCURRENCY);
    return all;
  }

  // Unknown total (CORS hid X-Pages). Walk pages until we run out.
  // Page 1 was full -> there might be more. Page 1 was short -> we're done.
  if (firstBody.length < PAGE_SIZE_HINT) return all;

  for (let p = 2; p <= HARD_PAGE_CAP; p++) {
    const resp = await esiFetch(withCompat(baseUrl, p), { method: "GET", headers });
    if (resp.status === 204 || resp.status === 404) break;
    if (!resp.ok) {
      throw new Error(`ESI error (page ${p}): ${resp.status} ${resp.statusText}`);
    }
    const data = await safeJsonArray(resp);
    if (data.length === 0) break;
    all.push(...data);
    if (data.length < PAGE_SIZE_HINT) break;
  }

  return all;
}

/* -------------------- Fuzzwork aggregates -------------------- */

async function fetchFuzzworkAggregates(locationId: number, typeIds: number[]): Promise<any> {
  const typesCsv = typeIds.join(",");
  const url = `https://market.fuzzwork.co.uk/aggregates/?region=${locationId}&types=${encodeURIComponent(typesCsv)}`;

  const resp = await fetch(url, { method: "GET", headers: { Accept: "application/json" } });
  if (!resp.ok) {
    throw new Error(`Fuzzwork error: ${resp.status} ${resp.statusText}`);
  }
  return await resp.json();
}

/* -------------------- Concurrency runner -------------------- */

async function runWithConcurrency(tasks: Array<() => Promise<void>>, max: number): Promise<void> {
  let i = 0;
  const workers = new Array(Math.max(1, max)).fill(0).map(async () => {
    while (i < tasks.length) {
      const idx = i++;
      await tasks[idx]();
    }
  });
  await Promise.all(workers);
}