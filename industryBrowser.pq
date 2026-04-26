let
    /* ============================================================
       industryBrowser.pq  (renamed from ownedBPs.pq)

       NOTE: Only the Power Query was renamed (ownedBPs ->
       industryBrowser). The Excel side keeps its original names:
         - Sheet:  ownedBPs        (formula lives in ownedBPs!B2)
         - Table:  ownedBPsTable   (referenced below)
       This is intentional so no Excel Name Manager / sheet rename
       is required.

       For each blueprint row { typeId, ME, TE, runs } emit:
         - Per-stage ISK/h (mfg, copy, invent-as-sell-BPC), each isolated
           and using the row's actual ME/TE plus the user's selected
           decryptor.
         - For T2-producing rows, three full-cycle ISK/h paths and the
           best of them:
              (a) invent T2 BPC ourselves
              (b) buy T2 BPC from contracts
              (c) the row IS itself a T2 BPO (rare)
         - bestT2Path / bestT2IskPerHour pick the winner.

       Row sources combined:
         - "owned"        : rows from ownedBPsTable (the user's BPs).
                            Owned BPCs are treated as already paid for
                            (no acquisition cost).
         - "boughtBpc"    : every BPC currently listed in jitaContracts
                            that isn't already in ownedBPsTable.
                            Synthesized at the user-configured
                            "bought BPC" ME/TE/runs. These rows MUST
                            pay for the BPC: per-run mfg profit
                            subtracts bpcMarketPrice/runs. If no
                            contract price exists -> mfg result is
                            null (the row is unbuildable).
         Dedup key: { typeId, ME, TE, runs } so a user-owned BPC at
         non-default ME/TE doesn't collide with the synthetic row,
         and the owned variant takes precedence.

       Best-product selection:
         Invention (sell BPCs) and the T2-build paths each iterate over
         every entry in inventionProductsRaw, compute lifecycle ISK/h
         for that one chosen product, and pick the maximum. The chosen
         productId is exposed in inventionBestProductId,
         t2BuildPath_invent_productId, t2BuildPath_buyBpc_productId,
         and bestT2ProductId.

       EVE formulas implemented:
         EIV       Σ qty × adjustedPrice (CCP daily, from ESI)
         JobFee    EIV × runs × cfgRate
         MatQty(N) max(N, ceil(round(N × baseQty × (1 - ME/100), 2)))
         Time      baseTime × (1 - TE/100) × (1 + cfgDurMod)
         SellMult  (1 - tax) × (1 - broker) × (1 + sellModifier)
       ============================================================ */

    /* -------- 1. Source --------
       ownedBPsTable now has 4 columns:
         Column1 = typeId
         Column2 = ME
         Column3 = TE
         Column4 = runs   (-1 for BPO, positive int for BPC)
    */
    Source  = Excel.CurrentWorkbook(){[Name="ownedBPsTable"]}[Content],
    Typed   = Table.TransformColumnTypes(Source,
                {{"Column1", Int64.Type},{"Column2", Int64.Type},
                 {"Column3", Int64.Type},{"Column4", Int64.Type}}),
    Renamed = Table.RenameColumns(Typed,
                {{"Column1","typeId"},{"Column2","ME"},
                 {"Column3","TE"},{"Column4","runs"}}),

    /* -------- 2. Config -------- */
    ConfigTbl  = Excel.CurrentWorkbook(){[Name="config"]}[Content],
    ConfigRow0 = ConfigTbl{0},
    NormKey    = (k as text) => Text.Lower(Text.Trim(k)),
    ConfigRow  = Record.FromList(
                    Record.FieldValues(ConfigRow0),
                    List.Transform(Record.FieldNames(ConfigRow0), NormKey)),
    GetCfg     = (k as text) => Record.FieldOrDefault(ConfigRow, NormKey(k), null),

    ToPct = (v as any) as number =>
        if v = null then 0
        else if v is number then v
        else
            let
                s     = Text.Trim(Text.From(v)),
                isPct = Text.EndsWith(s, "%"),
                core  = if isPct then Text.Trim(Text.RemoveRange(s, Text.Length(s)-1, 1)) else s,
                n     = try Number.FromText(core) otherwise 0
            in
                if isPct then n / 100 else n,

    ToNum = (v as any) as number =>
        if v = null then 0
        else if v is number then v
        else try Number.FromText(Text.Trim(Text.From(v))) otherwise 0,

    Lang        = let v = GetCfg("language") in if v = null then "en" else Text.Lower(Text.From(v)),
    BuyMethod   = GetCfg("buy method"),
    SellMethod  = GetCfg("sell method"),
    TaxPct      = ToPct(GetCfg("tax %")),
    BrokerPct   = ToPct(GetCfg("broker %")),
    SellMod     = ToPct(GetCfg("sell modifier")),
    SellMult    = (1 - TaxPct) * (1 - BrokerPct) * (1 + SellMod),

    BuyIsOrder  = if BuyMethod = null then false
                  else Text.StartsWith(Text.Lower(Text.From(BuyMethod)), "buy"),
    BuyOrderMult= if BuyIsOrder then 1 / (1 - BrokerPct) else 1,

    /* Activity modifiers (negative dur = faster). cfgRate = effective
       combined EIV multiplier supplied by user.
       Keys match the config table rows (prefix "additional " was added). */
    MfgDurMod   = ToPct(GetCfg("additional maunfacture duration")),
    CpyDurMod   = ToPct(GetCfg("additional copy duration")),
    InvDurMod   = ToPct(GetCfg("invention duration")),
    MfgRate     = ToPct(GetCfg("additional maunfacture cost")),
    CpyRate     = ToPct(GetCfg("additional copy cost")),
    InvRate     = ToPct(GetCfg("invention cost")),
    /* Reactions reuse manufacture rate / duration unless caller adds
       dedicated config keys later. */
    RxnDurMod   = MfgDurMod,
    RxnRate     = MfgRate,

    /* Assumed metadata for BPCs purchased from contracts. */
    BoughtBpcRuns = let v = GetCfg("bought bpc runs") in if v = null then 10 else ToNum(v),
    BoughtBpcME   = let v = GetCfg("bought bpc me")   in if v = null then 2  else ToNum(v),
    BoughtBpcTE   = let v = GetCfg("bought bpc te")   in if v = null then 4  else ToNum(v),

    /* Minimum SVR (sell volume ratio):
         SVR = market_daily_sellVolume / my_daily_output
       where my_daily_output = productQtyPerRun × (86400 / mfgTimeSec_adjusted).
       SVR >= MinSvr means the market sells at least MinSvr times what I can
       produce per day — I won't saturate it.
       0 = disabled (all rows pass). */
    MinSvr = let v = GetCfg("minimum svr")
             in if v = null then 0
                else if v is number then v
                else try Number.FromText(Text.Trim(Text.From(v))) otherwise 0,

    /* -------- 3. Selected decryptor -------- */
    DecName    = let v = GetCfg("decryptor")
                 in if v = null then "" else Text.Lower(Text.Trim(Text.From(v))),
    DecTbl     = Table.Buffer(decryptors),
    DecMatch   = try
                    Table.SelectRows(DecTbl,
                        each Text.Lower(Text.Trim(_[#"Decryptor Name"])) = DecName){0}
                 otherwise null,
    DecProb    = if DecMatch = null then 1
                 else try ToNum(DecMatch[#"Probability Multiplier"]) otherwise 1,
    DecRuns    = if DecMatch = null then 0
                 else try ToNum(DecMatch[#"Max Run Modifier"]) otherwise 0,
    DecME      = if DecMatch = null then 0
                 else try ToNum(DecMatch[#"ME Modifier"]) otherwise 0,
    DecTE      = if DecMatch = null then 0
                 else try ToNum(DecMatch[#"PE Modifier"]) otherwise 0,

    /* -------- 4. Indexed lookups -------- */

    /* Blueprint reference: BP typeId -> record */
    BpRefBuf    = Table.Buffer(blueprintReference),
    BpRefKeys   = List.Transform(BpRefBuf[_key], each Text.From(_)),
    BpRefRows   = Table.ToRecords(BpRefBuf),
    BpRefIndex  = Record.FromList(BpRefRows, BpRefKeys),
    BpRefOf     = (typeId as nullable number) as nullable record =>
        if typeId = null then null
        else Record.FieldOrDefault(BpRefIndex, Text.From(typeId), null),

    /* Reverse index: product typeId -> BP _key (for "is this row's product
       a T2 item that has its OWN BP entry?" detection). */
    BpByProductPairs = List.Transform(BpRefRows, each {Text.From([productTypeId]), [_key]}),
    BpByProductValid = List.Select(BpByProductPairs, each _{0} <> "null" and _{0} <> ""),
    BpByProductIdx   = Record.FromList(
                          List.Transform(BpByProductValid, each _{1}),
                          List.Transform(BpByProductValid, each _{0})),
    BpKeyForProduct  = (productTypeId as nullable number) as nullable number =>
        if productTypeId = null then null
        else Record.FieldOrDefault(BpByProductIdx, Text.From(productTypeId), null),

    /* Types: typeId -> { name (already JSON record), metaGroupID } */
    TypesBuf      = Table.Buffer(Table.SelectColumns(#"types-ref",
                                {"_key","name","metaGroupID"})),
    TypesKeys     = List.Transform(TypesBuf[_key], each Text.From(_)),
    TypesRecs     = Table.ToRecords(TypesBuf),
    TypesIndex    = Record.FromList(TypesRecs, TypesKeys),

    ResolveName = (raw as any) as nullable text =>
        let
            rec = if raw = null then null
                  else if raw is record then raw
                  else if raw is text then try Json.Document(raw) otherwise null
                  else null
        in
            if rec = null then null
            else
                let v = Record.FieldOrDefault(rec, Lang, null)
                in  if v <> null then v else Record.FieldOrDefault(rec, "en", null),

    NameOf = (typeId as nullable number) as nullable text =>
        if typeId = null then null
        else
            let r = Record.FieldOrDefault(TypesIndex, Text.From(typeId), null)
            in  if r = null then null else ResolveName(r[name]),

    MetaGroupOf = (typeId as nullable number) as number =>
        if typeId = null then 0
        else
            let r = Record.FieldOrDefault(TypesIndex, Text.From(typeId), null)
            in  if r = null then 0
                else let mg = r[metaGroupID] in if mg = null then 0 else mg,

    /* Contracts: typeId -> BPC price */
    ContractsBuf   = Table.Buffer(jitaContracts),
    ContractKeys   = List.Transform(ContractsBuf[typeId], each Text.From(_)),
    ContractVals   = ContractsBuf[price],
    ContractIndex  = Record.FromList(ContractVals, ContractKeys),
    BPCPriceOf = (typeId as nullable number) as nullable number =>
        if typeId = null then null
        else Record.FieldOrDefault(ContractIndex, Text.From(typeId), null),

    /* -------- 4b. Synthesize "boughtBpc" rows from contracts --------
       For every BPC listed in jitaContracts that the user doesn't
       already own, add a row at the user-configured bought-BPC
       ME/TE/runs. Dedup on { typeId, ME, TE, runs } so owned variants
       take precedence. A `source` column marks origin. */
    OwnedTagged = Table.AddColumn(Renamed, "source", each "owned", type text),

    /* Set of owned typeIds for fast exclusion.
       Use List.Distinct on keys: the user may own the same typeId at
       multiple ME/TE/runs, which would produce duplicate record keys. */
    OwnedTypeIdKeys = List.Distinct(List.Transform(Renamed[typeId], each Text.From(_))),
    OwnedTypeIdSet  = Record.FromList(List.Repeat({true}, List.Count(OwnedTypeIdKeys)),
                                      OwnedTypeIdKeys),

    /* Distinct contract typeIds not already owned. */
    BoughtBpcKeys  = List.Distinct(List.Select(
                        List.Transform(ContractsBuf[typeId], each Text.From(_)),
                        each not Record.FieldOrDefault(OwnedTypeIdSet, _, false))),

    BoughtBpcRecs  = List.Transform(BoughtBpcKeys, (k) =>
        [ typeId = Number.FromText(k),
          ME     = BoughtBpcME,
          TE     = BoughtBpcTE,
          runs   = BoughtBpcRuns,
          source = "boughtBpc" ]),
    BoughtBpcTbl   = Table.FromRecords(BoughtBpcRecs,
                        type table [typeId = Int64.Type, ME = Int64.Type,
                                    TE = Int64.Type, runs = Int64.Type,
                                    source = text]),

    Combined0      = Table.Combine({OwnedTagged, BoughtBpcTbl}),
    /* Owned rows come first, so Distinct keeps them and drops any
       synthetic duplicate at the same {typeId,ME,TE,runs}. */
    Combined       = Table.Distinct(Combined0, {"typeId","ME","TE","runs"}),

    /* -------- 5. Material / time helpers -------- */

    /* ME-adjusted total qty for N runs of a single material baseQty. */
    MeAdjustQty = (baseQty as number, runs as number, ME as number) as number =>
        if baseQty <= 0 then 0
        else
            let
                perRun = baseQty * (1 - ME/100),
                total  = Number.RoundUp(Number.Round(runs * perRun, 2))
            in
                List.Max({runs, total}),

    /* Σ MeAdjustQty(baseQty, runs, ME) × buyPrice × BuyOrderMult
       over a normalized materials list { typeId, quantity, buyPrice }. */
    MfgMatCostAt = (mats as nullable list, runs as number, ME as number) as nullable number =>
        if mats = null then null
        else
            let
                parts = List.Transform(mats, (r) =>
                    let
                        bq = r[quantity],
                        px = r[buyPrice],
                        adjQ = MeAdjustQty(bq, runs, ME)
                    in
                        if px = null then 0 else adjQ * px),
                s = List.Sum(parts)
            in
                if s = null then 0 else s * BuyOrderMult,

    /* TE-adjusted total time for N runs. */
    TimeAt = (baseSec as nullable number, runs as number,
              TE as number, durMod as number) as nullable number =>
        if baseSec = null then null
        else baseSec * runs * (1 - TE/100) * (1 + durMod),

    /* Job install fee. */
    JobFee = (eiv as nullable number, runs as number, rate as number) as nullable number =>
        if eiv = null then null else eiv * runs * rate,

    /* Build EVE multibuy text from a list of chunks {chunkMats, n, meAdj}.
       Aggregates duplicate typeIds across chunks, ME-adjusts each chunk's
       quantities (meAdj=0 means no adjustment), formats as "Name\tqty" lines.
       EVE multibuy clipboard accepts tab- or space-separated name+qty.
       Field names intentionally avoid collisions with outer-scope params
       (ME, runs) to dodge a Power Query nested-scope recompile bug. */
    BuildMultibuy = (entries as list, extras as nullable list) as nullable text =>
        let
            fromMats = List.Combine(List.Transform(entries, (e) =>
                let m = e[chunkMats] in
                if m = null then {}
                else List.Transform(m, (r) =>
                    [ typeId = r[typeId],
                      qty    = MeAdjustQty(r[quantity], e[n], e[meAdj]) ]))),
            grouped =
                if List.Count(fromMats) = 0 then {}
                else
                    let
                        tbl = Table.FromRecords(fromMats),
                        g   = Table.Group(tbl, {"typeId"},
                                {{"qty", each List.Sum([qty]), type number}})
                    in
                        Table.ToRecords(g),
            mainLines = List.Transform(grouped, (g) =>
                let n = NameOf(g[typeId]) in
                if n = null or g[qty] <= 0 then null
                else n & "#(tab)" & Text.From(Number.RoundUp(g[qty]))),
            extraLines =
                if extras = null then {}
                else List.Transform(extras, (x) =>
                    let n = x[name], q = x[qty] in
                    if n = null or n = "" or q = null or q <= 0 then null
                    else n & "#(tab)" & Text.From(Number.RoundUp(q))),
            kept = List.Select(mainLines & extraLines, each _ <> null)
        in
            if List.Count(kept) = 0 then null
            else Text.Combine(kept, "#(lf)"),

    /* -------- 6. Per-row stage calculations --------
       Each "Best" function returns a record:
         { iskPerHour, runsPerHour|jobsPerHour|attemptsPerHour,
           multibuy, path? }
       runsPerHour/etc. is rounded UP so the multibuy material count
       is always sufficient to fill an hour.
    */

    /* Manufacturing.
       BPC row (runs > 0):
         - purchased=false: direct mfg, no acquisition cost ("ownBpc").
         - purchased=true : must buy the BPC; charge bpcMarketPrice/runs
                            per mfg run. If no contract price exists,
                            the row is unbuildable -> null. ("buyBpc")
       BPO row (runs = -1): pick best of pathBuy (buy BPC from contracts;
         no extra time, BPC not multibuyable so omitted from list) and
         pathCopy (self-copy then mfg; amortize copy time/cost over cpyOut
         runs; multibuy includes copy materials × copies needed). */
    MfgBest = (bp as nullable record, ME as number, TE as number,
               runs as number, purchased as logical) as nullable record =>
        if bp = null then null
        else
            let
                mats   = bp[mfgBaseMaterials],
                tBase  = bp[mfgBaseTimeSec],
                gross  = bp[mfgProductGross],
                eiv    = bp[mfgEIV]
            in
                if mats = null or tBase = null or gross = null then null
                else
                    let
                        matCost = MfgMatCostAt(mats, 1, ME),
                        feeCol  = JobFee(eiv, 1, MfgRate),
                        fee     = if feeCol is number then feeCol else 0,
                        mfgCost = matCost + fee,
                        mfgTime = TimeAt(tBase, 1, TE, MfgDurMod),
                        rev     = gross * SellMult
                    in
                        if mfgTime = null or mfgTime <= 0 then null
                        else if runs > 0 then
                            let
                                /* Acquisition cost: bpcMarketPrice is already ISK/run
                                   (the scraper stores pricePerRun = contractPrice/runs).
                                   Opportunity cost for owned BPCs: forgoing the sale = bpcPx;
                                   if not on contracts it can't be sold either, so cost = 0.
                                   Hard requirement for purchased BPCs: null = unbuildable. */
                                bpcPx     = bp[bpcMarketPrice],
                                acqPerRun = if not purchased
                                            then if bpcPx = null then 0 else bpcPx
                                            else bpcPx   // null when bpcPx = null
                            in
                                if acqPerRun = null then null
                                else
                                    let
                                        iph = (rev - mfgCost - acqPerRun) / (mfgTime / 3600),
                                        rph = Number.RoundUp(3600 / mfgTime),
                                        mb  = BuildMultibuy(
                                                {[ chunkMats = mats, n = rph, meAdj = ME ]},
                                                null)
                                    in
                                        [ iskPerHour  = iph,
                                          runsPerHour = rph,
                                          multibuy    = mb,
                                          bpcCost     = acqPerRun,
                                          path        = if purchased then "buyBpc" else "ownBpc" ]
                        else
                            let
                                /* pathBuy: buy BPC from contracts (BPC not in multibuy).
                                   bpcMarketPrice is already ISK/run; no further division needed. */
                                bpcPx     = bp[bpcMarketPrice],
                                buyValid  = bpcPx <> null,
                                buyCostPR = bpcPx,
                                buyIPH    = if not buyValid then null
                                            else (rev - mfgCost - buyCostPR) / (mfgTime / 3600),
                                buyPick   = if buyIPH = null then null
                                            else
                                                let
                                                    rph = Number.RoundUp(3600 / mfgTime),
                                                    mb  = BuildMultibuy(
                                                            {[ chunkMats = mats, n = rph, meAdj = ME ]},
                                                            null)
                                                in
                                                    [ iskPerHour  = buyIPH,
                                                      runsPerHour = rph,
                                                      multibuy    = mb,
                                                      bpcCost     = buyCostPR,
                                                      path        = "buyBpc" ],

                                /* pathCopy: amortize copy job per mfg run. */
                                cpyTimeBase = bp[copyBaseTimeSec],
                                cpyOut      = bp[copyOutputRuns],
                                cpyEIV      = bp[copyEIV],
                                cpyMatCost  = bp[copyMaterialsCostBuy],
                                cpyMats     = bp[copyBaseMaterials],
                                cpyValid    = cpyTimeBase <> null and cpyOut <> null and cpyOut > 0,
                                cpyMatN     = if cpyMatCost is number then cpyMatCost else 0,
                                cpyFeeCol   = JobFee(cpyEIV, cpyOut, CpyRate),
                                cpyFee      = if cpyFeeCol is number then cpyFeeCol else 0,
                                cpyCostPR   = if cpyValid then (cpyMatN + cpyFee) / cpyOut else null,
                                cpyTimePerRun = if cpyValid
                                                then (cpyTimeBase * (1 + CpyDurMod)) / cpyOut
                                                else null,
                                cycleTime   = if cpyValid then mfgTime + cpyTimePerRun else null,
                                copyIPH     = if not cpyValid or cycleTime = null or cycleTime <= 0
                                              then null
                                              else (rev - mfgCost - cpyCostPR) / (cycleTime / 3600),
                                copyPick    = if copyIPH = null then null
                                              else
                                                  let
                                                      rph    = Number.RoundUp(3600 / cycleTime),
                                                      copies = Number.RoundUp(rph / cpyOut),
                                                      mb     = BuildMultibuy(
                                                                {[ chunkMats = mats,    n = rph,    meAdj = ME ],
                                                                 [ chunkMats = cpyMats, n = copies, meAdj = 0  ]},
                                                                null)
                                                  in
                                                      [ iskPerHour  = copyIPH,
                                                        runsPerHour = rph,
                                                        multibuy    = mb,
                                                        bpcCost     = cpyCostPR,
                                                        path        = "selfCopy" ],

                                cands = List.Select({buyPick, copyPick}, each _ <> null),
                                best  = if List.Count(cands) = 0 then null
                                        else List.First(List.Sort(cands,
                                                (a,b) => Value.Compare(b[iskPerHour], a[iskPerHour])))
                            in
                                best,

    /* Copy: returns iskPerHour + jobsPerHour + multibuy. Copy materials are
       usually empty in EVE but include them when present. */
    CopyBest = (bp as nullable record) as nullable record =>
        if bp = null then null
        else
            let
                tBase  = bp[copyBaseTimeSec],
                outRuns= bp[copyOutputRuns],
                bpcPx  = bp[bpcMarketPrice],
                eiv    = bp[copyEIV],
                matsC  = bp[copyMaterialsCostBuy],
                cpyMats= bp[copyBaseMaterials]
            in
                if tBase = null or outRuns = null or outRuns <= 0 or bpcPx = null then null
                else
                    let
                        matN   = if matsC is number then matsC else 0,
                        feeCol = JobFee(eiv, outRuns, CpyRate),
                        fee    = if feeCol is number then feeCol else 0,
                        cost   = matN + fee,
                        time   = tBase * outRuns * (1 + CpyDurMod),
                        /* bpcMarketPrice is per-run; one copy job produces outRuns BPC runs. */
                        rev    = bpcPx * outRuns * SellMult,
                        profit = rev - cost
                    in
                        if time <= 0 then null
                        else
                            let
                                iph = profit / (time / 3600),
                                jph = Number.RoundUp(3600 / time),
                                mb  = BuildMultibuy(
                                        {[ chunkMats = cpyMats, n = jph, meAdj = 0 ]},
                                        null)
                            in
                                [ iskPerHour = iph, jobsPerHour = jph, multibuy = mb ],

    /* Cost per single run on a freshly-self-copied BPC.
       Used as the "T1 BPC consumed" cost in invention. */
    SelfCopyCostPerRun = (bp as nullable record) as nullable number =>
        if bp = null then null
        else
            let
                outRuns= bp[copyOutputRuns],
                eiv    = bp[copyEIV],
                matsC  = bp[copyMaterialsCostBuy]
            in
                if outRuns = null or outRuns <= 0 then null
                else
                    let
                        matN = if matsC is number then matsC else 0,
                        fee  = let f = JobFee(eiv, outRuns, CpyRate) in if f = null then 0 else f
                    in
                        (matN + fee) / outRuns,

    /* Shared per-attempt cost/time (independent of which product is
       chosen). Players pick a target T2 product when starting invention,
       so each attempt has ONE specific product (not a portfolio).
       Cost per attempt: datacores + invention job fee + 1 freshly-copied
       T1 BPC consumed (whether or not the attempt succeeds).
       Time per attempt: (copy time / copyOutputRuns) + invention time.
       The T1 BPC consumed each attempt must be copied first, so its copy
       time is amortized into the attempt time. Formula:
         attemptTime = copyBaseTimeSec × (1 + CpyDurMod) / copyOutputRuns
                     + inventionBaseTimeSec × (1 + InvDurMod)
       bpcCost = amortized T1 copy cost per attempt (= SelfCopyCostPerRun). */
    InventionShared = (bp as nullable record) as nullable record =>
        if bp = null then null
        else
            let
                tBase  = bp[inventionBaseTimeSec],
                matsC  = bp[inventionMaterialsCostBuy],
                eiv    = bp[inventionEIV],
                bpcPR  = SelfCopyCostPerRun(bp)
            in
                if tBase = null then null
                else
                    let
                        attMat  = if matsC is number then matsC else 0,
                        attFee  = let f = JobFee(eiv, 1, InvRate) in if f = null then 0 else f,
                        attBpc  = if bpcPR is number then bpcPR else 0,
                        /* Copy time amortized per attempt (= per T1 BPC consumed). */
                        cpyTimeBase   = bp[copyBaseTimeSec],
                        cpyOut        = bp[copyOutputRuns],
                        cpyTimePerRun = if cpyTimeBase <> null and cpyOut <> null and cpyOut > 0
                                        then (cpyTimeBase * (1 + CpyDurMod)) / cpyOut
                                        else 0,
                        invTime       = tBase * (1 + InvDurMod)
                    in
                        [ attemptCost = attMat + attFee + attBpc,
                          attemptTime = invTime + cpyTimePerRun,
                          bpcCost     = attBpc ],

    /* For one specific T2 product:
         expected T2 BPC runs per attempt =
             quantity × (baseProb × decProb) × (baseRuns + decRuns)
         expected gross (selling resulting T2 BPCs at contract price) =
             expRuns × bpcMarketPrice(T2 BP)
       Returns null if no contract price for that T2 BPC (then
       "sell BPCs" path is unrealistic for this product). */
    InventionForProduct = (sh as nullable record, p as nullable record) as nullable record =>
        if sh = null or p = null then null
        else
            let
                expRuns = p[quantity] * (p[baseProbability] * DecProb) * (p[baseRuns] + DecRuns),
                bp2px   = BPCPriceOf(p[typeId]),
                gross   = if bp2px = null then null else expRuns * bp2px,
                time    = sh[attemptTime]
            in
                if gross = null or time = null or time <= 0 then null
                else
                    let
                        rev    = gross * SellMult,
                        profit = rev - sh[attemptCost]
                    in
                        [ productId    = p[typeId],
                          expectedRuns = expRuns,
                          grossRevenue = gross,
                          iskPerHour   = profit / (time / 3600) ],

    /* Pick the highest-ISK/h product for selling-BPCs-on-contract.
       Returns null when no product has a contract price.
       Falls back to the first product (with id only, no iph) so callers
       that need a target product (e.g. PathInvent below) still have a
       choice when no contract data exists. */
    BestInventionSell = (bp as nullable record) as nullable record =>
        if bp = null then null
        else
            let
                sh    = InventionShared(bp),
                prods = bp[inventionProductsRaw]
            in
                if sh = null or prods = null or List.Count(prods) = 0 then null
                else
                    let
                        cands = List.Select(
                                    List.Transform(prods, (p) => InventionForProduct(sh, p)),
                                    each _ <> null),
                        best  = if List.Count(cands) = 0 then null
                                else List.First(List.Sort(cands,
                                        (a,b) => Value.Compare(b[iskPerHour], a[iskPerHour])))
                    in
                        best,

    /* Backwards-compatible iph (best product). */
    InventionPerHour = (bp as nullable record) as nullable number =>
        let b = BestInventionSell(bp) in
        if b = null then null else b[iskPerHour],

    /* Invention with multibuy. Returns ISK/h for the best-product
       sell-BPCs path plus the chosen productId. Multibuy lists invention
       base materials (datacores) and the selected decryptor (1/attempt). */
    InventionBest = (bp as nullable record) as nullable record =>
        if bp = null then null
        else
            let
                sh = InventionShared(bp),
                b  = BestInventionSell(bp)
            in
                if sh = null or b = null then null
                else
                    let
                        time   = sh[attemptTime],
                        iph    = b[iskPerHour],
                        aph    = Number.RoundUp(3600 / time),
                        invMats= bp[inventionBaseMaterials],
                        decRow = if DecMatch = null then null
                                 else [ name = DecMatch[#"Decryptor Name"], qty = aph ],
                        extras = if decRow = null then null else {decRow},
                        mb     = BuildMultibuy(
                                    {[ chunkMats = invMats, n = aph, meAdj = 0 ]},
                                    extras)
                    in
                        [ iskPerHour      = iph,
                          attemptsPerHour = aph,
                          multibuy        = mb,
                          productId       = b[productId],
                          bpcCost         = sh[bpcCost] ],

    /* -------- 7. T2 build paths --------
       For each path compute (profit, time) for the same number of T2
       runs N, then ISK/h = profit / (time/3600). */

    /* Downstream T2 mfg numbers for N runs at given ME/TE on the T2 BP. */
    T2BuildBlock = (t2bp as nullable record, N as number,
                    t2ME as number, t2TE as number) as nullable record =>
        if t2bp = null or N <= 0 then null
        else
            let
                mats   = t2bp[mfgBaseMaterials],
                tBase  = t2bp[mfgBaseTimeSec],
                grossPR= t2bp[mfgProductGross],
                eiv    = t2bp[mfgEIV]
            in
                if mats = null or tBase = null or grossPR = null then null
                else
                    let
                        matCost = MfgMatCostAt(mats, N, t2ME),
                        feeCol  = JobFee(eiv, N, MfgRate),
                        fee     = if feeCol is number then feeCol else 0,
                        time    = TimeAt(tBase, N, t2TE, MfgDurMod),
                        rev     = grossPR * N * SellMult
                    in
                        [ cost = matCost + fee, time = time, revenue = rev ],

    /* (a) Invent ourselves: one attempt -> expected runs of one chosen
       T2 product, then build that T2. Iterate over every possible product,
       compute lifecycle ISK/h (sell finished T2 items, NOT BPCs), pick the
       max. Resulting T2 BP is invented at decryptor-modified ME/TE
       (default 2/4). */
    PathInventForProduct = (sh as nullable record, p as nullable record) as nullable record =>
        if sh = null or p = null then null
        else
            let
                expRuns = p[quantity] * (p[baseProbability] * DecProb) * (p[baseRuns] + DecRuns),
                t2bp    = BpRefOf(p[typeId]),
                block   = if expRuns <= 0 then null
                          else T2BuildBlock(t2bp, expRuns, 2 + DecME, 4 + DecTE)
            in
                if block = null then null
                else
                    let
                        cost   = sh[attemptCost] + block[cost],
                        time   = sh[attemptTime] + block[time],
                        profit = block[revenue] - cost
                    in
                        if time = null or time <= 0 then null
                        else [ productId  = p[typeId],
                               iskPerHour = profit / (time / 3600) ],

    PathInvent = (t1bp as nullable record) as nullable record =>
        if t1bp = null then null
        else
            let
                sh    = InventionShared(t1bp),
                prods = t1bp[inventionProductsRaw]
            in
                if sh = null or prods = null or List.Count(prods) = 0 then null
                else
                    let
                        cands = List.Select(
                                    List.Transform(prods, (p) => PathInventForProduct(sh, p)),
                                    each _ <> null),
                        best  = if List.Count(cands) = 0 then null
                                else List.First(List.Sort(cands,
                                        (a,b) => Value.Compare(b[iskPerHour], a[iskPerHour])))
                    in
                        best,

    /* (b) Buy BPC from contracts: skip copy/invent; pay BPC price + mfg.
       Iterates over every potential T2 product, picks the one with the
       highest ISK/h. Skips products with no contract price. */
    PathBuyBpcForProduct = (p as nullable record) as nullable record =>
        if p = null then null
        else
            let
                t2BpId = p[typeId],
                bpcPx  = BPCPriceOf(t2BpId),
                t2bp   = BpRefOf(t2BpId),
                N      = BoughtBpcRuns,
                block  = if N <= 0 then null
                         else T2BuildBlock(t2bp, N, BoughtBpcME, BoughtBpcTE)
            in
                if bpcPx = null or block = null then null
                else
                    let
                        cost   = bpcPx + block[cost],
                        time   = block[time],
                        profit = block[revenue] - cost
                    in
                        if time = null or time <= 0 then null
                        else [ productId  = t2BpId,
                               iskPerHour = profit / (time / 3600) ],

    PathBuyBpc = (t1bp as nullable record) as nullable record =>
        if t1bp = null then null
        else
            let
                prods = t1bp[inventionProductsRaw]
            in
                if prods = null or List.Count(prods) = 0 then null
                else
                    let
                        cands = List.Select(
                                    List.Transform(prods, (p) => PathBuyBpcForProduct(p)),
                                    each _ <> null),
                        best  = if List.Count(cands) = 0 then null
                                else List.First(List.Sort(cands,
                                        (a,b) => Value.Compare(b[iskPerHour], a[iskPerHour])))
                    in
                        best,

    /* (c) Row IS itself a T2 BPO: only valid when the row's BP produces
           a meta-group-2 item AND that BP itself has no invention activity
           (true T2 BPOs are rare but exist for some structures/ships). */
    PathOwnBpo = (rowBp as nullable record, ME as number, TE as number) as nullable record =>
        if rowBp = null then null
        else
            let
                prodId  = rowBp[productTypeId],
                metaGrp = MetaGroupOf(prodId),
                hasInv  = rowBp[inventionProductsRaw] <> null,
                isT2Bpo = (metaGrp = 2) and (not hasInv),
                block   = if isT2Bpo then T2BuildBlock(rowBp, 1, ME, TE) else null,
                profit  = if block = null then null else block[revenue] - block[cost],
                time    = if block = null then null else block[time]
            in
                if block = null or time = null or time <= 0 then null
                else [ iskPerHour = profit / (time / 3600) ],

    /* -------- 8. Reaction ISK/h -------- */
    ReactionBest = (bp as nullable record) as nullable record =>
        if bp = null then null
        else
            let
                mats  = bp[reactionBaseMaterials],
                tBase = bp[reactionBaseTimeSec],
                gross = bp[reactionProductGross],
                eiv   = bp[reactionEIV]
            in
                if mats = null or tBase = null or gross = null then null
                else
                    let
                        matCost = MfgMatCostAt(mats, 1, 0),
                        feeCol  = JobFee(eiv, 1, RxnRate),
                        fee     = if feeCol is number then feeCol else 0,
                        cost    = matCost + fee,
                        time    = tBase * (1 + RxnDurMod),
                        rev     = gross * SellMult
                    in
                        if time <= 0 then null
                        else
                            let
                                iph = (rev - cost) / (time / 3600),
                                rph = Number.RoundUp(3600 / time),
                                mb  = BuildMultibuy(
                                        {[ chunkMats = mats, n = rph, meAdj = 0 ]},
                                        null)
                            in
                                [ iskPerHour = iph, runsPerHour = rph, multibuy = mb ],

    /* -------- 9. Build per-row output -------- */
    Computed = Table.AddColumn(Combined, "_calc", each
        let
            tid    = [typeId],
            me     = [ME],
            te     = [TE],
            rns    = [runs],
            src    = [source],
            isBpo  = rns = -1,
            /* Purchased = the BPC was acquired from contracts (and so its
               cost must be amortized into per-run mfg profit). True for
               every synthetic invention-only row; false for user-owned BPs
               (BPOs and owned BPCs alike). */
            purchased = src = "boughtBpc",
            bp     = BpRefOf(tid),
            prodId = if bp = null then null else bp[productTypeId],

            /* SVR check for manufacturing.
               SVR = market daily sellVolume / my daily output
                   = sellVolume / (productQtyPerRun x 86400 / mfgTimeSec)
               Pass when MinSvr = 0 or SVR >= MinSvr.
               If market data missing, SVR is unknown -> pass (null). */
            mfgTimeSec   = if bp = null then null
                           else TimeAt(bp[mfgBaseTimeSec], 1, te, MfgDurMod),
            mfgProdQty   = if bp = null then 1
                           else let q = bp[productQuantityPerRun]
                                in  if q = null or q <= 0 then 1 else q,
            mfgSellVol   = if bp = null then null else bp[mfgProductSellVolume],
            mfgMyPerDay  = if mfgTimeSec = null or mfgTimeSec <= 0 then null
                           else mfgProdQty * (86400 / mfgTimeSec),
            mfgSvr       = if mfgMyPerDay = null or mfgMyPerDay <= 0
                               or mfgSellVol = null
                           then null
                           else mfgSellVol / mfgMyPerDay,
            mfgSvrOk     = MinSvr <= 0 or mfgSvr = null or mfgSvr >= MinSvr,

            /* SVR check for reactions (same logic, different sell volume). */
            rxnTimeSec   = if bp = null then null
                           else let t = bp[reactionBaseTimeSec]
                                in  if t = null then null else t * (1 + RxnDurMod),
            rxnSellVol   = if bp = null then null else bp[reactionProductSellVolume],
            rxnMyPerDay  = if rxnTimeSec = null or rxnTimeSec <= 0 then null
                           else 86400 / rxnTimeSec,
            rxnSvr       = if rxnMyPerDay = null or rxnMyPerDay <= 0
                               or rxnSellVol = null
                           then null
                           else rxnSellVol / rxnMyPerDay,
            rxnSvrOk     = MinSvr <= 0 or rxnSvr = null or rxnSvr >= MinSvr,

            mfg    = if not mfgSvrOk then null else MfgBest(bp, me, te, rns, purchased),
            /* Copy and invention require a BPO (runs = -1). BPCs can't be
               copied or used as invention input. */
            cpy    = if isBpo then CopyBest(bp) else null,
            inv    = if isBpo then InventionBest(bp) else null,
            rxn    = if not rxnSvrOk then null else ReactionBest(bp),

            mfgPH       = if mfg = null then null else mfg[iskPerHour],
            mfgRPH      = if mfg = null then null else mfg[runsPerHour],
            mfgMB       = if mfg = null then null else mfg[multibuy],
            mfgPath     = if mfg = null then null else mfg[path],
            mfgBpcCost  = if mfg = null then null else mfg[bpcCost],

            cpyPH       = if cpy = null then null else cpy[iskPerHour],
            cpyJPH      = if cpy = null then null else cpy[jobsPerHour],
            cpyMB       = if cpy = null then null else cpy[multibuy],

            invPH       = if inv = null then null else inv[iskPerHour],
            invAPH      = if inv = null then null else inv[attemptsPerHour],
            invMB       = if inv = null then null else inv[multibuy],
            invProdId   = if inv = null then null else inv[productId],
            invProdName = NameOf(invProdId),
            invBpcCost  = if inv = null then null else inv[bpcCost],

            rxnPH       = if rxn = null then null else rxn[iskPerHour],
            rxnRPH      = if rxn = null then null else rxn[runsPerHour],
            rxnMB       = if rxn = null then null else rxn[multibuy],

            invents = isBpo and bp <> null and bp[inventionProductsRaw] <> null,
            invPath = if invents then PathInvent(bp) else null,
            buyPath = if bp <> null and bp[inventionProductsRaw] <> null
                      then PathBuyBpc(bp) else null,
            ownPath = PathOwnBpo(bp, me, te),

            invIPH    = if invPath = null then null else invPath[iskPerHour],
            invPathId = if invPath = null then null else invPath[productId],
            buyIPH    = if buyPath = null then null else buyPath[iskPerHour],
            buyPathId = if buyPath = null then null else buyPath[productId],
            ownIPH    = if ownPath = null then null else ownPath[iskPerHour],

            cands   = List.Select(
                        {{"invent", invIPH, invPathId},
                         {"buyBpc", buyIPH, buyPathId},
                         {"ownBpo", ownIPH, prodId}},
                        each _{1} <> null),
            best    = if List.Count(cands) = 0 then null
                      else List.First(List.Sort(cands, (a,b) => Value.Compare(b{1}, a{1}))),
            bestLbl = if best = null then null else best{0},
            bestVal = if best = null then null else best{1},
            bestPid = if best = null then null else best{2}
        in
            [
                productTypeId           = prodId,
                productName             = NameOf(prodId),
                name                    = NameOf(tid),
                mfgProfitPerHour        = mfgPH,
                mfgRunsPerHour          = mfgRPH,
                mfgMultibuy             = mfgMB,
                mfgPath                 = mfgPath,
                mfgBpcCost              = mfgBpcCost,
                copyProfitPerHour       = cpyPH,
                copyJobsPerHour         = cpyJPH,
                copyMultibuy            = cpyMB,
                inventionProfitPerHour  = invPH,
                inventionAttemptsPerHour= invAPH,
                inventionMultibuy       = invMB,
                inventionBestProductId  = invProdId,
                inventionBestProductName= invProdName,
                inventionBpcCost        = invBpcCost,
                reactionProfitPerHour   = rxnPH,
                reactionRunsPerHour     = rxnRPH,
                reactionMultibuy        = rxnMB,
                t2BuildPath_invent_iskPerHour = invIPH,
                t2BuildPath_invent_productId  = invPathId,
                t2BuildPath_buyBpc_iskPerHour = buyIPH,
                t2BuildPath_buyBpc_productId  = buyPathId,
                t2BuildPath_ownBpo_iskPerHour = ownIPH,
                bestT2Path              = bestLbl,
                bestT2IskPerHour        = bestVal,
                bestT2ProductId         = bestPid,
                bestT2ProductName       = NameOf(bestPid)
            ]),

    Expanded = Table.ExpandRecordColumn(Computed, "_calc",
        {"productTypeId","productName","name",
         "mfgProfitPerHour","mfgRunsPerHour","mfgMultibuy","mfgPath","mfgBpcCost",
         "copyProfitPerHour","copyJobsPerHour","copyMultibuy",
         "inventionProfitPerHour","inventionAttemptsPerHour","inventionMultibuy",
         "inventionBestProductId","inventionBestProductName","inventionBpcCost",
         "reactionProfitPerHour","reactionRunsPerHour","reactionMultibuy",
         "t2BuildPath_invent_iskPerHour","t2BuildPath_invent_productId",
         "t2BuildPath_buyBpc_iskPerHour","t2BuildPath_buyBpc_productId",
         "t2BuildPath_ownBpo_iskPerHour",
         "bestT2Path","bestT2IskPerHour","bestT2ProductId","bestT2ProductName"}),

    Final = Table.SelectColumns(Expanded,
        {"typeId","name","ME","TE","runs","source",
         "productTypeId","productName",
         "mfgProfitPerHour","mfgRunsPerHour","mfgMultibuy","mfgPath","mfgBpcCost",
         "copyProfitPerHour","copyJobsPerHour","copyMultibuy",
         "inventionProfitPerHour","inventionAttemptsPerHour","inventionMultibuy",
         "inventionBestProductId","inventionBestProductName","inventionBpcCost",
         "reactionProfitPerHour","reactionRunsPerHour","reactionMultibuy",
         "t2BuildPath_invent_iskPerHour","t2BuildPath_invent_productId",
         "t2BuildPath_buyBpc_iskPerHour","t2BuildPath_buyBpc_productId",
         "t2BuildPath_ownBpo_iskPerHour",
         "bestT2Path","bestT2IskPerHour","bestT2ProductId","bestT2ProductName"})
in
    Final
