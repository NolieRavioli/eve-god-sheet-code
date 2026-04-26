let
    /* ============================================================
       blueprintReference.pq

       Exposes RAW per-BP building blocks. No ME, TE, decryptor, or
       config-driven economics applied here — those belong in
       ownedBPs.pq where ME/TE per row are known.

       Per row (_key = blueprint typeID), columns emitted:
         productTypeId, productQuantityPerRun
         mfgBaseMaterials      list of { typeId, quantity, buyPrice }
         mfgBaseTimeSec
         mfgEIV                Σ qty × adjustedPrice (CCP daily)
         mfgMaterialsCostBuy   Σ qty × buyPrice × BuyOrderMult, at ME 0
         mfgProductGross       productQty × sellMethodPrice (no SellMult)
         bpcMarketPrice        contract price of this BP's BPC, if listed
         copyBaseTimeSec, copyOutputRuns,
         copyMaterialsCostBuy, copyEIV
         inventionBaseTimeSec,
         inventionMaterialsCostBuy, inventionEIV,
         inventionProductsRaw  list of { typeId, quantity, baseProbability, baseRuns }
         reactionBaseMaterials, reactionBaseTimeSec,
         reactionMaterialsCostBuy, reactionEIV,
         reactionProductGross
       ============================================================ */

    /* -------- 1. Source -------- */
    Source = Csv.Document(
                File.Contents("C:\Users\nolan\eve-god-sheet-code\sde-csv\blueprints.csv"),
                [Delimiter=",", Columns=4, Encoding=1252, QuoteStyle=QuoteStyle.None]),
    Promoted   = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed      = Table.TransformColumnTypes(Promoted,
                    {{"_key", Int64.Type},
                     {"activities", type text},
                     {"blueprintTypeID", Int64.Type},
                     {"maxProductionLimit", Int64.Type}}),
    ParsedJSON = Table.TransformColumns(Typed, {{"activities", Json.Document}}),

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

    BuyMethod  = GetCfg("buy method"),
    SellMethod = GetCfg("sell method"),
    BrokerPct  = ToPct(GetCfg("broker %")),

    /* Buy-side broker only applies when our buy method represents
       placing buy orders ourselves (buyMax / buyWeightedAverage / etc).
       Instant-buying from sellMin / sellWeightedAverage pays no broker. */
    BuyIsOrder    = if BuyMethod = null then false
                    else Text.StartsWith(Text.Lower(Text.From(BuyMethod)), "buy"),
    BuyOrderMult  = if BuyIsOrder then 1 / (1 - BrokerPct) else 1,

    /* -------- 3. Indexed lookups -------- */

    /* Market: typeId -> record-of-prices */
    MarketTbl   = Table.Buffer(jitaMarket),
    MarketKeys  = List.Transform(MarketTbl[typeId], each Text.From(_)),
    MarketRows  = Table.ToRecords(MarketTbl),
    MarketIndex = Record.FromList(MarketRows, MarketKeys),

    PriceOf = (typeId as nullable number, col as nullable text) as nullable number =>
        if typeId = null or col = null then null
        else
            let r = Record.FieldOrDefault(MarketIndex, Text.From(typeId), null)
            in  if r = null then null else Record.FieldOrDefault(r, col, null),

    /* Adjusted prices (EIV source) from ESI.
       Wrapped in try so that a missing or erroring adjustedPrices query
       degrades gracefully: all EIV fields become null, job fees become 0. */
    AdjTbl    = Table.Buffer(
                    try adjustedPrices
                    otherwise #table(
                        type table [typeId = Int64.Type, adjustedPrice = number, averagePrice = number],
                        {})),
    AdjKeys   = List.Transform(AdjTbl[typeId], each Text.From(_)),
    AdjVals   = AdjTbl[adjustedPrice],
    AdjIndex  = Record.FromList(AdjVals, AdjKeys),
    AdjustedPriceOf = (typeId as nullable number) as nullable number =>
        if typeId = null then null
        else Record.FieldOrDefault(AdjIndex, Text.From(typeId), null),

    /* Contracts: typeId -> BPC price */
    ContractsTbl  = Table.Buffer(jitaContracts),
    ContractKeys  = List.Transform(ContractsTbl[typeId], each Text.From(_)),
    ContractVals  = ContractsTbl[price],
    ContractIndex = Record.FromList(ContractVals, ContractKeys),
    BPCPriceOf = (typeId as nullable number) as nullable number =>
        if typeId = null then null
        else Record.FieldOrDefault(ContractIndex, Text.From(typeId), null),

    /* -------- 4. Helpers -------- */

    /* Normalize a materials list to { typeId, quantity, buyPrice }. */
    NormMaterials = (lst as nullable list) as nullable list =>
        if lst = null then null
        else
            List.Transform(lst, (m) =>
                let
                    tid = try Record.FieldOrDefault(m, "typeID",   0) otherwise 0,
                    q   = try Record.FieldOrDefault(m, "quantity", 0) otherwise 0,
                    qN  = if q is number then q else 0,
                    p   = PriceOf(tid, BuyMethod)
                in
                    [ typeId = tid, quantity = qN, buyPrice = p ]),

    /* Σ qty × buyPrice × BuyOrderMult for a normalized materials list. */
    MaterialsBuyCost = (mats as nullable list) as nullable number =>
        if mats = null then null
        else
            let
                parts = List.Transform(mats, (r) =>
                    let p = r[buyPrice] in
                    if p = null then 0 else r[quantity] * p),
                s = List.Sum(parts)
            in
                if s = null then 0 else s * BuyOrderMult,

    /* EIV = Σ qty × adjustedPrice. Returns null if ANY material lacks
       an adjustedPrice (so downstream knows EIV is incomplete). */
    EIV = (mats as nullable list) as nullable number =>
        if mats = null then null
        else
            let
                parts = List.Transform(mats, (r) =>
                    let ap = AdjustedPriceOf(r[typeId]) in
                    if ap = null then null else r[quantity] * ap),
                hasNull = List.AnyTrue(List.Transform(parts, each _ = null))
            in
                if hasNull then null else List.Sum(parts),

    /* First product of an activity, or null. */
    FirstProduct = (act as nullable record) as nullable record =>
        if act = null then null
        else
            let prods = Record.FieldOrDefault(act, "products", null)
            in  if prods = null or List.Count(prods) = 0 then null else prods{0},

    /* Single-product gross revenue at sell method (no SellMult applied). */
    ProductGross = (act as nullable record) as nullable number =>
        let p = FirstProduct(act) in
        if p = null then null
        else
            let
                tid = try p[typeID]   otherwise null,
                q   = try p[quantity] otherwise 1,
                qN  = if q is number then q else 1,
                px  = PriceOf(tid, SellMethod)
            in
                if px = null then null else qN * px,

    /* Daily sell volume for the first product of an activity. Used by
       industryBrowser to compute SVR (market volume / my output / day). */
    ProductSellVolume = (act as nullable record) as nullable number =>
        let p = FirstProduct(act) in
        if p = null then null
        else
            let tid = try p[typeID] otherwise null
            in  PriceOf(tid, "sellVolume"),

    /* Normalize invention products to raw SDE values
       (no decryptor, no skill, no contract price applied here). */
    NormInvProducts = (act as nullable record) as nullable list =>
        if act = null then null
        else
            let lst = Record.FieldOrDefault(act, "products", null) in
            if lst = null then null
            else
                List.Transform(lst, (m) =>
                    let
                        tid  = try Record.FieldOrDefault(m, "typeID",      0) otherwise 0,
                        q    = try Record.FieldOrDefault(m, "quantity",    1) otherwise 1,
                        prob = try Record.FieldOrDefault(m, "probability", 1) otherwise 1,
                        runs = try Record.FieldOrDefault(m, "runs",        1) otherwise 1
                    in
                        [
                            typeId          = tid,
                            quantity        = if q    is number then q    else 1,
                            baseProbability = if prob is number then prob else 1,
                            baseRuns        = if runs is number then runs else 1
                        ]),

    /* -------- 5. Manufacturing -------- */
    WithMfgMats = Table.AddColumn(ParsedJSON, "mfgBaseMaterials",
        each NormMaterials(try [activities][manufacturing][materials] otherwise null),
        type nullable list),

    WithMfgTime = Table.AddColumn(WithMfgMats, "mfgBaseTimeSec",
        each try [activities][manufacturing][time] otherwise null, Int64.Type),

    WithMfgProd = Table.AddColumn(WithMfgTime, "_mfgProduct",
        each FirstProduct(try [activities][manufacturing] otherwise null)),

    WithProdId = Table.AddColumn(WithMfgProd, "productTypeId",
        each let p = [_mfgProduct] in
             if p = null then null else try p[typeID] otherwise null,
        Int64.Type),

    WithProdQty = Table.AddColumn(WithProdId, "productQuantityPerRun",
        each let p = [_mfgProduct] in
             if p = null then null
             else let q = try p[quantity] otherwise 1 in
                  if q is number then q else 1,
        Int64.Type),

    WithMfgEIV = Table.AddColumn(WithProdQty, "mfgEIV",
        each EIV([mfgBaseMaterials]), type nullable number),

    WithMfgCost = Table.AddColumn(WithMfgEIV, "mfgMaterialsCostBuy",
        each MaterialsBuyCost([mfgBaseMaterials]), type nullable number),

    WithMfgGross = Table.AddColumn(WithMfgCost, "mfgProductGross",
        each ProductGross(try [activities][manufacturing] otherwise null),
        type nullable number),

    WithMfgSellVol = Table.AddColumn(WithMfgGross, "mfgProductSellVolume",
        each ProductSellVolume(try [activities][manufacturing] otherwise null),
        type nullable number),

    /* -------- 6. BPC market price -------- */
    WithBpcPx = Table.AddColumn(WithMfgSellVol, "bpcMarketPrice",
        each BPCPriceOf([_key]), type nullable number),

    /* -------- 7. Copy --------
       Copy install fee per CCP uses the EIV of the manufactured product,
       so copyEIV = mfgEIV. */
    WithCpyMats = Table.AddColumn(WithBpcPx, "copyBaseMaterials",
        each NormMaterials(try [activities][copying][materials] otherwise null),
        type nullable list),

    WithCpyTime = Table.AddColumn(WithCpyMats, "copyBaseTimeSec",
        each try [activities][copying][time] otherwise null, Int64.Type),

    WithCpyOut = Table.AddColumn(WithCpyTime, "copyOutputRuns",
        each try [maxProductionLimit] otherwise null, Int64.Type),

    WithCpyCost = Table.AddColumn(WithCpyOut, "copyMaterialsCostBuy",
        each MaterialsBuyCost([copyBaseMaterials]), type nullable number),

    WithCpyEIV = Table.AddColumn(WithCpyCost, "copyEIV",
        each [mfgEIV], type nullable number),

    /* -------- 8. Invention -------- */
    WithInvMats = Table.AddColumn(WithCpyEIV, "inventionBaseMaterials",
        each NormMaterials(try [activities][invention][materials] otherwise null),
        type nullable list),

    WithInvTime = Table.AddColumn(WithInvMats, "inventionBaseTimeSec",
        each try [activities][invention][time] otherwise null, Int64.Type),

    WithInvCost = Table.AddColumn(WithInvTime, "inventionMaterialsCostBuy",
        each MaterialsBuyCost([inventionBaseMaterials]), type nullable number),

    WithInvEIV = Table.AddColumn(WithInvCost, "inventionEIV",
        each [mfgEIV], type nullable number),

    WithInvProds = Table.AddColumn(WithInvEIV, "inventionProductsRaw",
        each NormInvProducts(try [activities][invention] otherwise null),
        type nullable list),

    /* -------- 9. Reaction -------- */
    WithRxnMats = Table.AddColumn(WithInvProds, "reactionBaseMaterials",
        each NormMaterials(try [activities][reaction][materials] otherwise null),
        type nullable list),

    WithRxnTime = Table.AddColumn(WithRxnMats, "reactionBaseTimeSec",
        each try [activities][reaction][time] otherwise null, Int64.Type),

    WithRxnCost = Table.AddColumn(WithRxnTime, "reactionMaterialsCostBuy",
        each MaterialsBuyCost([reactionBaseMaterials]), type nullable number),

    WithRxnEIV = Table.AddColumn(WithRxnCost, "reactionEIV",
        each EIV([reactionBaseMaterials]), type nullable number),

    WithRxnGross = Table.AddColumn(WithRxnEIV, "reactionProductGross",
        each ProductGross(try [activities][reaction] otherwise null),
        type nullable number),

    WithRxnSellVol = Table.AddColumn(WithRxnGross, "reactionProductSellVolume",
        each ProductSellVolume(try [activities][reaction] otherwise null),
        type nullable number),

    /* -------- 10. Final shape -------- */
    Final = Table.SelectColumns(WithRxnSellVol, {
        "_key",
        "productTypeId", "productQuantityPerRun",
        "mfgBaseMaterials", "mfgBaseTimeSec",
        "mfgEIV", "mfgMaterialsCostBuy", "mfgProductGross", "mfgProductSellVolume",
        "bpcMarketPrice",
        "copyBaseMaterials",
        "copyBaseTimeSec", "copyOutputRuns",
        "copyMaterialsCostBuy", "copyEIV",
        "inventionBaseMaterials",
        "inventionBaseTimeSec",
        "inventionMaterialsCostBuy", "inventionEIV",
        "inventionProductsRaw",
        "reactionBaseMaterials", "reactionBaseTimeSec",
        "reactionMaterialsCostBuy", "reactionEIV",
        "reactionProductGross", "reactionProductSellVolume"
    })
in
    Final
