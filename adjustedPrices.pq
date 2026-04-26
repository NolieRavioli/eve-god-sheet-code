let
    /* -------- 1. Source --------
       ESI /markets/prices/ returns the daily CCP-computed
       adjusted_price (used for Estimated Item Value / job install fees)
       and average_price (rolling 28d). One row per typeID that has data.   */
    Source = Json.Document(
        Web.Contents("https://esi.evetech.net/latest/markets/prices/")),

    AsTable = Table.FromList(Source, Splitter.SplitByNothing(),
                             null, null, ExtraValues.Error),

    Expanded = Table.ExpandRecordColumn(
        AsTable, "Column1",
        {"type_id", "adjusted_price", "average_price"},
        {"typeId",  "adjustedPrice",  "averagePrice"}),

    Typed = Table.TransformColumnTypes(Expanded,
                {{"typeId",        Int64.Type},
                 {"adjustedPrice", type number},
                 {"averagePrice",  type number}})
in
    Typed
