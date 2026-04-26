let
    Source = Excel.CurrentWorkbook(){[Name="jitaContractsTable"]}[Content],
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    /* Defensive: the upstream Office Script can occasionally spill an
       error/diagnostic row whose typeId cell is non-numeric text.
       Filter those out BEFORE typing so the whole query doesn't fail. */
    Cleaned = Table.SelectRows(#"Promoted Headers", each
        let
            t = try Number.FromText(Text.Trim(Text.From([typeId]))) otherwise null,
            p = try Number.FromText(Text.Trim(Text.From([price])))  otherwise null
        in
            t <> null and t > 0 and p <> null and p > 0),
    #"Changed Type" = Table.TransformColumnTypes(Cleaned,{{"typeId", Int64.Type}, {"price", type number}})
in
    #"Changed Type"