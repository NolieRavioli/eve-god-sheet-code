let
    Source = Csv.Document(File.Contents("C:\Users\nolan\eve-god-sheet-code\sde-csv\types.csv"),[Delimiter=",", Columns=20, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",
        {{"_key", Int64.Type}, {"basePrice", Int64.Type}, {"capacity", type number},
         {"description", type text}, {"factionID", type text}, {"graphicID", Int64.Type},
         {"groupID", Int64.Type}, {"iconID", Int64.Type}, {"marketGroupID", Int64.Type},
         {"mass", type number}, {"metaGroupID", Int64.Type}, {"metaLevel", Int64.Type},
         {"portionSize", Int64.Type}, {"published", type logical}, {"raceID", Int64.Type},
         {"radius", Int64.Type}, {"soundID", Int64.Type},
         {"variationParentTypeID", Int64.Type}, {"volume", type number}}),
    #"Parsed JSON" = Table.TransformColumns(#"Changed Type",{{"name", Json.Document}})
in
    #"Parsed JSON"