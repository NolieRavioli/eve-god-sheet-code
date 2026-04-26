let
    Source = Excel.CurrentWorkbook(){[Name="settings"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"setting", type text}, {"value", type any}}),
    #"Transposed Table" = Table.Transpose(#"Changed Type"),
    #"Promoted Headers" = Table.PromoteHeaders(#"Transposed Table", [PromoteAllScalars=true]),
    #"Changed Type1" = Table.TransformColumnTypes(#"Promoted Headers",{{"Language", type text}, {"Location", type text}, {"buy method", type text}, {"sell method", type text}, {"tax %", type number}, {"broker %", type number}, {"Sell modifier", type number}, {"Target Profit", type number}, {"store", type number}, {"store type", type text}, {"buyback", type number}, {"buyback type", type text}})
in
    #"Changed Type1"