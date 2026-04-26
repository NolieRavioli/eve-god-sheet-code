let
    Source = Excel.CurrentWorkbook(){[Name="decryptors"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Decryptor Name", type text}, {"Probability Multiplier", type number}, {"Max Run Modifier", Int64.Type}, {"ME Modifier", Int64.Type}, {"PE Modifier", Int64.Type}})
in
    #"Changed Type"