let
    Source = Excel.CurrentWorkbook(){[Name="jitaMarketTable"]}[Content],
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"typeId", Int64.Type}, {"buyWeightedAverage", type number}, {"buyMax", type number}, {"buyMin", type number}, {"buyStddev", type number}, {"buyMedian", type number}, {"buyVolume", Int64.Type}, {"buyOrderCount", Int64.Type}, {"buyPercentile", type number}, {"sellWeightedAverage", type number}, {"sellMax", type number}, {"sellMin", type number}, {"sellStddev", type number}, {"sellMedian", type number}, {"sellVolume", Int64.Type}, {"sellOrderCount", Int64.Type}, {"sellPercentile", type number}})
in
    #"Changed Type"