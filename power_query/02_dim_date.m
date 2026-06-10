// Power Query M: Date Dimension
let
    MinDate = List.Min(fact_ar_performance[month_date]),
    MaxDate = List.Max(fact_ar_performance[month_date]),
    Count   = Number.RoundUp(Duration.Days(MaxDate - MinDate) / 30) + 2,
    Dates   = List.Dates(MinDate, Count, #duration(30,0,0,0)),
    T       = Table.TransformColumnTypes(Table.FromList(Dates, Splitter.SplitByNothing(), {"date"}), {{"date", type date}}),
    Y  = Table.AddColumn(T,  "year",        each Date.Year([date]),            Int64.Type),
    M  = Table.AddColumn(Y,  "month_num",   each Date.Month([date]),           Int64.Type),
    MN = Table.AddColumn(M,  "month_name",  each Date.ToText([date],"MMMM"),   type text),
    MS = Table.AddColumn(MN, "month_short", each Date.ToText([date],"MMM-yy"), type text),
    Q  = Table.AddColumn(MS, "quarter",     each "Q"&Text.From(Date.QuarterOfYear([date])), type text),
    SK = Table.AddColumn(Q,  "sort_key",    each Date.Year([date])*100+Date.Month([date]), Int64.Type),
    Out = Table.Sort(SK, {{"date", Order.Ascending}})
in Out
