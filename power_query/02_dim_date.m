// Power Query M Script: Generate Date Dimension Table
// Covers the full range of months present in fact_ar_performance

let
    // ── Pull min / max dates from the fact table ──────────────────────────────
    MinDate = List.Min(fact_ar_performance[month_date]),
    MaxDate = List.Max(fact_ar_performance[month_date]),

    // ── Generate list of months between min and max ───────────────────────────
    MonthCount = Duration.Days(MaxDate - MinDate) / 30 + 2,
    DateList   = List.Dates(MinDate, Number.RoundUp(MonthCount), #duration(30, 0, 0, 0)),

    // ── Convert to table ──────────────────────────────────────────────────────
    DateTable = Table.FromList(DateList, Splitter.SplitByNothing(), {"date"}),
    TypedDates = Table.TransformColumnTypes(DateTable, {{"date", type date}}),

    // ── Add calendar columns ──────────────────────────────────────────────────
    WithYear    = Table.AddColumn(TypedDates, "year",        each Date.Year([date]),            Int64.Type),
    WithMonth   = Table.AddColumn(WithYear,   "month_num",   each Date.Month([date]),           Int64.Type),
    WithMonthNm = Table.AddColumn(WithMonth,  "month_name",  each Date.ToText([date], "MMMM"),  type text),
    WithMonthSh = Table.AddColumn(WithMonthNm,"month_short", each Date.ToText([date], "MMM-yy"),type text),
    WithQuarter = Table.AddColumn(WithMonthSh,"quarter",     each "Q" & Text.From(Date.QuarterOfYear([date])), type text),
    WithSortKey = Table.AddColumn(WithQuarter,"sort_key",
                    each Date.Year([date]) * 100 + Date.Month([date]),
                    Int64.Type),

    FinalDimDate = Table.Sort(WithSortKey, {{"date", Order.Ascending}})

in
    FinalDimDate
