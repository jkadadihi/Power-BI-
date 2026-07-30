// Power BI — dimension tables
//
// Three separate queries. Paste each block into its own Blank Query and name
// it as marked. They are in one file only so they are easy to find.
//
// WHY DIMENSIONS AT ALL
// fact_ar_daily and fact_ar_month_end are two different grains of the same
// business entities. If visuals slice each fact by its OWN Country/Customer
// columns, a slicer on one will not filter the other and the two bases will
// disagree on screen. Shared dimensions fix that: one Country slicer filters
// both facts, so the daily and month-end numbers on an SPR page always refer
// to the same account.
//
// RELATIONSHIPS TO BUILD (Model view, all single-direction, one-to-many
// from the dimension to the fact):
//   dim_date[date]        -> fact_ar_daily[Date]
//   dim_date[date]        -> fact_ar_monthly[MonthDate]
//   dim_date[date]        -> fact_ar_month_end[MonthDate]
//   dim_customer[CustKey] -> fact_ar_daily[CustKey]      (add via the step below)
//   dim_customer[CustKey] -> fact_ar_month_end[CustKey]
//
// NOTE ON CustKey: a customer name can appear in more than one country and is
// a DIFFERENT account each time, so the key must be Country + Customer, never
// Customer alone. Add this step to the end of each fact query, just before the
// final Sort:
//
//     WithCustKey = Table.AddColumn(<previous step>, "CustKey",
//         each Text.Upper(Text.Trim(Text.From([Country]))) & "|" &
//              Text.Upper(Text.Trim(Text.From([Customer]))), type text),
//
// The upper/trim normalisation matters: the source has trailing spaces on some
// names, and an unnormalised key splits one account into two phantom rows.


// =====================================================================
// QUERY NAME: dim_date
// =====================================================================
let
    // Span both facts so neither has dates outside the dimension. A fact row
    // with no matching date row lands in a blank member and disappears from
    // every date-filtered visual.
    DailyDates = List.RemoveNulls(fact_ar_daily[Date]),
    EomDates   = List.RemoveNulls(fact_ar_month_end[MonthDate]),
    AllDates   = List.Combine({DailyDates, EomDates}),

    MinDate = Date.StartOfYear(List.Min(AllDates)),
    MaxDate = Date.EndOfYear(List.Max(AllDates)),
    DayCount = Duration.Days(MaxDate - MinDate) + 1,

    Dates = List.Dates(MinDate, DayCount, #duration(1, 0, 0, 0)),
    T  = Table.TransformColumnTypes(
            Table.FromList(Dates, Splitter.SplitByNothing(), {"date"}),
            {{"date", type date}}),
    Y  = Table.AddColumn(T,  "year",         each Date.Year([date]),                    Int64.Type),
    M  = Table.AddColumn(Y,  "month_num",    each Date.Month([date]),                   Int64.Type),
    MD = Table.AddColumn(M,  "month_start",  each Date.StartOfMonth([date]),            type date),
    MK = Table.AddColumn(MD, "month_key",    each Date.ToText([date], "yyyy-MM"),       type text),
    MN = Table.AddColumn(MK, "month_name",   each Date.ToText([date], "MMMM", "en-US"), type text),
    MS = Table.AddColumn(MN, "month_short",  each Date.ToText([date], "MMM-yy", "en-US"), type text),
    Q  = Table.AddColumn(MS, "quarter",      each "Q" & Text.From(Date.QuarterOfYear([date])), type text),
    // Sort keys so month names order chronologically instead of alphabetically.
    // Without these, a chart axis reads April, August, December, February.
    SK = Table.AddColumn(Q,  "month_sort",   each Date.Year([date]) * 100 + Date.Month([date]), Int64.Type),
    EOM = Table.AddColumn(SK, "is_month_end", each [date] = Date.EndOfMonth([date]), type logical),
    Out = Table.Sort(EOM, {{"date", Order.Ascending}})
in
    Out


// =====================================================================
// QUERY NAME: dim_customer
//
// Built from both facts so an account that exists in one basis but not the
// other still gets a dimension row and does not vanish from slicers.
// =====================================================================
let
    FromDaily = Table.SelectColumns(fact_ar_daily, {"Country", "Customer"}),
    FromEom   = Table.SelectColumns(fact_ar_month_end, {"Country", "Customer"}),
    Combined  = Table.Combine({FromDaily, FromEom}),

    Trimmed = Table.TransformColumns(Combined, {
        {"Country",  each Text.Trim(Text.From(_)), type text},
        {"Customer", each Text.Trim(Text.From(_)), type text}
    }),
    NoBlanks = Table.SelectRows(Trimmed, each [Country] <> "" and [Customer] <> ""),
    Unique = Table.Distinct(NoBlanks, {"Country", "Customer"}),

    WithKey = Table.AddColumn(Unique, "CustKey",
        each Text.Upper([Country]) & "|" & Text.Upper([Customer]), type text),
    // Re-dedupe on the NORMALISED key: "Nokia" and "Nokia " are one account.
    KeyUnique = Table.Distinct(WithKey, {"CustKey"}),

    Sorted = Table.Sort(KeyUnique, {{"Country", Order.Ascending}, {"Customer", Order.Ascending}})
in
    Sorted


// =====================================================================
// QUERY NAME: dim_basis
//
// A tiny disconnected table so a report page can offer a "Current data /
// EOM data" toggle, the same confirmation label the HTML report shows.
// Pair it with the SelectedBasis measures in pbi_README.md.
// =====================================================================
let
    Source = #table(
        type table [Basis = text, BasisSort = Int64.Type],
        {
            {"Month-end close", 1},
            {"Daily snapshot",  2}
        })
in
    Source
