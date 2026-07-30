// Power BI — fact_ar_monthly (latest snapshot per month)
//
// WHY THIS EXISTS
// fact_ar_daily holds every daily snapshot. AR is a BALANCE, not a flow: you
// cannot sum it across days. Summing 20 daily snapshots of a €4m book gives
// €80m, which is meaningless. Any month-level visual therefore has to pick ONE
// snapshot per month, and the right one is the last one taken in that month.
//
// You can do this in DAX with LASTNONBLANK, but doing it here keeps the model
// simple and the visuals fast, and it makes the choice visible to whoever
// picks this file up next.
//
// IMPORTANT — THIS IS NOT MONTH-END
// The last daily snapshot of a month is the last one the pipeline captured,
// which is not the same as the closed-month position: month close shifts by a
// few days, and postings land after the final snapshot. For SPR / Account
// Business Review reporting, use fact_ar_month_end instead, which reads the
// real closed figures. This query is for trend visuals, where a consistent
// same-basis series matters more than exact close.
//
// SETUP
//   New Query > Blank Query > Advanced Editor, paste, name it fact_ar_monthly.
//   It references fact_ar_daily, so create that one first.
let
    Source = fact_ar_daily,

    // The latest date actually captured in each month.
    LatestPerMonth = Table.Group(Source, {"MonthKey"}, {
        {"SnapshotDate", each List.Max([Date]), type date}
    }),

    // Keep only rows from that date. An inner join is the cheap way to do it.
    Joined = Table.Join(Source, {"MonthKey"}, LatestPerMonth, {"MonthKey"}, JoinKind.Inner),
    LatestOnly = Table.SelectRows(Joined, each [Date] = [SnapshotDate]),

    // SnapshotDate is now identical to Date on every surviving row, so drop it
    // but keep a flag the report can surface: how stale is this month's number?
    Cleaned = Table.RemoveColumns(LatestOnly, {"SnapshotDate"}),

    WithBasis = Table.AddColumn(Cleaned, "Basis", each "Daily snapshot", type text),
    WithLag = Table.AddColumn(WithBasis, "DaysFromMonthEnd",
        each Duration.Days(Date.EndOfMonth([MonthDate]) - [Date]), Int64.Type),

    Sorted = Table.Sort(WithLag,
        {{"MonthDate", Order.Ascending}, {"Country", Order.Ascending}, {"Customer", Order.Ascending}})
in
    Sorted
