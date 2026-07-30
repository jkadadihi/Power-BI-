// Power BI — fact_ar_month_end (real closed-month figures)
//
// WHY THIS EXISTS
// SPRs / Account Business Reviews are reported on month-end closed numbers.
// The daily pipeline cannot produce those: month close shifts by days, and
// postings land after the final daily snapshot, so the last snapshot of a
// month is close to the close but not the close.
//
// This reads VW_AMKAD_Source_Debits out of the monthly Debits workbooks, which
// is the closed position. It is the Power BI port of power_query/spr_month_end_ar.m,
// with one deliberate change: the Excel version pulls the LATEST CLOSED MONTH
// ONLY, because it loads to a worksheet that gets replaced on every refresh.
// Power BI has no such constraint, so this pulls EVERY month it can find and
// gives you a real month-end time series to trend on.
//
// SETUP
//   New Query > Blank Query > Advanced Editor, paste, name it fact_ar_month_end.
//   Uses the same SharePointSite_Url parameter as fact_ar_daily.
let
    Source = SharePoint.Files(SharePointSite_Url, [ApiVersion = 15]),

    DebitsFiles = Table.SelectRows(Source,
        each (Text.EndsWith([Extension], ".xlsx") or Text.EndsWith([Extension], ".xlsm"))
             and Text.Contains([Name], "Debits", Comparer.OrdinalIgnoreCase)
             and not Text.StartsWith([Name], "~$")),

    AddData = Table.AddColumn(DebitsFiles, "Data", each
        try Excel.Workbook([Content], null, true) otherwise null),
    Readable = Table.SelectRows(AddData, each [Data] <> null),
    ExpandSheets = Table.ExpandTableColumn(Readable, "Data",
        {"Item", "Kind", "Data"}, {"Item", "Kind", "Data"}),

    // Match on [Item] only. VW_AMKAD_Source_Debits is a loaded QUERY table,
    // not a worksheet, so filtering [Kind] = "Sheet" finds nothing.
    SourceSheetOnly = Table.SelectRows(ExpandSheets, each [Item] = "VW_AMKAD_Source_Debits"),

    // Debits header spellings are inconsistent between files: lowercase
    // "overdue" next to "Overdue €", a space after ">" in the day buckets.
    // Power Query column names are case- and space-sensitive, so an unmatched
    // name silently expands to nulls, which then zero-fill and look like real
    // zeros. Normalise before expanding.
    CanonicalNames = {
        "Month", "Country", "Customer", "Go Live Customer", "Payment Term (days)",
        "Gross Sales €", "Total AR €", "Overdue €", ">60 days €", ">90 days €",
        "Total UAC €", "Total Payments €"
    },
    NormalizeKey = (name as text) as text => Text.Lower(Text.Remove(name, {" "})),
    NormalizeHeaders = Table.TransformColumns(SourceSheetOnly, {{"Data", each
        let
            actual = Table.ColumnNames(_),
            renames = List.RemoveNulls(List.Transform(actual, (a) =>
                let
                    match = List.First(
                        List.Select(CanonicalNames, (c) => NormalizeKey(c) = NormalizeKey(a)),
                        null)
                in
                    if match <> null and match <> a then {a, match} else null))
        in
            Table.RenameColumns(_, renames)}}),

    ExpandRows = Table.ExpandTableColumn(NormalizeHeaders, "Data", {
        "Month", "Country", "Customer",
        "Gross Sales €", "Total AR €", "Overdue €", ">60 days €", ">90 days €",
        "Total UAC €", "Total Payments €"
    }),

    RemoveHeaderNoise = Table.SelectRows(ExpandRows,
        each [Country] <> null and [Country] <> "Country" and [Customer] <> null),

    Selected = Table.SelectColumns(RemoveHeaderNoise, {
        "Month", "Country", "Customer",
        "Total AR €", "Overdue €", ">60 days €", ">90 days €",
        "Total UAC €", "Gross Sales €", "Total Payments €"
    }),

    Renamed = Table.RenameColumns(Selected, {
        {"Total AR €", "Total AR"}, {"Overdue €", "Overdue"},
        {">60 days €", "GT60"}, {">90 days €", "GT90"},
        {"Total UAC €", "Total UAC"}, {"Gross Sales €", "Gross Sales"},
        {"Total Payments €", "Total Payments"}
    }),

    // Month may arrive as a date or as text depending on the file. Coerce to a
    // real month-start date so it can join dim_date, and keep the text key too
    // so it lines up with fact_ar_daily and the SPR sheets.
    WithMonthDate = Table.AddColumn(Renamed, "MonthDate", each
        let d = try Date.From([Month]) otherwise
                try Date.FromText(Text.From([Month]) & "-01") otherwise null
        in if d = null then null else Date.StartOfMonth(d), type date),
    DatedOnly = Table.SelectRows(WithMonthDate, each [MonthDate] <> null),
    WithMonthKey = Table.AddColumn(DatedOnly, "MonthKey",
        each Date.ToText([MonthDate], "yyyy-MM"), type text),

    Typed = Table.TransformColumnTypes(WithMonthKey, {
        {"Country", type text}, {"Customer", type text},
        {"Total AR", Currency.Type}, {"Overdue", Currency.Type},
        {"GT60", Currency.Type}, {"GT90", Currency.Type},
        {"Total UAC", Currency.Type}, {"Gross Sales", Currency.Type},
        {"Total Payments", Currency.Type}
    }),

    // SAP exports payments as negatives; AMCOD reports them positive.
    PaymentsPositive = Table.TransformColumns(Typed, {
        {"Total Payments", each if _ = null then null else Number.Abs(_), Currency.Type}
    }),

    MetricColumns = {"Total AR", "Overdue", "GT60", "GT90",
                     "Total UAC", "Gross Sales", "Total Payments"},
    ZeroFilled = Table.ReplaceValue(PaymentsPositive, null, 0, Replacer.ReplaceValue, MetricColumns),

    // If the same month appears in two Debits workbooks, the later file wins by
    // being kept last is NOT safe to assume, so key the grain explicitly.
    Deduped = Table.Distinct(ZeroFilled, {"MonthKey", "Country", "Customer"}),

    WithBasis = Table.AddColumn(Deduped, "Basis", each "Month-end close", type text),
    Final = Table.SelectColumns(WithBasis, {
        "MonthDate", "MonthKey", "Country", "Customer", "Basis",
        "Total AR", "Overdue", "GT60", "GT90",
        "Total UAC", "Gross Sales", "Total Payments"
    }),

    Sorted = Table.Sort(Final,
        {{"MonthDate", Order.Ascending}, {"Country", Order.Ascending}, {"Customer", Order.Ascending}})
in
    Sorted
