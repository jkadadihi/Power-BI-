// Power Query M: latest closed month's Debits -> MonthEndAR (for the SPR tab)
//
// WHY THIS EXISTS
// The daily pipeline gives the report a *daily* position. SPRs are reported on
// *month-end* numbers, and month close can shift by days, so the last daily
// snapshot of a month is NOT the month-end close. This query pulls the real
// closed-month figures so the SPR tab can show them instead of a mid-month
// snapshot. Every other tab keeps using the daily data, untouched.
//
// WHAT IT DOES
// Reads VW_AMKAD_Source_Debits (a loaded query TABLE inside the monthly Debits
// workbook, not a worksheet tab) for the LATEST CLOSED month only, and loads it
// to a table named MonthEndAR in this workbook. The report's Office Script
// reads that table by name, so no Power Automate changes are needed.
//
// SETUP
//   1. Data > Get Data > Blank Query > Advanced Editor, paste this in.
//   2. Manage Parameters: add a text parameter SharePointSite_Url, e.g.
//      "https://dpdhl.sharepoint.com/teams/EXP-USQIA-BS33384_AMKAD"
//   3. Name the query MonthEndAR, Close & Load to a NEW worksheet as a table.
//   4. In Table Design, confirm the Table Name is exactly: MonthEndAR
//
// SCOPE NOTE
// This holds the latest closed month only; each refresh replaces it. The SPR
// tab therefore shows month-end figures for that month and falls back to the
// daily snapshot (clearly labelled) for any other month selected.
//
// The header normalization, payment sign flip, and zero-fill below are carried
// over from the AMCOD month-end query, where they were found to be necessary
// against the real files: Debits header spellings are inconsistent ("overdue"
// lowercase next to "Overdue €", and a space after ">" in the day buckets), and
// Power Query column names are case- and space-sensitive, so a mismatch
// silently expands to nulls that then zero-fill and look like real zeros.
let
    // A month is closed during the following month, so "latest closed" is last
    // month. In July 2026 this yields "June 2026", matching the folder naming.
    ReportMonthName = Date.ToText(Date.AddMonths(DateTime.Date(DateTime.LocalNow()), -1), "MMMM yyyy", "en-US"),

    Source = SharePoint.Files(SharePointSite_Url, [ApiVersion = 15]),

    ThisMonthFile = Table.SelectRows(Source,
        each Text.Contains([Folder Path], ReportMonthName)
             and (Text.EndsWith([Extension], ".xlsx") or Text.EndsWith([Extension], ".xlsm"))
             and Text.Contains([Name], "Debits", Comparer.OrdinalIgnoreCase)),

    AddData = Table.AddColumn(ThisMonthFile, "Data", each Excel.Workbook([Content], null, true)),
    ExpandSheets = Table.ExpandTableColumn(AddData, "Data", {"Item", "Kind", "Data"}, {"Item", "Kind", "Data"}),

    // Match on [Item] only. VW_AMKAD_Source_Debits is a loaded query table, not
    // a worksheet, so restricting to [Kind]="Sheet" finds nothing.
    SourceSheetOnly = Table.SelectRows(ExpandSheets, each [Item] = "VW_AMKAD_Source_Debits"),

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

    // EUR set only. The SPR panel reports in euros, matching the review deck.
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
        {"Total AR €", "Total AR"},
        {"Overdue €", "Overdue"},
        {">60 days €", "GT60"},
        {">90 days €", "GT90"},
        {"Total UAC €", "Total UAC"},
        {"Gross Sales €", "Gross Sales"},
        {"Total Payments €", "Total Payments"}
    }),

    Typed = Table.TransformColumnTypes(Renamed, {
        {"Country", type text}, {"Customer", type text},
        {"Total AR", Currency.Type}, {"Overdue", Currency.Type},
        {"GT60", Currency.Type}, {"GT90", Currency.Type},
        {"Total UAC", Currency.Type}, {"Gross Sales", Currency.Type},
        {"Total Payments", Currency.Type}
    }),

    // Month must be text in YYYY-MM to line up with the SPR sheets' Month
    // column and the report's own month keys. The source may hand it over as a
    // date or as text, so handle both rather than assuming.
    MonthAsKey = Table.TransformColumns(Typed, {{"Month", each
        let
            asDate = try Date.From(_) otherwise null
        in
            if asDate <> null then Date.ToText(asDate, "yyyy-MM")
            else Text.From(_), type text}}),

    // SAP exports payments as negatives; AMCOD reports them positive.
    PaymentsPositive = Table.TransformColumns(MonthAsKey, {
        {"Total Payments", each if _ = null then null else Number.Abs(_), Currency.Type}
    }),

    MetricColumns = {"Total AR", "Overdue", "GT60", "GT90",
                     "Total UAC", "Gross Sales", "Total Payments"},
    ZeroFilled = Table.ReplaceValue(PaymentsPositive, null, 0, Replacer.ReplaceValue, MetricColumns),

    Sorted = Table.Sort(ZeroFilled, {{"Country", Order.Ascending}, {"Customer", Order.Ascending}})
in
    Sorted
