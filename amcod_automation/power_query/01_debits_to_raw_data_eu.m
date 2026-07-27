// Power Query M: Current Month's Debits -> Staging (Live/EUR columns)
//
// DESIGN B ("file per month, full history carried inside the file"):
// Each monthly workbook already contains the entire 2023 -> prior-month
// history, carried forward every time the file is duplicated for a new
// month (Stage 1 - "Month Roll Forward"). So this query does NOT rebuild
// history and does NOT recurse every month folder - the history is already
// in the file. Its ONLY job is to pull THIS month's Debits block, which an
// Office Script (see office_scripts/AppendCurrentMonthToRawData.ts) then
// appends to the bottom of RawData as permanent values - exactly the manual
// "scroll to bottom, paste the new month's block" step from Stage 3.
//
// Because the appended rows are written as values, they become part of the
// history the next duplicate carries forward. Old Debits source files are
// therefore never needed - only the current month's file has to exist.
//
// LOAD THIS QUERY to a dedicated staging worksheet/table (e.g.
// "Staging_CurrentMonth"). Do NOT load it onto RawData directly - RawData
// has formula columns (DSO, TDSO Gap, %) that Power Query would destroy,
// and the append is done by the Office Script, not by this query.
//
// CONFIRMED SOURCE TABLE (a loaded query table inside the Debits workbook,
// not a worksheet tab): VW_AMKAD_Source_Debits
//
// CONFIRMED DEBITS COLUMNS (table VW_AMKAD_Source_Debits):
//   A Month | B Country | C Customer | D Go Live Customer
//   E Payment Term (days) | F DSO (NOT used - DSO is a formula in RawData)
//   G Gross Sales € | H Total AR € | I Overdue € | J >60 days € | K >90 days €
//   L Total UAC € | M Total Payments €
//   N-T = same metrics in local currency (NOT used - Live/EUR only)
//
// TARGET (RawData) Live/EUR columns the append script writes into. These are
// RawData's EXACT header strings - note "Onboard Dt" (not "Onboard Date") and
// the space after ">" in the day-bucket columns:
//   Month                   <- Debits A Month
//   Country                 <- Debits B Country
//   Customer                <- Debits C Customer
//   Onboard Dt              <- Debits D Go Live Customer
//   Payment Term (days)     <- Debits E Payment Term (days)
//   Total AR € (Live)       <- Debits H Total AR €
//   Overdue € (Live)        <- Debits I Overdue €
//   > 60 days € (Live)      <- Debits J >60 days €
//   Gross Sales € (Live)    <- Debits G Gross Sales €
//   > 90 days € (Live)      <- Debits K >90 days €
//   Total UAC € (Live)      <- Debits L Total UAC €
//   Total Payments € (Live) <- Debits M Total Payments €
//
// PARAMETER (Query Editor > Manage Parameters):
//   SharePointSite_Url  e.g. "https://dpdhl.sharepoint.com/teams/EXP-USQIA-BS33384_AMKAD"
//
// The report month is auto-calculated (see ReportMonthName below) as LAST
// month, since a month is closed during the following month. So there is no
// month parameter to advance - just refresh. To force a specific month
// (re-run an old month, or a late close), replace the ReportMonthName line
// with a literal, e.g.  ReportMonthName = "June 2026",

let
    // Auto-pick LAST month's folder. In July 2026 this yields "June 2026".
    ReportMonthName = Date.ToText(Date.AddMonths(DateTime.Date(DateTime.LocalNow()), -1), "MMMM yyyy", "en-US"),

    Source = SharePoint.Files(SharePointSite_Url, [ApiVersion = 15]),

    // Narrow to the ONE Debits file in that month's folder.
    ThisMonthFile = Table.SelectRows(Source,
        each Text.Contains([Folder Path], ReportMonthName)
             and (Text.EndsWith([Extension], ".xlsx") or Text.EndsWith([Extension], ".xlsm"))
             and Text.Contains([Name], "Debits", Comparer.OrdinalIgnoreCase)),

    AddData = Table.AddColumn(ThisMonthFile, "Data", each Excel.Workbook([Content], null, true)),
    ExpandSheets = Table.ExpandTableColumn(AddData, "Data", {"Item", "Kind", "Data"}, {"Item", "Kind", "Data"}),
    // VW_AMKAD_Source_Debits is a loaded query TABLE inside the Debits file,
    // not a worksheet tab, so match on [Item] only - do NOT restrict to
    // [Kind]="Sheet" or it finds nothing (this caused an early 0-rows bug).
    SourceSheetOnly = Table.SelectRows(ExpandSheets, each [Item] = "VW_AMKAD_Source_Debits"),

    // Debits spells the local-currency column "overdue" (lowercase) while the
    // EUR one is "Overdue €". Power Query column names are case-sensitive, so
    // asking for "Overdue" silently yields a column of nulls. Normalize first,
    // tolerating either spelling.
    FixOverdueCase = Table.TransformColumns(SourceSheetOnly, {{"Data", each
        if List.Contains(Table.ColumnNames(_), "overdue")
        then Table.RenameColumns(_, {{"overdue", "Overdue"}})
        else _}}),

    // BOTH currency sets: Debits G-M are EUR (-> RawData's "€ (Live)" columns)
    // and Debits N-T are the same metrics in local currency (-> RawData's bare
    // columns). Pulling only the EUR set left every local-currency column in
    // RawData blank, which is what they are for.
    ExpandRows = Table.ExpandTableColumn(FixOverdueCase, "Data", {
        "Month", "Country", "Customer", "Go Live Customer", "Payment Term (days)",
        "Gross Sales €", "Total AR €", "Overdue €", ">60 days €", ">90 days €",
        "Total UAC €", "Total Payments €",
        "Gross Sales", "Total AR", "Overdue", ">60 days", ">90 days",
        "Total UAC", "Total Payments"
    }),

    RemoveHeaderNoise = Table.SelectRows(ExpandRows, each [Country] <> null and [Country] <> "Country"),

    // DSO is excluded - it is a formula in RawData, not a Debits value.
    Selected = Table.SelectColumns(RemoveHeaderNoise, {
        "Month", "Country", "Customer", "Go Live Customer", "Payment Term (days)",
        "Total AR €", "Overdue €", ">60 days €", "Gross Sales €", ">90 days €",
        "Total UAC €", "Total Payments €",
        "Total AR", "Overdue", ">60 days", "Gross Sales", ">90 days",
        "Total UAC", "Total Payments"
    }),

    // Rename to RawData's EXACT header strings so the append script maps them
    // 1:1. "Payment Term (days)" and the local-currency columns already match
    // (the script's normalized matching absorbs the ">60" vs "> 60" spacing),
    // so only the EUR columns need the "€ (Live)" suffix.
    Renamed = Table.RenameColumns(Selected, {
        {"Go Live Customer", "Onboard Dt"},
        {"Total AR €", "Total AR € (Live)"},
        {"Overdue €", "Overdue € (Live)"},
        {">60 days €", "> 60 days € (Live)"},
        {"Gross Sales €", "Gross Sales € (Live)"},
        {">90 days €", "> 90 days € (Live)"},
        {"Total UAC €", "Total UAC € (Live)"},
        {"Total Payments €", "Total Payments € (Live)"}
    }),

    Typed = Table.TransformColumnTypes(Renamed, {
        {"Month", type date},
        {"Country", type text}, {"Customer", type text},
        {"Onboard Dt", type text}, {"Payment Term (days)", Int64.Type},
        {"Total AR € (Live)", Currency.Type}, {"Overdue € (Live)", Currency.Type},
        {"> 60 days € (Live)", Currency.Type}, {"Gross Sales € (Live)", Currency.Type},
        {"> 90 days € (Live)", Currency.Type},
        {"Total UAC € (Live)", Currency.Type}, {"Total Payments € (Live)", Currency.Type},
        {"Total AR", Currency.Type}, {"Overdue", Currency.Type},
        {">60 days", Currency.Type}, {"Gross Sales", Currency.Type},
        {">90 days", Currency.Type},
        {"Total UAC", Currency.Type}, {"Total Payments", Currency.Type}
    }),

    // Stage 4 "Payment Cleanup": SAP exports payments as negatives, but AMCOD
    // reports them positive (-1923 in Debits is 1923 in RawData). Flip both
    // currency versions.
    PaymentsPositive = Table.TransformColumns(Typed, {
        {"Total Payments € (Live)", each if _ = null then null else Number.Abs(_), Currency.Type},
        {"Total Payments", each if _ = null then null else Number.Abs(_), Currency.Type}
    }),

    // Empty metric cells are reported as 0, matching how the month block reads
    // today - a blank bucket means nothing aged into it, not unknown.
    MetricColumns = {
        "Total AR € (Live)", "Overdue € (Live)", "> 60 days € (Live)",
        "Gross Sales € (Live)", "> 90 days € (Live)",
        "Total UAC € (Live)", "Total Payments € (Live)",
        "Total AR", "Overdue", ">60 days", "Gross Sales", ">90 days",
        "Total UAC", "Total Payments"
    },
    ZeroFilled = Table.ReplaceValue(PaymentsPositive, null, 0, Replacer.ReplaceValue, MetricColumns),

    Sorted = Table.Sort(ZeroFilled, {{"Country", Order.Ascending}, {"Customer", Order.Ascending}})
in
    Sorted
