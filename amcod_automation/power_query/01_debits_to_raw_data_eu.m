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
// CONFIRMED SHEET NAME (inside the Debits workbook): VW_AMKAD_Source_Debits
//
// CONFIRMED DEBITS COLUMNS (sheet VW_AMKAD_Source_Debits):
//   A Month | B Country | C Customer | D Go Live Customer
//   E Payment Term (days) | F DSO (NOT used - DSO is a formula in RawData)
//   G Gross Sales € | H Total AR € | I Overdue € | J >60 days € | K >90 days €
//   L Total UAC € | M Total Payments €
//   N-T = same metrics in local currency (NOT used - Live/EUR only)
//
// TARGET (RawData) Live/EUR columns the append script writes into:
//   Month              <- Debits A Month
//   Country            <- Debits B Country
//   Customer           <- Debits C Customer
//   Onboard Date       <- Debits D Go Live Customer
//   Payment (days)     <- Debits E Payment Term (days)
//   Total AR € (Live)      <- Debits H Total AR €
//   Overdue € (Live)       <- Debits I Overdue €
//   >60 days € (Live)      <- Debits J >60 days €
//   Gross Sales (Live)     <- Debits G Gross Sales €
//   >90 days (Live)        <- Debits K >90 days €
//   Total UAC € (Live)     <- Debits L Total UAC €
//   Total Payments € (Live) <- Debits M Total Payments €
//
// PARAMETERS (Query Editor > Manage Parameters):
//   SharePointSite_Url      e.g. "https://dhl.sharepoint.com/sites/AMROFinance"
//   ReportMonth_FolderName  the current month's folder name, e.g. "July 2026".
//                           The Power Automate flow sets this each month so a
//                           freshly duplicated file pulls the RIGHT month
//                           instead of the month it was copied from.

let
    Source = SharePoint.Files(SharePointSite_Url, [ApiVersion = 15]),

    // Narrow to the ONE Debits file in this month's folder.
    ThisMonthFile = Table.SelectRows(Source,
        each Text.Contains([Folder Path], ReportMonth_FolderName)
             and (Text.EndsWith([Extension], ".xlsx") or Text.EndsWith([Extension], ".xlsm"))
             and Text.Contains([Name], "Debits", Comparer.OrdinalIgnoreCase)),

    AddData = Table.AddColumn(ThisMonthFile, "Data", each Excel.Workbook([Content], null, true)),
    ExpandSheets = Table.ExpandTableColumn(AddData, "Data", {"Item", "Kind", "Data"}, {"Item", "Kind", "Data"}),
    SourceSheetOnly = Table.SelectRows(ExpandSheets, each [Kind] = "Sheet" and [Item] = "VW_AMKAD_Source_Debits"),

    ExpandRows = Table.ExpandTableColumn(SourceSheetOnly, "Data", {
        "Month", "Country", "Customer", "Go Live Customer", "Payment Term (days)",
        "Gross Sales €", "Total AR €", "Overdue €", ">60 days €", ">90 days €",
        "Total UAC €", "Total Payments €"
    }),

    RemoveHeaderNoise = Table.SelectRows(ExpandRows, each [Country] <> null and [Country] <> "Country"),

    // Live/EUR columns only. DSO and local-currency columns are excluded.
    Selected = Table.SelectColumns(RemoveHeaderNoise, {
        "Month", "Country", "Customer", "Go Live Customer", "Payment Term (days)",
        "Total AR €", "Overdue €", ">60 days €", "Gross Sales €", ">90 days €",
        "Total UAC €", "Total Payments €"
    }),

    Renamed = Table.RenameColumns(Selected, {
        {"Go Live Customer", "Onboard Date"},
        {"Payment Term (days)", "Payment (days)"},
        {"Total AR €", "Total AR € (Live)"},
        {"Overdue €", "Overdue € (Live)"},
        {">60 days €", ">60 days € (Live)"},
        {"Gross Sales €", "Gross Sales (Live)"},
        {">90 days €", ">90 days (Live)"},
        {"Total UAC €", "Total UAC € (Live)"},
        {"Total Payments €", "Total Payments € (Live)"}
    }),

    Typed = Table.TransformColumnTypes(Renamed, {
        {"Month", type date},
        {"Country", type text}, {"Customer", type text},
        {"Onboard Date", type text}, {"Payment (days)", Int64.Type},
        {"Total AR € (Live)", Currency.Type}, {"Overdue € (Live)", Currency.Type},
        {">60 days € (Live)", Currency.Type}, {"Gross Sales (Live)", Currency.Type},
        {">90 days (Live)", Currency.Type},
        {"Total UAC € (Live)", Currency.Type}, {"Total Payments € (Live)", Currency.Type}
    }),

    Sorted = Table.Sort(Typed, {{"Country", Order.Ascending}, {"Customer", Order.Ascending}})
in
    Sorted
