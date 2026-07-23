// Power Query M: Debits (Refreshable) -> Raw Data EU "Live"/EUR columns
//
// Replaces AMCOD Stage 3 ("Update Raw Data EU"), the single biggest manual
// step in the process: open the Debits workbook, open the Monthly workbook,
// scroll to the bottom of Raw Data EU, copy a block, paste values, fix
// formatting, update the month indicator - repeated every close.
//
// CONFIRMED FOLDER STRUCTURE (from the actual SharePoint library):
//   .../2026 Presentations/2026 AMKAD Summary/Qtr 2/June 2026/
//     AMKAD Main Metrics Development Template...
//     AMKAD_KPI_ViewRefreshable_Debits Only ...      <- the source file
//     GPMR_CFSC_AMKAD_202606_V1.xlsm
//     Monthly_Performance_Repo...
//     SSC (KAD) Monthly Data File.zip
// A new dated subfolder is created every month (Stage 1 - "Month Roll
// Forward"), nested under a quarter folder, nested under a year folder. The
// Debits file itself keeps the SAME name every month - only its folder
// changes - so:
//   - the month/year is parsed from the FOLDER path, not the filename
//   - the query points at the stable parent folder ("2026 AMKAD Summary")
//     and lets SharePoint.Files recurse through every Qtr/Month subfolder,
//     so a brand new month folder is picked up on refresh with NO edits to
//     this query. The only maintenance this needs is once a year, when the
//     "2026" year folder itself rolls to "2027" - update
//     DebitsRoot_RelativePath then.
//   - other workbooks living in the same folder (Main Metrics Template,
//     GPMR, Monthly Performance Report, the .zip) are excluded by filtering
//     on filename containing "Debits".
//
// CONFIRMED SHEET NAME (inside the Debits workbook): VW_AMKAD_Source_Debits
//
// CONFIRMED DEBITS COLUMNS (workbook: "AMKAD_KPI_ViewRefreshable_Debits Only"):
//   A Month | B Country | C Customer | D Go Live Customer
//   E Payment Term (days) | F DSO
//   G Gross Sales € | H Total AR € | I Overdue € | J >60 days € | K >90 days €
//   L Total UAC € | M Total Payments €
//   N Gross Sales | O Total AR | P Overdue | Q >60 days | R >90 days
//   S Total UAC | T Total Payments
// (columns G-M are EUR; N-T are the same metrics in local currency)
//
// This query ONLY maps the EUR ("Live") columns into Raw Data EU. The
// local-currency columns (N-T) and the Monthly file's TDSO/TDSO Gap/% and
// K-O columns are intentionally left untouched - they are not sourced from
// Debits, so nothing is pasted into them.
//
// TARGET (Raw Data EU / Monthly file) MAPPING:
//   A Month              <- Debits A Month
//   B Country             <- Debits B Country
//   C Customer            <- Debits C Customer
//   D Onboard Date        <- Debits D Go Live Customer
//   E Payment (days)      <- Debits E Payment Term (days)
//   G DSO                 <- Debits F DSO
//   P Total AR (Live)     <- Debits H Total AR €
//   Q Overdue (Live)      <- Debits I Overdue €
//   R >60 days (Live)     <- Debits J >60 days €
//   S Gross Sales (Live)  <- Debits G Gross Sales €
//   T >90 days (Live)     <- Debits K >90 days €
//   Z Total UAC € (Live)      <- Debits L Total UAC €
//   AB Total Payments € (Live) <- Debits M Total Payments €
//
// Load this query's output to a staging table (e.g. "Debits Live Feed")
// next to Raw Data EU rather than directly overwriting Raw Data EU itself,
// if Raw Data EU already has manual formula columns (TDSO Gap, %, local
// currency, etc.) mixed in - Power Query owns every column of whatever
// table it's loaded into, so a table with hand-built formula columns next
// to query-fed columns can't be the direct load target of this query.
//
// PARAMETERS (set these once, Query Editor > Manage Parameters):
//   SharePointSite_Url        e.g. "https://dhl.sharepoint.com/sites/AMROFinance"
//   DebitsRoot_RelativePath   e.g. "/2026 Presentations/2026 AMKAD Summary"
//                             (update the "2026" segment once a year)

let
    Source = SharePoint.Files(SharePointSite_Url, [ApiVersion = 15]),

    UnderDebitsRoot = Table.SelectRows(Source,
        each Text.Contains([Folder Path], DebitsRoot_RelativePath)),

    ExcelFiles = Table.SelectRows(UnderDebitsRoot,
        each (Text.EndsWith([Extension], ".xlsx") or Text.EndsWith([Extension], ".xlsm"))
             and Text.Contains([Name], "Debits", Comparer.OrdinalIgnoreCase)),

    // Folder path ends in ".../Qtr 2/June 2026/" - the last non-empty
    // segment is the month folder name, e.g. "June 2026".
    AddMonthRaw = Table.AddColumn(ExcelFiles, "month_raw",
        each List.Last(List.RemoveItems(Text.Split(Text.TrimEnd([Folder Path], "/"), "/"), {""})),
        type text),

    AddData = Table.AddColumn(AddMonthRaw, "Data", each Excel.Workbook([Content], null, true)),
    ExpandSheets = Table.ExpandTableColumn(AddData, "Data", {"Item", "Kind", "Data"}, {"Item", "Kind", "Data"}),
    SourceSheetOnly = Table.SelectRows(ExpandSheets, each [Kind] = "Sheet" and [Item] = "VW_AMKAD_Source_Debits"),

    ExpandRows = Table.ExpandTableColumn(SourceSheetOnly, "Data", {
        "Month", "Country", "Customer", "Go Live Customer", "Payment Term (days)", "DSO",
        "Gross Sales €", "Total AR €", "Overdue €", ">60 days €", ">90 days €",
        "Total UAC €", "Total Payments €"
    }),

    RemoveHeaderNoise = Table.SelectRows(ExpandRows, each [Country] <> null and [Country] <> "Country"),

    // Map straight through to the Raw Data EU column names confirmed above.
    // Local-currency columns (N-T in Debits) are deliberately not selected.
    Selected = Table.SelectColumns(RemoveHeaderNoise, {
        "Month", "Country", "Customer", "Go Live Customer", "Payment Term (days)", "DSO",
        "Total AR €", "Overdue €", ">60 days €", "Gross Sales €", ">90 days €",
        "Total UAC €", "Total Payments €"
    }),

    Renamed = Table.RenameColumns(Selected, {
        {"Month", "Month"},
        {"Country", "Country"},
        {"Customer", "Customer"},
        {"Go Live Customer", "Onboard Date"},
        {"Payment Term (days)", "Payment (days)"},
        {"DSO", "DSO"},
        {"Total AR €", "Total AR (Live)"},
        {"Overdue €", "Overdue (Live)"},
        {">60 days €", ">60 days (Live)"},
        {"Gross Sales €", "Gross Sales (Live)"},
        {">90 days €", ">90 days (Live)"},
        {"Total UAC €", "Total UAC € (Live)"},
        {"Total Payments €", "Total Payments € (Live)"}
    }),

    Typed = Table.TransformColumnTypes(Renamed, {
        {"Month", type date},
        {"Country", type text}, {"Customer", type text},
        {"Onboard Date", type text}, {"Payment (days)", Int64.Type},
        {"DSO", type number},
        {"Total AR (Live)", Currency.Type}, {"Overdue (Live)", Currency.Type},
        {">60 days (Live)", Currency.Type}, {"Gross Sales (Live)", Currency.Type},
        {">90 days (Live)", Currency.Type},
        {"Total UAC € (Live)", Currency.Type}, {"Total Payments € (Live)", Currency.Type}
    }),

    Sorted = Table.Sort(Typed, {{"Month", Order.Ascending}, {"Country", Order.Ascending}})
in
    Sorted
