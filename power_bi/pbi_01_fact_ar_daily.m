// Power BI — fact_ar_daily
//
// WHAT THIS IS
// The daily AR position, one row per Date + Country + Customer, built for a
// Power BI model rather than for an Excel worksheet.
//
// HOW THIS DIFFERS FROM THE EXCEL RawData QUERY
// In Excel, RawData reads the newest Daily_Performance_Report workbook, which
// is itself a copy-forward chain: each day's file is a copy of the previous
// day plus one appended day. A failed pipeline run drops a day out of the
// chain permanently, which is why the Excel side needs a self-healing repair
// step to patch holes back in from the source files.
//
// Power BI does not have to inherit that fragility. It has no row limit worth
// worrying about here and no reason to carry a workbook forward, so this query
// reads EVERY Cust Sol source file directly and unions them. The history is
// rebuilt from source on each refresh, so a failed pipeline run is invisible:
// the day is missing only until the source file exists, and then it is simply
// there. No backfill logic, no gap repair, no dependence on the Excel chain.
//
// SETUP (Power BI Desktop)
//   1. Home > Transform data > New Parameter:
//        Name: SharePointSite_Url   Type: Text   Current value:
//        https://dpdhl.sharepoint.com/teams/EXP-USQIA-BS33384_AMKAD
//   2. Home > New Query > Blank Query > Advanced Editor, paste this in.
//   3. Name it fact_ar_daily.
//   4. Sign in with Organizational account when prompted for the SharePoint
//      credential. Privacy level: Organizational.
//   5. Close & Apply.
//
// PERFORMANCE
// This opens every source file, so the FIRST refresh is slow (roughly a second
// or two per file). Enable incremental refresh later if the folder grows past
// a couple of years — see pbi_README.md.
let
    Site = SharePointSite_Url,
    AllFiles = SharePoint.Files(Site, [ApiVersion = 15]),

    // ---------------------------------------------------------------
    // 1. Every Cust Sol daily source file.
    // ---------------------------------------------------------------
    SrcFiles = Table.SelectRows(AllFiles,
        each Text.Contains([Folder Path], "Cust_Sol Daily File", Comparer.OrdinalIgnoreCase)
             and (Text.Lower([Extension]) = ".xlsx" or Text.Lower([Extension]) = ".xlsm")
             and not Text.StartsWith([Name], "~$")),

    // ---------------------------------------------------------------
    // 2. Map one file.
    //
    // Columns are taken by POSITION, not by header name. The source sheet has
    // "Total Payments" twice (native currency and EUR); promoting headers would
    // make Power Query rename one of them to "Total Payments_1" and the mapping
    // would silently bind to the wrong one. Position is the only stable key.
    // This is the same mapping the Office Script writer uses.
    // ---------------------------------------------------------------
    Round0 = (v) => let n = try Number.From(v) otherwise null
                    in if n = null then null else Number.Round(n, 0),

    MapOne = (content) =>
        let
            sheet = Excel.Workbook(content, null, true){0}[Data],
            // Drop the header row and any trailing blank rows. Column5 is
            // Customer: a row without one is not a data row.
            body  = Table.SelectRows(sheet, each
                        [Column1] <> null and [Column1] <> "Report Date" and [Column5] <> null),
            mapped = Table.FromRecords(List.Transform(Table.ToRecords(body), (r) => [
                Date                    = Record.Field(r, "Column1"),
                Country                 = Record.Field(r, "Column2"),
                Customer                = Record.Field(r, "Column5"),
                #"Go Live"              = Record.Field(r, "Column4"),
                #"Payment Term"         = Round0(Record.Field(r, "Column6")),
                // EUR set — what every visual should report on.
                #"Total AR"             = Round0(Record.Field(r, "Column18")),
                #"Gross Sales"          = Round0(Record.Field(r, "Column16")),
                Overdue                 = Round0(Record.Field(r, "Column19")),
                GT60                    = Round0(Record.Field(r, "Column15")),
                GT90                    = Round0(Record.Field(r, "Column20")),
                #"Total UAC"            = Round0(Record.Field(r, "Column21")),
                #"Total Payments"       = Round0(Record.Field(r, "Column22")),
                // Native-currency set, kept for reconciliation only.
                #"Total AR LC"          = Round0(Record.Field(r, "Column10")),
                #"Overdue LC"           = Round0(Record.Field(r, "Column11")),
                #"GT60 LC"              = Round0(Record.Field(r, "Column12")),
                #"Gross Sales LC"       = Round0(Record.Field(r, "Column13")),
                #"GT90 LC"              = Round0(Record.Field(r, "Column14")),
                #"Total UAC LC"         = Round0(Record.Field(r, "Column8")),
                #"Total Payments LC"    = Round0(Record.Field(r, "Column7"))
            ]))
        in
            mapped,

    // BEST EFFORT per file. One unreadable or locked file must not take the
    // whole model down — it costs that day and nothing else.
    MappedOrNull = List.Transform(SrcFiles[Content], each try MapOne(_) otherwise null),
    Usable = List.RemoveNulls(MappedOrNull),
    Combined = if List.IsEmpty(Usable)
        then error "No readable Cust Sol daily files were found. Check SharePointSite_Url and your SharePoint credential."
        else Table.Combine(Usable),

    // ---------------------------------------------------------------
    // 3. Type, then de-duplicate.
    //
    // A source file can legitimately be re-issued for the same day. Keeping
    // both copies would double every measure, so keep one row per grain.
    // ---------------------------------------------------------------
    Typed = Table.TransformColumnTypes(Combined, {
        {"Date", type date},
        {"Country", type text}, {"Customer", type text},
        {"Go Live", type date}, {"Payment Term", Int64.Type},
        {"Total AR", Currency.Type}, {"Gross Sales", Currency.Type},
        {"Overdue", Currency.Type}, {"GT60", Currency.Type}, {"GT90", Currency.Type},
        {"Total UAC", Currency.Type}, {"Total Payments", Currency.Type},
        {"Total AR LC", Currency.Type}, {"Overdue LC", Currency.Type},
        {"GT60 LC", Currency.Type}, {"Gross Sales LC", Currency.Type},
        {"GT90 LC", Currency.Type}, {"Total UAC LC", Currency.Type},
        {"Total Payments LC", Currency.Type}
    }),

    NoBlankKeys = Table.SelectRows(Typed,
        each [Date] <> null and [Country] <> null and [Customer] <> null),

    Deduped = Table.Distinct(NoBlankKeys, {"Date", "Country", "Customer"}),

    // SAP exports payments as negatives; the AR reporting shows them positive.
    PaymentsPositive = Table.TransformColumns(Deduped, {
        {"Total Payments", each if _ = null then null else Number.Abs(_), Currency.Type},
        {"Total Payments LC", each if _ = null then null else Number.Abs(_), Currency.Type}
    }),

    // ---------------------------------------------------------------
    // 4. Model helpers.
    //
    // MonthKey is text YYYY-MM so it matches the SPR sheets and the HTML
    // report's month keys. MonthDate is the real date the relationship to
    // dim_date is built on.
    // ---------------------------------------------------------------
    WithMonth = Table.AddColumn(PaymentsPositive, "MonthDate",
        each Date.StartOfMonth([Date]), type date),
    WithMonthKey = Table.AddColumn(WithMonth, "MonthKey",
        each Date.ToText([MonthDate], "yyyy-MM"), type text),

    // The percentage and DSO columns that live as formulas in the Excel sheet
    // are deliberately NOT recreated here. Percentages must not be averaged
    // across rows, so they belong in DAX as measures (see pbi_README.md);
    // a calculated column would give wrong subtotals at every level.

    Sorted = Table.Sort(WithMonthKey,
        {{"Date", Order.Ascending}, {"Country", Order.Ascending}, {"Customer", Order.Ascending}})
in
    Sorted
