// Power Query (M) — Raw Data EUR engine
// ---------------------------------------------------------------------------
// Combines EVERY daily file in the Cust_Sol Daily File SharePoint folder into
// one full-history table mapped to the Raw Data EUR layout. Each refresh
// rebuilds the complete union from whatever files are in the folder, so it is
// self-healing: drop in a corrected file (or delete a bad one) and the next
// refresh just reflects reality. No append logic, no duplicate guard needed —
// re-reading the same folder produces the same result.
//
// Headers are intentionally NOT promoted. The source report has the header
// "Total Payments" TWICE (native col G and Euro col V); promoting headers would
// force Power Query to rename the second to "Total Payments.1" and make the
// mapping fragile. Working by column POSITION instead keeps it unambiguous and
// also survives the source ever renaming a header.
//
// Output = 19 columns (the A:E + Q:AD data columns). F:P are omitted on purpose
// (not needed per current scope). If a downstream reference needs the full
// A:AD shape, add null placeholder columns for F:P after this query.
//
// SETUP: paste into a blank query (Data > Get Data > Blank Query > Advanced
// Editor) in the engine workbook. Name the query "RawData", Close & Load to a
// Table, then in Table Design set the Table Name to exactly "RawData" — the
// downstream HTML query finds the data by that literal table name.
let
    // ==== CONFIG — change only these two if the location moves ====
    SiteUrl      = "https://dpdhl.sharepoint.com/teams/EXP-USQIA-BS33384_AMKAD",
    FolderMarker = "Cust_Sol Daily File",

    Source    = SharePoint.Files(SiteUrl, [ApiVersion = 15]),
    InFolder  = Table.SelectRows(Source, each Text.Contains([Folder Path], FolderMarker)),
    OnlyExcel = Table.SelectRows(InFolder, each Text.EndsWith(Text.Lower([Name]), ".xlsx")),

    // Pull each file's first worksheet as a raw grid (headers not promoted).
    GetSheet = Table.AddColumn(OnlyExcel, "SheetData", each
        let
            wb     = Excel.Workbook([Content], false, true),
            sheets = Table.SelectRows(wb, each [Kind] = "Sheet")
        in
            try sheets{0}[Data] otherwise null),

    KeptSheets = Table.SelectRows(GetSheet, each [SheetData] <> null),
    Combined   = Table.Combine(KeptSheets[SheetData]),

    // Drop each file's repeated header row and any blank rows.
    DataRows = Table.SelectRows(Combined, each
        [Column1] <> "Report Date" and [Column1] <> null and [Column5] <> null),

    // Select the source columns we need, by position, in destination order.
    Picked = Table.SelectColumns(DataRows, {
        "Column1","Column2","Column5","Column4","Column6",
        "Column18","Column16","Column19","Column15","Column20",
        "Column10","Column11","Column12","Column13","Column14",
        "Column21","Column8","Column22","Column7"
    }),

    // Rename to the Raw Data EUR destination headers (A:E, then Q:AD).
    Renamed = Table.RenameColumns(Picked, {
        {"Column1",  "Date"},
        {"Column2",  "Country"},
        {"Column5",  "Customer"},
        {"Column4",  "Go Live"},
        {"Column6",  "Payment Term (days)"},
        {"Column18", "Total AR € (Live)"},
        {"Column16", "Gross Sales € (Live)"},
        {"Column19", "Overdue € (Live)"},
        {"Column15", "> 60 days € (Live)"},
        {"Column20", "> 90 days € (Live)"},
        {"Column10", "Total AR"},
        {"Column11", "Overdue"},
        {"Column12", "> 60 days"},
        {"Column13", "Gross Sales"},
        {"Column14", "> 90 days"},
        {"Column21", "Total UAC € (Live)"},
        {"Column8",  "Total UAC"},
        {"Column22", "Total Payments € (Live)"},
        {"Column7",  "Total Payments"}
    }),

    // Type the money columns as numbers. Date/Country/Customer/Go Live left as-is.
    Typed = Table.TransformColumnTypes(Renamed, {
        {"Payment Term (days)",       type number},
        {"Total AR € (Live)",         type number},
        {"Gross Sales € (Live)",      type number},
        {"Overdue € (Live)",          type number},
        {"> 60 days € (Live)",        type number},
        {"> 90 days € (Live)",        type number},
        {"Total AR",                  type number},
        {"Overdue",                   type number},
        {"> 60 days",                 type number},
        {"Gross Sales",               type number},
        {"> 90 days",                 type number},
        {"Total UAC € (Live)",        type number},
        {"Total UAC",                 type number},
        {"Total Payments € (Live)",   type number},
        {"Total Payments",            type number}
    })
in
    Typed
