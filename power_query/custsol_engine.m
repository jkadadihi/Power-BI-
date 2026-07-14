// Power Query (M) — Daily engine: "yesterday's daily file + today's source"
// ---------------------------------------------------------------------------
// This is the query the team asked for, literally: start from yesterday's
// Daily_Performance_Report (which already holds the FULL running history in its
// RawData table) and add today's new rows from the raw Cust_Sol source file.
//
// Why this works even though Power Query "replaces, never appends": the history
// is READ from a SEPARATE file (yesterday's daily), not accumulated in this
// query's own output. So each refresh rebuilds  (yesterday's history) + (today)
// = today's full history. Nothing old is lost, and the Cust_Sol source folder
// only needs to contain today's file — the older history rides along inside the
// daily-file chain.
//
// SETUP: paste into a blank query in the ENGINE workbook (.xlsm). Name the
// query "RawData", Close & Load to a Table, then set the Table Name to exactly
// "RawData" in Table Design. Power Automate then copies the refreshed engine to
// "Daily_Performance_Report- <date>  AMKAD.xlsm" in the Daily folder.
let
    Site     = "https://dpdhl.sharepoint.com/teams/EXP-USQIA-BS33384_AMKAD",
    AllFiles = SharePoint.Files(Site, [ApiVersion = 15]),

    // ===== Source 1: newest existing Daily_Performance_Report = full history so far =====
    DailyFiles = Table.SelectRows(AllFiles, each
        Text.Contains([Folder Path], "2026 AMKAD Summaries/Daily", Comparer.OrdinalIgnoreCase)
        and Text.StartsWith([Name], "Daily_Performance_Report", Comparer.OrdinalIgnoreCase)
        and Text.EndsWith(Text.Lower([Name]), ".xlsm")),

    // Parse the date out of "Daily_Performance_Report- June 30 2026  AMKAD.xlsm"
    // (authoritative — same logic as the existing downstream query).
    MonthMap = [January=1, February=2, March=3, April=4, May=5, June=6,
                July=7, August=8, September=9, October=10, November=11, December=12],
    WithDate = Table.AddColumn(DailyFiles, "FileDate", each
        try
            let
                middle = Text.BetweenDelimiters([Name], "Report-", "AMKAD"),
                toks   = List.Select(Text.Split(Text.Trim(middle), " "), each _ <> ""),
                mon    = Record.Field(MonthMap, Text.Proper(Text.Lower(toks{0}))),
                dd     = Number.From(toks{1}),
                yy     = Number.From(toks{2})
            in
                #date(yy, mon, dd)
        otherwise null, type date),
    RankedDaily = Table.Sort(WithDate, {{"FileDate", Order.Descending}, {"Date modified", Order.Descending}}),
    History = if Table.RowCount(RankedDaily) = 0
              then #table(type table [], {})
              else Excel.Workbook(RankedDaily{0}[Content]){[Item = "RawData", Kind = "Table"]}[Data],

    // Normalize the history's Date column to text so it stacks/dedupes cleanly
    // against today's rows (which come in as text).
    HistoryTxt = if Table.HasColumns(History, "Date")
                 then Table.TransformColumns(History, {{"Date", each Text.From(_), type text}})
                 else History,

    // ===== Source 2: newest Cust_Sol raw file = today's new rows =====
    SrcFiles  = Table.SelectRows(AllFiles, each
        Text.Contains([Folder Path], "Cust_Sol Daily File", Comparer.OrdinalIgnoreCase)
        and Text.EndsWith(Text.Lower([Name]), ".xlsx")),
    RankedSrc = Table.Sort(SrcFiles, {{"Date modified", Order.Descending}}),
    TodaySheet = if Table.RowCount(RankedSrc) = 0 then null else
        let
            wb = Excel.Workbook(RankedSrc{0}[Content], false, true),
            sh = Table.SelectRows(wb, each [Kind] = "Sheet")
        in try sh{0}[Data] otherwise null,

    TodayRows = if TodaySheet = null then #table(type table [], {}) else
        let
            noHdr  = Table.SelectRows(TodaySheet, each
                        [Column1] <> "Report Date" and [Column1] <> null and [Column5] <> null),
            picked = Table.SelectColumns(noHdr, {
                        "Column1","Column2","Column5","Column4","Column6",
                        "Column18","Column16","Column19","Column15","Column20",
                        "Column10","Column11","Column12","Column13","Column14",
                        "Column21","Column8","Column22","Column7"}),
            named  = Table.RenameColumns(picked, {
                        {"Column1","Date"},{"Column2","Country"},{"Column5","Customer"},
                        {"Column4","Go Live"},{"Column6","Payment Term (days)"},
                        {"Column18","Total AR € (Live)"},{"Column16","Gross Sales € (Live)"},
                        {"Column19","Overdue € (Live)"},{"Column15","> 60 days € (Live)"},
                        {"Column20","> 90 days € (Live)"},{"Column10","Total AR"},
                        {"Column11","Overdue"},{"Column12","> 60 days"},{"Column13","Gross Sales"},
                        {"Column14","> 90 days"},{"Column21","Total UAC € (Live)"},
                        {"Column8","Total UAC"},{"Column22","Total Payments € (Live)"},
                        {"Column7","Total Payments"}}),
            asText = Table.TransformColumns(named, {{"Date", each Text.From(_), type text}})
        in asText,

    // ===== Stack history + today, dedupe on Date+Country+Customer (re-run safe) =====
    Combined = Table.Combine({HistoryTxt, TodayRows}),
    Result   = Table.Distinct(Combined, {"Date", "Country", "Customer"})
in
    Result
