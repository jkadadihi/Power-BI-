// This is YOUR EXISTING "RawData" query from the current AMKAD input file,
// unchanged, PLUS one addition: after it finds yesterday's full history (same
// as it does today), it also reads today's newest Cust_Sol source file and
// stacks those rows on top. That combined result is what gets loaded — so the
// file this query lives in becomes, after refresh, the new full-history table.
// Power Automate's only job is then to save a copy of this file under today's
// date. No Office Script anywhere in this.
let
    Source = SharePoint.Files(
        "https://dpdhl.sharepoint.com/teams/EXP-USQIA-BS33384_AMKAD",
        [ApiVersion = 15]
    ),

    // ===== UNCHANGED: your existing logic that finds the latest daily file =====
    FilteredFiles = Table.SelectRows(
        Source,
        each Text.Contains([Folder Path], "2026 AMKAD Summaries/Daily", Comparer.OrdinalIgnoreCase)
            and Text.StartsWith([Name], "Daily_Performance_Report", Comparer.OrdinalIgnoreCase)
            and [Extension] = ".xlsm"
    ),
    MonthMap = [January=1, February=2, March=3, April=4, May=5, June=6,
                July=7, August=8, September=9, October=10, November=11, December=12],
    WithFileDate = Table.AddColumn(
        FilteredFiles, "FileDate", each
        try
            let
                middle  = Text.BetweenDelimiters([Name], "Report-", "AMKAD"),
                tokens  = List.Select(Text.Split(Text.Trim(middle), " "), each _ <> ""),
                mName   = Text.Proper(Text.Lower(tokens{0})),
                dayNum  = Number.From(tokens{1}),
                yearNum = Number.From(tokens{2}),
                monNum  = Record.Field(MonthMap, mName)
            in
                #date(yearNum, monNum, dayNum)
        otherwise null, type date
    ),
    Ranked = Table.Sort(WithFileDate, {{"FileDate", Order.Descending}, {"Date modified", Order.Descending}}),
    LatestFile = if Table.RowCount(Ranked) = 0
        then error "No Daily_Performance_Report .xlsm files found in the 2026 Daily folder."
        else Ranked{0}[Content],
    ImportedWorkbook = Excel.Workbook(LatestFile, true),
    RawDataTable = ImportedWorkbook{[Item = "RawData", Kind = "Table"]}[Data],
    // ===== end of your existing logic — RawDataTable = full history through yesterday =====

    // ===== NEW: pull today's rows from the raw Cust_Sol source and add them =====
    SrcFiles = Table.SelectRows(
        Source,
        each Text.Contains([Folder Path], "Cust_Sol Daily File", Comparer.OrdinalIgnoreCase)
            and Text.EndsWith(Text.Lower([Name]), ".xlsx")
    ),
    RankedSrc = Table.Sort(SrcFiles, {{"Date modified", Order.Descending}}),
    TodayRaw = if Table.RowCount(RankedSrc) = 0 then null else
        let
            wb = Excel.Workbook(RankedSrc{0}[Content], false, true),
            sh = Table.SelectRows(wb, each [Kind] = "Sheet")
        in try sh{0}[Data] otherwise null,

    TodayRows = if TodayRaw = null then #table(type table [], {}) else
        let
            noHeader = Table.SelectRows(TodayRaw, each
                [Column1] <> "Report Date" and [Column1] <> null and [Column5] <> null),
            picked = Table.SelectColumns(noHeader, {
                "Column1","Column2","Column5","Column4","Column6",
                "Column18","Column16","Column19","Column15","Column20",
                "Column10","Column11","Column12","Column13","Column14",
                "Column21","Column8","Column22","Column7"
            }),
            named = Table.RenameColumns(picked, {
                {"Column1","Date"},{"Column2","Country"},{"Column5","Customer"},
                {"Column4","Go Live"},{"Column6","Payment Term (days)"},
                {"Column18","Total AR € (Live)"},{"Column16","Gross Sales € (Live)"},
                {"Column19","Overdue € (Live)"},{"Column15","> 60 days € (Live)"},
                {"Column20","> 90 days € (Live)"},{"Column10","Total AR"},
                {"Column11","Overdue"},{"Column12","> 60 days"},{"Column13","Gross Sales"},
                {"Column14","> 90 days"},{"Column21","Total UAC € (Live)"},
                {"Column8","Total UAC"},{"Column22","Total Payments € (Live)"},
                {"Column7","Total Payments"}
            })
        in named,

    // Stack today's rows on top of the history the existing logic already found,
    // then drop any exact repeat of Date+Country+Customer (safe if this refreshes
    // more than once on the same day — last one in wins, nothing doubles up).
    Combined = Table.Combine({RawDataTable, TodayRows}),
    Result   = Table.Distinct(Combined, {"Date", "Country", "Customer"})
in
    Result
