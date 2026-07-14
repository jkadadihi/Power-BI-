// ORIGINAL "RawData" query from the current AMKAD input file, exactly as
// pasted (from the file's Mashup XML), BEFORE any of the Cust Sol automation
// edits. Keep this — if the updated query in custsol_rawdata.m ever causes a
// problem in production, paste THIS back into Advanced Editor to revert to
// exactly what was running before.
let
    Source = SharePoint.Files(
        "https://dpdhl.sharepoint.com/teams/EXP-USQIA-BS33384_AMKAD",
        [ApiVersion = 15]
    ),
    FilteredFiles = Table.SelectRows(
        Source,
        each Text.Contains([Folder Path], "2026 AMKAD Summaries/Daily", Comparer.OrdinalIgnoreCase)
            and Text.StartsWith([Name], "Daily_Performance_Report", Comparer.OrdinalIgnoreCase)
            and [Extension] = ".xlsm"
    ),
    // Month-name lookup
    MonthMap = [January=1, February=2, March=3, April=4, May=5, June=6,
                July=7, August=8, September=9, October=10, November=11, December=12],
    // Parse "June 30 2026" out of "Daily_Performance_Report- June 30 2026 AMKAD"
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
    // Newest by the date IN THE FILENAME (authoritative), tie-break on modified
    Ranked = Table.Sort(
        WithFileDate,
        {{"FileDate", Order.Descending}, {"Date modified", Order.Descending}}
    ),
    LatestFile = if Table.RowCount(Ranked) = 0
        then error "No Daily_Performance_Report .xlsm files found in the 2026 Daily folder."
        else Ranked{0}[Content],
    ImportedWorkbook = Excel.Workbook(LatestFile, true),
    RawDataTable = ImportedWorkbook{[Item = "RawData", Kind = "Table"]}[Data]
in
    RawDataTable
