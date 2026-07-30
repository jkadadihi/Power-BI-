// Power Query M: RawData, with missing days repaired from the source files
//
// WHAT THIS IS FOR
// The live RawData query reads the newest Daily_Performance_Report and stops.
// Each daily file is built by copying the previous one and appending that day,
// so if a run ever fails, that day is absent from every file built afterwards
// and no later refresh brings it back. The raw source files, however, are all
// still sitting in Cust_Sol Daily File. This query reads the same daily file as
// before and then fills any missing days straight from those sources.
//
// SAFE TO TEST ALONGSIDE THE LIVE QUERY. Load it to its own worksheet and
// compare: on a normal day it should return exactly the same rows as the live
// RawData. It should only ever return MORE rows, never fewer, and only for
// dates that are genuinely absent.
//
// FAST PATH
// Opening every source file on every refresh would be slow, so the repair only
// runs when a source file exists whose date is not already in the history. The
// check uses each file's Report Date, read from inside the file, so it does not
// depend on the source file naming (which we have already seen vary).
//
// SCOPE
// Only gaps INSIDE the existing history range are filled. Source files older
// than the first date in RawData are ignored, so this repairs holes rather than
// silently extending history backwards.
//
// SETUP
//   1. Data > Get Data > Blank Query > Advanced Editor, paste this in.
//   2. Needs the same SharePointSite_Url text parameter as the other queries.
//   3. Name it RawData_SelfHealing and Close & Load to a NEW worksheet.
//   4. Compare against the live RawData table before switching anything over.
let
    Site = SharePointSite_Url,
    AllFiles = SharePoint.Files(Site, [ApiVersion = 15]),

    // ---------------------------------------------------------------
    // 1. The history: newest Daily_Performance_Report, exactly as the
    //    live query finds it.
    // ---------------------------------------------------------------
    DailyFiles = Table.SelectRows(AllFiles,
        each Text.Contains([Folder Path], "2026 AMKAD Summaries/Daily", Comparer.OrdinalIgnoreCase)
             and Text.StartsWith([Name], "Daily_Performance_Report", Comparer.OrdinalIgnoreCase)
             and [Extension] = ".xlsm"),

    MonthMap = [January=1, February=2, March=3, April=4, May=5, June=6,
                July=7, August=8, September=9, October=10, November=11, December=12],
    WithFileDate = Table.AddColumn(DailyFiles, "FileDate", each
        try
            let
                middle  = Text.BetweenDelimiters([Name], "Report-", "AMKAD"),
                tokens  = List.Select(Text.Split(Text.Trim(middle), " "), each _ <> ""),
                mName   = Text.Proper(Text.Lower(tokens{0})),
                dayNum  = Number.From(tokens{1}),
                yearNum = Number.From(tokens{2})
            in
                #date(yearNum, Record.Field(MonthMap, mName), dayNum)
        otherwise null, type date),
    Ranked = Table.Sort(WithFileDate,
        {{"FileDate", Order.Descending}, {"Date modified", Order.Descending}}),
    LatestFile = if Table.RowCount(Ranked) = 0
        then error "No Daily_Performance_Report .xlsm files found in the 2026 Daily folder."
        else Ranked{0}[Content],
    History = Excel.Workbook(LatestFile, true){[Item = "RawData", Kind = "Table"]}[Data],

    // Dates already covered. Month may arrive as a date or as text, so coerce.
    AsDate = (v) => try Date.From(v) otherwise null,
    HistDates = List.Distinct(List.RemoveNulls(
        List.Transform(Table.Column(History, "Month"), each AsDate(_)))),
    EarliestHist = if List.IsEmpty(HistDates) then null else List.Min(HistDates),

    // ---------------------------------------------------------------
    // 2. The source files that could fill a gap.
    // ---------------------------------------------------------------
    SrcFiles = try Table.SelectRows(AllFiles,
        each Text.Contains([Folder Path], "Cust_Sol Daily File", Comparer.OrdinalIgnoreCase)
             and (Text.Lower([Extension]) = ".xlsx" or Text.Lower([Extension]) = ".xlsm"))
        otherwise #table({"Content", "Name"}, {}),

    // Read each source file's own Report Date (first data row, column 1). Every
    // row in a source file shares one report date, so one read settles it.
    SrcWithDate = Table.AddColumn(SrcFiles, "SrcDate", each
        try
            let
                sheet = Excel.Workbook([Content], null, true){0}[Data],
                body  = Table.SelectRows(sheet, each
                            [Column1] <> null and [Column1] <> "Report Date")
            in
                AsDate(Table.FirstValue(Table.FirstN(body, 1)))
        otherwise null, type date),

    // Only genuine gaps: dated, not already present, and not older than the
    // history we hold.
    MissingFiles = Table.SelectRows(SrcWithDate, each
        [SrcDate] <> null
        and not List.Contains(HistDates, [SrcDate])
        and (EarliestHist = null or [SrcDate] >= EarliestHist)),

    // ---------------------------------------------------------------
    // 3. Repair — only touched when something is actually missing.
    // ---------------------------------------------------------------
    Round0 = (v) => let n = try Number.From(v) otherwise null
                    in if n = null then null else Number.Round(n, 0),

    // Columns are taken by POSITION, not header name: the source has "Total
    // Payments" twice (native and EUR), and promoting headers would rename one
    // of them and quietly break the mapping.
    MapOne = (content) =>
        let
            sheet = Excel.Workbook(content, null, true){0}[Data],
            body  = Table.SelectRows(sheet, each
                        [Column1] <> null and [Column1] <> "Report Date" and [Column5] <> null),
            mapped = Table.FromRecords(List.Transform(Table.ToRecords(body), (r) => [
                #"Month"                    = Record.Field(r, "Column1"),
                #"Country"                  = Record.Field(r, "Column2"),
                #"Customer"                 = Record.Field(r, "Column5"),
                #"Go Live"                  = Record.Field(r, "Column4"),
                #"Payment Term (days)"      = Record.Field(r, "Column6"),
                #"Total AR € (Live)"        = Round0(Record.Field(r, "Column18")),
                #"Gross Sales € (Live)"     = Round0(Record.Field(r, "Column16")),
                #"Overdue € (Live)"         = Round0(Record.Field(r, "Column19")),
                #"> 60 days € (Live)"       = Round0(Record.Field(r, "Column15")),
                #"> 90 days € (Live)"       = Round0(Record.Field(r, "Column20")),
                #"Total AR"                 = Round0(Record.Field(r, "Column10")),
                #"Overdue"                  = Round0(Record.Field(r, "Column11")),
                #"> 60 days"                = Round0(Record.Field(r, "Column12")),
                #"Gross Sales"              = Round0(Record.Field(r, "Column13")),
                #"> 90 days"                = Round0(Record.Field(r, "Column14")),
                #"Total UAC € (Live)"       = Round0(Record.Field(r, "Column21")),
                #"Total UAC"                = Round0(Record.Field(r, "Column8")),
                #"Total Payments € (Live)"  = Round0(Record.Field(r, "Column22")),
                #"Total Payments"           = Round0(Record.Field(r, "Column7"))
            ]))
        in
            mapped,

    // BEST EFFORT. The repair opens one extra file per missing day, so it makes
    // far more network calls than the old single-file query did. A transient
    // SharePoint failure on any one of them must NOT take down RawData: a
    // backfill is a nice-to-have, the history is not. Each file is wrapped
    // individually, so one bad download costs that day and nothing else, and if
    // the whole repair fails the query still returns the history unchanged.
    MappedOrNull = List.Transform(MissingFiles[Content],
        each try MapOne(_) otherwise null),
    UsableRepairs = List.RemoveNulls(MappedOrNull),

    Repaired = if Table.RowCount(MissingFiles) = 0 or List.IsEmpty(UsableRepairs)
        then null
        else try Table.Combine(UsableRepairs) otherwise null,

    // ---------------------------------------------------------------
    // 4. Combine. Columns present in History but not in the repaired rows
    //    (the formula columns: TDSO, DSO, TDSO Gap, the % columns) come
    //    through as null, which is what they are for a backfilled day.
    // ---------------------------------------------------------------
    Result = if Repaired = null then History else Table.Combine({History, Repaired}),

    Sorted = Table.Sort(Result, {{"Month", Order.Ascending}, {"Country", Order.Ascending},
                                 {"Customer", Order.Ascending}})
in
    Sorted
