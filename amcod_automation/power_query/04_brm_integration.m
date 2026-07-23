// Power Query M: BRM -> AMCOD Metrics
//
// Replaces AMCOD Stage 9 ("BRM Update"): Marcia opens the BRM file, copies
// Gross Sales / Open AR / 60+ / 90+ / DSO by hand, and pastes them into the
// AMCOD reporting tabs - a second major copy/paste step (Priority 6).
//
// EXPECTED SOURCE COLUMNS (BRM export, first sheet):
//   Country | Gross Sales (EUR) | Open AR (EUR) | 60+ (EUR) | 90+ (EUR) | DSO
// Adjust the column list in the two Table.SelectColumns/RenameColumns steps
// if Marcia's BRM export uses different headers - the shape of the output
// (one row per country per month, ready to feed the AMCOD tabs) doesn't
// need to change.
//
// BrmFile_Path and BrmMonth_Date are Power Query parameters: point
// BrmFile_Path at the current month's BRM export, and set BrmMonth_Date to
// the reporting month so this query's output can be appended to a running
// AMCOD BRM history table instead of overwriting last month's row.

let
    Source = Excel.Workbook(File.Contents(BrmFile_Path), null, true),
    RawSheet = Source{[Item = "BRM Export", Kind = "Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(RawSheet, [PromoteAllScalars = true]),

    Cols = Table.SelectColumns(PromotedHeaders, {
        "Country", "Gross Sales (EUR)", "Open AR (EUR)", "60+ (EUR)", "90+ (EUR)", "DSO"
    }),
    Renamed = Table.RenameColumns(Cols, {
        {"Country", "country"}, {"Gross Sales (EUR)", "gross_sales_eur"},
        {"Open AR (EUR)", "open_ar_eur"}, {"60+ (EUR)", "gt60_eur"},
        {"90+ (EUR)", "gt90_eur"}, {"DSO", "brm_dso"}
    }),

    RemovedBlanks = Table.SelectRows(Renamed, each [country] <> null and [country] <> ""),

    AddMonth = Table.AddColumn(RemovedBlanks, "month_date", each BrmMonth_Date, type date),

    Typed = Table.TransformColumnTypes(AddMonth, {
        {"country", type text},
        {"gross_sales_eur", Currency.Type}, {"open_ar_eur", Currency.Type},
        {"gt60_eur", Currency.Type}, {"gt90_eur", Currency.Type},
        {"brm_dso", type number}
    })
in
    Typed
