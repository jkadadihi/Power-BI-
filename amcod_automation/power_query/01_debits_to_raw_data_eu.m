// Power Query M: Debits Workbook(s) -> Raw Data EU
//
// Replaces AMCOD Stage 3 ("Update Raw Data EU"), the single biggest manual
// step in the process: open Debits workbook, open Monthly workbook, scroll
// to bottom of Raw Data EU, copy a block, paste values, fix formatting,
// update the month indicator - repeated for every section/country.
//
// How it works: point DebitsFolderPath at the folder where each month's
// Debits-Only file lands (Stage 2 - "Get Month-End Data"). Every file in
// the folder is combined automatically, so a new month's file just needs
// to be dropped in the folder - no copy/paste, and nothing to remember to
// "scroll to the bottom" for.
//
// EXPECTED SOURCE COLUMNS (per Debits-Only workbook, first sheet):
//   Country | Customer | Invoice Number | Invoice Date | Due Date
//   Amount (EUR) | Debit/Credit | Reference
// If Marcia's actual Debits export uses different headers, only the
// Table.RenameColumns step below needs to change - everything downstream
// keys off the renamed (lowercase) column names.
//
// EXPECTED FILENAME CONVENTION: "Debits Only - <Month> <Year>.xlsx"
// e.g. "Debits Only - June 2026.xlsx". The month/year used for the
// "month" indicator column is parsed from the filename, not the file's
// modified date, so it stays correct even if a file is re-saved later.

let
    DebitsFolderPath = DebitsFolder_Path,

    Source = Folder.Files(DebitsFolderPath),
    ExcelFiles = Table.SelectRows(Source, each Text.EndsWith([Extension], ".xlsx") or Text.EndsWith([Extension], ".xlsm")),

    // "Debits Only - June 2026.xlsx" -> "June 2026": take everything after the
    // last " - " in the filename, then strip the file extension.
    AddMonthRaw = Table.AddColumn(ExcelFiles, "month_raw",
        each Text.Trim(Text.BeforeDelimiter(
                Text.AfterDelimiter([Name], " - ", {0, RelativePosition.FromEnd}),
                ".")),
        type text),

    AddData = Table.AddColumn(AddMonthRaw, "Data", each Excel.Workbook([Content], null, true)),
    ExpandSheets = Table.ExpandTableColumn(AddData, "Data", {"Item", "Kind", "Data"}, {"Item", "Kind", "Data"}),
    FirstSheetOnly = Table.SelectRows(ExpandSheets, each [Kind] = "Sheet" and [Item] <> "Cover" and [Item] <> "ReadMe"),

    ExpandRows = Table.ExpandTableColumn(FirstSheetOnly, "Data",
        {"Country", "Customer", "Invoice Number", "Invoice Date", "Due Date", "Amount (EUR)", "Debit/Credit", "Reference"},
        {"Country", "Customer", "Invoice Number", "Invoice Date", "Due Date", "Amount (EUR)", "Debit/Credit", "Reference"}),

    RemoveHeaderNoise = Table.SelectRows(ExpandRows, each [Country] <> null and [Country] <> "Country"),

    Renamed = Table.RenameColumns(RemoveHeaderNoise, {
        {"Country", "country"}, {"Customer", "customer"},
        {"Invoice Number", "invoice_number"}, {"Invoice Date", "invoice_date"},
        {"Due Date", "due_date"}, {"Amount (EUR)", "amount_eur"},
        {"Debit/Credit", "debit_credit"}, {"Reference", "reference"}
    }),

    AddMonthDate = Table.AddColumn(Renamed, "month_date",
        each try Date.FromText([month_raw], [Format = "MMMM yyyy", Culture = "en-US"]) otherwise null, type date),

    Typed = Table.TransformColumnTypes(AddMonthDate, {
        {"country", type text}, {"customer", type text},
        {"invoice_number", type text},
        {"invoice_date", type date}, {"due_date", type date},
        {"amount_eur", Currency.Type}, {"debit_credit", type text},
        {"reference", type text}
    }),

    RemovedHelperCols = Table.RemoveColumns(Typed, {"month_raw", "Item", "Kind", "Name", "Extension", "Date accessed", "Date modified", "Date created", "Attributes", "Folder Path", "Content"}, MissingField.Ignore),

    Sorted = Table.Sort(RemovedHelperCols, {{"month_date", Order.Ascending}, {"country", Order.Ascending}})
in
    Sorted
