// Power Query M Script: Load & Clean AR Performance Data
// Source: Monthly_Performance_Report- April 2026 AMKAD - All Customers.xlsm
// Sheet: Raw Data EUR Onboarding

let
    // ── 1. Connect to Excel workbook ──────────────────────────────────────────
    Source = Excel.Workbook(
        File.Contents("C:\Reports\Monthly_Performance_Report- April 2026 AMKAD - All Customers.xlsm"),
        null, true
    ),

    // ── 2. Select the correct sheet ──────────────────────────────────────────
    RawSheet = Source{[Item="Raw Data EUR Onboarding", Kind="Sheet"]}[Data],

    // ── 3. Promote first real header row ─────────────────────────────────────
    //      (adjust row index if there are title rows above the header)
    PromotedHeaders = Table.PromoteHeaders(
        Table.Skip(RawSheet, 0),   // change 0 → N if junk rows exist above header
        [PromoteAllScalars=true]
    ),

    // ── 4. Remove fully empty rows ───────────────────────────────────────────
    RemovedBlanks = Table.SelectRows(
        PromotedHeaders,
        each not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), {"", null}))
    ),

    // ── 5. Keep only the columns we need ─────────────────────────────────────
    //      Rename to standardised snake_case names
    SelectedColumns = Table.SelectColumns(
        RemovedBlanks,
        {
            "Customer",
            "Country",
            "Month",
            "DSO",
            "AR Amount (EUR)",
            "Overdue Amount (EUR)",
            "Overdue %",
            ">60 Days (EUR)",
            ">60 Days %",
            ">90 Days (EUR)",
            ">90 Days %"
        }
    ),

    RenamedColumns = Table.RenameColumns(
        SelectedColumns,
        {
            {"Customer",            "customer"},
            {"Country",             "country"},
            {"Month",               "month_raw"},
            {"DSO",                 "dso"},
            {"AR Amount (EUR)",     "ar_eur"},
            {"Overdue Amount (EUR)","overdue_eur"},
            {"Overdue %",           "overdue_pct"},
            {">60 Days (EUR)",      "gt60_eur"},
            {">60 Days %",          "gt60_pct"},
            {">90 Days (EUR)",      "gt90_eur"},
            {">90 Days %",          "gt90_pct"}
        }
    ),

    // ── 6. Parse month to a proper Date (first day of month) ─────────────────
    //      Assumes format like "Apr-26" or "April 2026" — adjust if different
    ParsedMonth = Table.AddColumn(
        RenamedColumns,
        "month_date",
        each
            let
                raw = Text.Trim([month_raw]),
                // Try "MMM-YY" format first, then "MMMM YYYY"
                parsed = try Date.FromText(raw, [Format="MMM-yy", Culture="en-US"])
                         otherwise try Date.FromText(raw, [Format="MMMM yyyy", Culture="en-US"])
                         otherwise null
            in parsed,
        type date
    ),

    // ── 7. Set correct column types ──────────────────────────────────────────
    TypedTable = Table.TransformColumnTypes(
        ParsedMonth,
        {
            {"customer",    type text},
            {"country",     type text},
            {"dso",         type number},
            {"ar_eur",      Currency.Type},
            {"overdue_eur", Currency.Type},
            {"overdue_pct", Percentage.Type},
            {"gt60_eur",    Currency.Type},
            {"gt60_pct",    Percentage.Type},
            {"gt90_eur",    Currency.Type},
            {"gt90_pct",    Percentage.Type}
        }
    ),

    // ── 8. Drop the raw month column ─────────────────────────────────────────
    FinalTable = Table.RemoveColumns(TypedTable, {"month_raw"})

in
    FinalTable
