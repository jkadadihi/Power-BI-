// Power Query M: Load & Clean AR Performance Data
// Sheet: Raw Data EUR Onboarding
let
    FilePath = Excel_File_Path,
    Source = Excel.Workbook(File.Contents(FilePath), null, true),
    RawSheet = Source{[Item="Raw Data EUR Onboarding", Kind="Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(Table.Skip(RawSheet, 0), [PromoteAllScalars=true]),
    RemovedBlanks = Table.SelectRows(PromotedHeaders,
        each not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), {"", null}))),
    Cols = Table.SelectColumns(RemovedBlanks, {
        "Customer","Country","Month","DSO",
        "AR Amount (EUR)","Overdue Amount (EUR)","Overdue %",
        ">60 Days (EUR)",">60 Days %",">90 Days (EUR)",">90 Days %"
    }),
    Renamed = Table.RenameColumns(Cols, {
        {"Customer","customer"},{"Country","country"},{"Month","month_raw"},
        {"DSO","dso"},{"AR Amount (EUR)","ar_eur"},
        {"Overdue Amount (EUR)","overdue_eur"},{"Overdue %","overdue_pct"},
        {">60 Days (EUR)","gt60_eur"},{">60 Days %","gt60_pct"},
        {">90 Days (EUR)","gt90_eur"},{">90 Days %","gt90_pct"}
    }),
    Parsed = Table.AddColumn(Renamed, "month_date",
        each let r = Text.Trim([month_raw]),
                 d = try Date.FromText(r,[Format="MMM-yy",Culture="en-US"])
                     otherwise try Date.FromText(r,[Format="MMMM yyyy",Culture="en-US"])
                     otherwise null
        in d, type date),
    Typed = Table.TransformColumnTypes(Parsed, {
        {"customer",type text},{"country",type text},
        {"dso",type number},{"ar_eur",Currency.Type},
        {"overdue_eur",Currency.Type},{"overdue_pct",Percentage.Type},
        {"gt60_eur",Currency.Type},{"gt60_pct",Percentage.Type},
        {"gt90_eur",Currency.Type},{"gt90_pct",Percentage.Type}
    }),
    Final = Table.RemoveColumns(Typed, {"month_raw"})
in Final
