"""
generate_pbit.py
Generates a Power BI Template (.pbit) for the DHL AMKAD OTC/SPR report.

Usage:
    python generate_pbit.py
    -> writes  AMKAD_OTC_Report.pbit  in the current directory

Open the .pbit in Power BI Desktop. It will ask for the Excel file path once,
then load all tables, relationships, measures, and 4 report pages automatically.

Requirements:  Python 3.8+  (stdlib only, no pip installs needed)
"""

import json
import zipfile
import io
import os

M_FACT = r"""
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
""".strip()

M_DIM_DATE = r"""
let
    MinDate = List.Min(fact_ar_performance[month_date]),
    MaxDate = List.Max(fact_ar_performance[month_date]),
    Count   = Number.RoundUp(Duration.Days(MaxDate - MinDate) / 30) + 2,
    Dates   = List.Dates(MinDate, Count, #duration(30,0,0,0)),
    T       = Table.TransformColumnTypes(
                  Table.FromList(Dates, Splitter.SplitByNothing(), {"date"}),
                  {{"date", type date}}),
    Y   = Table.AddColumn(T,  "year",        each Date.Year([date]),               Int64.Type),
    M   = Table.AddColumn(Y,  "month_num",   each Date.Month([date]),              Int64.Type),
    MN  = Table.AddColumn(M,  "month_name",  each Date.ToText([date],"MMMM"),      type text),
    MS  = Table.AddColumn(MN, "month_short", each Date.ToText([date],"MMM-yy"),    type text),
    Q   = Table.AddColumn(MS, "quarter",     each "Q"&Text.From(Date.QuarterOfYear([date])), type text),
    SK  = Table.AddColumn(Q,  "sort_key",    each Date.Year([date])*100+Date.Month([date]), Int64.Type),
    Out = Table.Sort(SK, {{"date", Order.Ascending}})
in Out
""".strip()

M_PARAM = '"C:\\Users\\jp1lq1\\OneDrive - DPDHL\\Desktop\\BI AMKAD\\Monthly_Performance_Report- April 2026 AMKAD - All Customers.xlsm"'


def build_model_schema():
    measures_fact = [
        {"name": "AR Total", "expression": "SUM(fact_ar_performance[ar_eur])", "formatString": "#,##0.00"},
        {"name": "Overdue Amount", "expression": "SUM(fact_ar_performance[overdue_eur])", "formatString": "#,##0.00"},
        {"name": "GT60 Amount", "expression": "SUM(fact_ar_performance[gt60_eur])", "formatString": "#,##0.00"},
        {"name": "GT90 Amount", "expression": "SUM(fact_ar_performance[gt90_eur])", "formatString": "#,##0.00"},
        {"name": "Overdue %", "expression": "DIVIDE([Overdue Amount],[AR Total])", "formatString": "0.0%"},
        {"name": "GT60 %", "expression": "DIVIDE([GT60 Amount],[AR Total])", "formatString": "0.0%"},
        {"name": "GT90 %", "expression": "DIVIDE([GT90 Amount],[AR Total])", "formatString": "0.0%"},
        {"name": "Overdue Amount PM", "expression": "CALCULATE([Overdue Amount],PREVIOUSMONTH(dim_date[date]))", "formatString": "#,##0.00"},
        {"name": "Overdue % PM", "expression": "CALCULATE([Overdue %],PREVIOUSMONTH(dim_date[date]))", "formatString": "0.0%"},
        {"name": "GT90 Amount PM", "expression": "CALCULATE([GT90 Amount],PREVIOUSMONTH(dim_date[date]))", "formatString": "#,##0.00"},
        {"name": "Overdue MoM Change EUR", "expression": "[Overdue Amount]-[Overdue Amount PM]", "formatString": "#,##0.00"},
        {"name": "Overdue MoM Change Pct", "expression": "[Overdue %]-[Overdue % PM]", "formatString": "0.0%"},
        {"name": "GT90 MoM Change EUR", "expression": "[GT90 Amount]-[GT90 Amount PM]", "formatString": "#,##0.00"},
        {"name": "Overdue Impact Rank", "expression": "RANKX(ALL(fact_ar_performance[customer]),[Overdue Amount],,DESC,DENSE)", "formatString": "0"},
        {"name": "DSO", "expression": "VAR TotalAR = SUM(fact_ar_performance[ar_eur])\nRETURN DIVIDE(SUMX(fact_ar_performance,fact_ar_performance[dso]*fact_ar_performance[ar_eur]),TotalAR)", "formatString": "0.0"},
        {"name": "DSO PM", "expression": "CALCULATE([DSO],PREVIOUSMONTH(dim_date[date]))", "formatString": "0.0"},
        {"name": "DSO MoM Change", "expression": "[DSO]-[DSO PM]", "formatString": "0.0"},
        {"name": "DSO Impact", "expression": "[DSO]-[DSO PM]", "formatString": "0.0"},
        {"name": "DSO Impact Rank", "expression": "RANKX(ALL(fact_ar_performance[customer]),[DSO Impact],,DESC,DENSE)", "formatString": "0"},
        {"name": "DSO Trend Indicator", "expression": "VAR delta=[DSO MoM Change]\nRETURN SWITCH(TRUE(),delta>2,\"Worsening\",delta<-2,\"Improving\",\"Stable\")"},
        {"name": "Payments Collected", "expression": "SUM(fact_ar_performance[payments_collected_eur])", "formatString": "#,##0.00"},
        {"name": "Payments Collected PM", "expression": "CALCULATE([Payments Collected],PREVIOUSMONTH(dim_date[date]))", "formatString": "#,##0.00"},
        {"name": "Payments MoM Change", "expression": "[Payments Collected]-[Payments Collected PM]", "formatString": "#,##0.00"},
        {"name": "Payments MoM Change %", "expression": "DIVIDE([Payments MoM Change],[Payments Collected PM])", "formatString": "0.0%"},
        {"name": "Collection Rate", "expression": "DIVIDE([Payments Collected],[Overdue Amount PM])", "formatString": "0.0%"},
    ]
    fact_columns = [
        {"name": "customer", "dataType": "string", "sourceColumn": "customer"},
        {"name": "country", "dataType": "string", "sourceColumn": "country"},
        {"name": "month_date", "dataType": "dateTime", "sourceColumn": "month_date", "formatString": "MMM-yy"},
        {"name": "dso", "dataType": "double", "sourceColumn": "dso"},
        {"name": "ar_eur", "dataType": "decimal", "sourceColumn": "ar_eur", "formatString": "#,##0.00"},
        {"name": "overdue_eur", "dataType": "decimal", "sourceColumn": "overdue_eur", "formatString": "#,##0.00"},
        {"name": "overdue_pct", "dataType": "double", "sourceColumn": "overdue_pct", "formatString": "0.0%"},
        {"name": "gt60_eur", "dataType": "decimal", "sourceColumn": "gt60_eur", "formatString": "#,##0.00"},
        {"name": "gt60_pct", "dataType": "double", "sourceColumn": "gt60_pct", "formatString": "0.0%"},
        {"name": "gt90_eur", "dataType": "decimal", "sourceColumn": "gt90_eur", "formatString": "#,##0.00"},
        {"name": "gt90_pct", "dataType": "double", "sourceColumn": "gt90_pct", "formatString": "0.0%"},
        {"name": "payments_collected_eur", "dataType": "decimal", "sourceColumn": "payments_collected_eur", "formatString": "#,##0.00"},
    ]
    date_columns = [
        {"name": "date", "dataType": "dateTime", "sourceColumn": "date", "formatString": "MMM-yy", "isKey": True},
        {"name": "year", "dataType": "int64", "sourceColumn": "year"},
        {"name": "month_num", "dataType": "int64", "sourceColumn": "month_num"},
        {"name": "month_name", "dataType": "string", "sourceColumn": "month_name"},
        {"name": "month_short", "dataType": "string", "sourceColumn": "month_short"},
        {"name": "quarter", "dataType": "string", "sourceColumn": "quarter"},
        {"name": "sort_key", "dataType": "int64", "sourceColumn": "sort_key"},
    ]
    return {
        "name": "AMKAD_OTC",
        "compatibilityLevel": 1550,
        "model": {
            "culture": "en-US",
            "dataAccessOptions": {"legacyRedirects": True, "returnErrorValuesAsNull": True},
            "defaultPowerBIDataSourceVersion": "powerBI_V3",
            "expressions": [{"name": "Excel_File_Path", "kind": "m", "expression": M_PARAM}],
            "tables": [
                {"name": "fact_ar_performance", "columns": fact_columns, "measures": measures_fact,
                 "partitions": [{"name": "fact_ar_performance", "mode": "import", "source": {"type": "m", "expression": M_FACT}}]},
                {"name": "dim_date", "columns": date_columns,
                 "partitions": [{"name": "dim_date", "mode": "import", "source": {"type": "m", "expression": M_DIM_DATE}}]}
            ],
            "relationships": [{
                "name": "fact_to_dim_date",
                "fromTable": "fact_ar_performance", "fromColumn": "month_date",
                "toTable": "dim_date", "toColumn": "date",
                "crossFilteringBehavior": "oneDirection"
            }],
            "annotations": [
                {"name": "PBIDesktopVersion", "value": "2.130"},
                {"name": "PBI_QueryOrder", "value": '["fact_ar_performance","dim_date"]'}
            ]
        }
    }


def kpi_card(name, x, y, w, h, tbl, measure):
    return {
        "id": abs(hash(name)) % 1000000, "x": x, "y": y, "z": 0, "width": w, "height": h,
        "config": json.dumps({"name": name, "singleVisual": {"visualType": "card",
            "projections": {"Values": [{"queryRef": f"{tbl}.{measure}", "active": True}]},
            "prototypeQuery": {"Select": [{"Measure": {"Expression": {"SourceRef": {"Source": "f"}}, "Property": measure}, "Name": f"{tbl}.{measure}"}], "From": [{"Name": "f", "Entity": tbl, "Type": 0}]}}})
    }


def table_visual(name, x, y, w, h, columns):
    select, src, vals = [], [], []
    for i, (tbl, col, col_name) in enumerate(columns):
        alias = chr(ord('a') + i)
        src.append({"Name": alias, "Entity": tbl, "Type": 0})
        vals.append({"queryRef": f"{tbl}.{col_name}"})
        if col.startswith("["):
            select.append({"Measure": {"Expression": {"SourceRef": {"Source": alias}}, "Property": col_name}, "Name": f"{tbl}.{col_name}"})
        else:
            select.append({"Column": {"Expression": {"SourceRef": {"Source": alias}}, "Property": col_name}, "Name": f"{tbl}.{col_name}"})
    return {"id": abs(hash(name)) % 1000000, "x": x, "y": y, "z": 0, "width": w, "height": h,
            "config": json.dumps({"name": name, "singleVisual": {"visualType": "tableEx", "projections": {"Values": vals}, "prototypeQuery": {"Select": select, "From": src}}})}


def line_chart(name, x, y, w, h, x_tbl, x_col, y_tbl, y_measure):
    return {"id": abs(hash(name)) % 1000000, "x": x, "y": y, "z": 0, "width": w, "height": h,
            "config": json.dumps({"name": name, "singleVisual": {"visualType": "lineChart",
                "projections": {"Category": [{"queryRef": f"{x_tbl}.{x_col}"}], "Y": [{"queryRef": f"{y_tbl}.{y_measure}", "active": True}]},
                "prototypeQuery": {"Select": [
                    {"Column": {"Expression": {"SourceRef": {"Source": "d"}}, "Property": x_col}, "Name": f"{x_tbl}.{x_col}"},
                    {"Measure": {"Expression": {"SourceRef": {"Source": "f"}}, "Property": y_measure}, "Name": f"{y_tbl}.{y_measure}"}
                ], "From": [{"Name": "d", "Entity": x_tbl, "Type": 0}, {"Name": "f", "Entity": y_tbl, "Type": 0}]}}})}


def build_report_layout():
    W, cw, ch, g = 1280, 250, 110, 20
    p1 = [
        kpi_card("dso_card", g, g, cw, ch, "fact_ar_performance", "DSO"),
        kpi_card("ar_card", g*2+cw, g, cw, ch, "fact_ar_performance", "AR Total"),
        kpi_card("dso_chg", g*3+cw*2, g, cw, ch, "fact_ar_performance", "DSO MoM Change"),
        kpi_card("dso_trend", g*4+cw*3, g, cw, ch, "fact_ar_performance", "DSO Trend Indicator"),
        line_chart("dso_line", g, ch+g*3, W-g*2, 260, "dim_date", "month_short", "fact_ar_performance", "DSO"),
        table_visual("dso_top10", g, ch+260+g*5, W-g*2, 230, [
            ("fact_ar_performance", "customer", "customer"),
            ("fact_ar_performance", "[DSO]", "DSO"),
            ("fact_ar_performance", "[DSO MoM Change]", "DSO MoM Change"),
            ("fact_ar_performance", "[DSO Impact]", "DSO Impact"),
        ]),
    ]
    p2 = [
        kpi_card("pay_card", g, g, cw, ch, "fact_ar_performance", "Payments Collected"),
        kpi_card("pay_chg", g*2+cw, g, cw, ch, "fact_ar_performance", "Payments MoM Change %"),
        kpi_card("coll_rate", g*3+cw*2, g, cw, ch, "fact_ar_performance", "Collection Rate"),
        line_chart("pay_line", g, ch+g*3, W-g*2, 350, "dim_date", "month_short", "fact_ar_performance", "Payments Collected"),
    ]
    p3 = [
        kpi_card("ovd_pct", g, g, cw, ch, "fact_ar_performance", "Overdue %"),
        kpi_card("gt60_pct", g*2+cw, g, cw, ch, "fact_ar_performance", "GT60 %"),
        kpi_card("gt90_pct", g*3+cw*2, g, cw, ch, "fact_ar_performance", "GT90 %"),
        kpi_card("ovd_eur", g*4+cw*3, g, cw, ch, "fact_ar_performance", "Overdue Amount"),
        table_visual("aging_tbl", g, ch+g*3, W-g*2, 300, [
            ("fact_ar_performance", "customer", "customer"),
            ("fact_ar_performance", "[Overdue Amount]", "Overdue Amount"),
            ("fact_ar_performance", "[Overdue %]", "Overdue %"),
            ("fact_ar_performance", "[GT60 Amount]", "GT60 Amount"),
            ("fact_ar_performance", "[GT90 Amount]", "GT90 Amount"),
        ]),
    ]
    p4 = [
        kpi_card("e_dso", g, g, cw, ch, "fact_ar_performance", "DSO"),
        kpi_card("e_dso_chg", g*2+cw, g, cw, ch, "fact_ar_performance", "DSO MoM Change"),
        kpi_card("e_ovd", g*3+cw*2, g, cw, ch, "fact_ar_performance", "Overdue %"),
        kpi_card("e_gt90", g*4+cw*3, g, cw, ch, "fact_ar_performance", "GT90 %"),
        table_visual("exp_kpi", g, ch+g*3, W//2-g*2, 300, [
            ("fact_ar_performance", "customer", "customer"),
            ("fact_ar_performance", "[DSO]", "DSO"),
            ("fact_ar_performance", "[DSO MoM Change]", "DSO MoM Change"),
            ("fact_ar_performance", "[Overdue %]", "Overdue %"),
            ("fact_ar_performance", "[GT90 %]", "GT90 %"),
        ]),
        table_visual("exp_contrib", W//2+g, ch+g*3, W//2-g*2, 300, [
            ("fact_ar_performance", "customer", "customer"),
            ("fact_ar_performance", "[DSO Impact]", "DSO Impact"),
            ("fact_ar_performance", "[Overdue MoM Change EUR]", "Overdue MoM Change EUR"),
            ("fact_ar_performance", "[GT90 MoM Change EUR]", "GT90 MoM Change EUR"),
        ]),
    ]
    return {
        "id": 0, "resourcePackages": [],
        "sections": [
            {"name": "dso_page",   "displayName": "DSO Performance", "displayOption": 1, "height": 720, "width": 1280, "visualContainers": p1, "ordinal": 0},
            {"name": "cash_page",  "displayName": "Cash Collection",  "displayOption": 1, "height": 720, "width": 1280, "visualContainers": p2, "ordinal": 1},
            {"name": "aging_page", "displayName": "Overdue & Aging",  "displayOption": 1, "height": 720, "width": 1280, "visualContainers": p3, "ordinal": 2},
            {"name": "ppt_page",   "displayName": "PPT Export",       "displayOption": 1, "height": 720, "width": 1280, "visualContainers": p4, "ordinal": 3},
        ],
        "config": json.dumps({"version": "5.49", "activeSectionIndex": 0, "defaultDrillFilterOtherVisuals": True}),
        "filters": "[]", "layoutOptimization": 0
    }


CONTENT_TYPES = """<?xml version=\"1.0\" encoding=\"utf-8\"?>
<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">
  <Default Extension=\"json\" ContentType=\"application/json\" />
  <Override PartName=\"/DataModelSchema\" ContentType=\"application/json\" />
  <Override PartName=\"/DiagramLayout\"   ContentType=\"application/json\" />
  <Override PartName=\"/Report/Layout\"   ContentType=\"application/json\" />
  <Override PartName=\"/Metadata\"        ContentType=\"application/json\" />
  <Override PartName=\"/Version\"         ContentType=\"application/octet-stream\" />
  <Override PartName=\"/SecurityBindings\" ContentType=\"application/octet-stream\" />
</Types>"""

METADATA = json.dumps({"version": "4.0", "createdFromTemplate": False, "description": "DHL AMKAD OTC/SPR AR Reporting Template"})
DIAGRAM_LAYOUT = json.dumps({"version": 1, "diagramLayout": {"nodeLayouts": {
    "fact_ar_performance": {"x": 400, "y": 200, "width": 280, "height": 400},
    "dim_date": {"x": 50, "y": 100, "width": 200, "height": 200}
}}})


def write_pbit(output_path):
    schema = build_model_schema()
    layout = build_report_layout()
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("DataModelSchema",     json.dumps(schema, ensure_ascii=False, indent=2))
        zf.writestr("DiagramLayout",       DIAGRAM_LAYOUT)
        zf.writestr("Metadata",            METADATA)
        zf.writestr("Report/Layout",       json.dumps(layout, ensure_ascii=False, indent=2))
        zf.writestr("SecurityBindings",    b"")
        zf.writestr("Version",             b"4.0")
    with open(output_path, "wb") as f:
        f.write(buf.getvalue())
    print(f"Written: {output_path}  ({os.path.getsize(output_path)/1024:.1f} KB)")
    print("Open in Power BI Desktop and set Excel_File_Path when prompted.")


if __name__ == "__main__":
    write_pbit(os.path.join(os.path.dirname(os.path.abspath(__file__)), "AMKAD_OTC_Report.pbit"))
