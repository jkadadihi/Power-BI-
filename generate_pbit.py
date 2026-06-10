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
import struct

# ─────────────────────────────────────────────────────────────────────────────
#  POWER QUERY M  (embedded as named expressions in the model)
# ─────────────────────────────────────────────────────────────────────────────

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

M_PARAM = r"""
"C:\Reports\Monthly_Performance_Report- April 2026 AMKAD - All Customers.xlsm"
""".strip()

# ─────────────────────────────────────────────────────────────────────────────
#  DATA MODEL SCHEMA  (Tabular Object Model JSON)
# ─────────────────────────────────────────────────────────────────────────────

def build_model_schema():
    measures_fact = [
        # ── Overdue & Aging ──────────────────────────────────────────────────
        {
            "name": "AR Total",
            "expression": "SUM(fact_ar_performance[ar_eur])",
            "formatString": "#,##0.00 €"
        },
        {
            "name": "Overdue Amount",
            "expression": "SUM(fact_ar_performance[overdue_eur])",
            "formatString": "#,##0.00 €"
        },
        {
            "name": "GT60 Amount",
            "expression": "SUM(fact_ar_performance[gt60_eur])",
            "formatString": "#,##0.00 €"
        },
        {
            "name": "GT90 Amount",
            "expression": "SUM(fact_ar_performance[gt90_eur])",
            "formatString": "#,##0.00 €"
        },
        {
            "name": "Overdue %",
            "expression": "DIVIDE([Overdue Amount],[AR Total])",
            "formatString": "0.0%"
        },
        {
            "name": "GT60 %",
            "expression": "DIVIDE([GT60 Amount],[AR Total])",
            "formatString": "0.0%"
        },
        {
            "name": "GT90 %",
            "expression": "DIVIDE([GT90 Amount],[AR Total])",
            "formatString": "0.0%"
        },
        {
            "name": "Overdue Amount PM",
            "expression": "CALCULATE([Overdue Amount],PREVIOUSMONTH(dim_date[date]))",
            "formatString": "#,##0.00 €"
        },
        {
            "name": "Overdue % PM",
            "expression": "CALCULATE([Overdue %],PREVIOUSMONTH(dim_date[date]))",
            "formatString": "0.0%"
        },
        {
            "name": "GT90 Amount PM",
            "expression": "CALCULATE([GT90 Amount],PREVIOUSMONTH(dim_date[date]))",
            "formatString": "#,##0.00 €"
        },
        {
            "name": "Overdue MoM Change EUR",
            "expression": "[Overdue Amount]-[Overdue Amount PM]",
            "formatString": "#,##0.00 €"
        },
        {
            "name": "Overdue MoM Change Pct",
            "expression": "[Overdue %]-[Overdue % PM]",
            "formatString": "0.0%"
        },
        {
            "name": "GT90 MoM Change EUR",
            "expression": "[GT90 Amount]-[GT90 Amount PM]",
            "formatString": "#,##0.00 €"
        },
        {
            "name": "Overdue Impact Rank",
            "expression": "RANKX(ALL(fact_ar_performance[customer]),[Overdue Amount],,DESC,DENSE)",
            "formatString": "0"
        },
        {
            "name": "Overdue % Label",
            "expression": (
                'FORMAT([Overdue %],"0.0%")'
                '& IF(NOT ISBLANK([Overdue % PM]),'
                '" (" & IF([Overdue MoM Change Pct]>=0,"+","")'
                '& FORMAT([Overdue MoM Change Pct],"0.0pp") & " MoM)","")'
            )
        },
        {
            "name": "GT90 % Label",
            "expression": (
                'FORMAT([GT90 %],"0.0%")'
                '& IF(NOT ISBLANK([GT90 Amount PM]),'
                '" (" & IF([GT90 MoM Change EUR]>=0,"+","")'
                '& FORMAT([GT90 MoM Change EUR]/1000000,"#,##0.0M €") & " MoM)","")'
            )
        },
        # ── DSO ──────────────────────────────────────────────────────────────
        {
            "name": "DSO",
            "expression": (
                "VAR TotalAR = SUM(fact_ar_performance[ar_eur])\n"
                "RETURN DIVIDE(\n"
                "    SUMX(fact_ar_performance,fact_ar_performance[dso]*fact_ar_performance[ar_eur]),\n"
                "    TotalAR)"
            ),
            "formatString": "0.0"
        },
        {
            "name": "DSO PM",
            "expression": "CALCULATE([DSO],PREVIOUSMONTH(dim_date[date]))",
            "formatString": "0.0"
        },
        {
            "name": "DSO MoM Change",
            "expression": "[DSO]-[DSO PM]",
            "formatString": "0.0"
        },
        {
            "name": "DSO MoM Change Label",
            "expression": (
                "VAR delta=[DSO MoM Change]\n"
                'RETURN IF(ISBLANK(delta),BLANK(),IF(delta>=0,"+"&FORMAT(delta,"0.0"),FORMAT(delta,"0.0")))'
            )
        },
        {
            "name": "DSO Trend Indicator",
            "expression": (
                "VAR delta=[DSO MoM Change]\n"
                'RETURN SWITCH(TRUE(),delta>2,"▲ Worsening",delta<-2,"▼ Improving","→ Stable")'
            )
        },
        {
            "name": "DSO Impact",
            "expression": "[DSO]-[DSO PM]",
            "formatString": "0.0"
        },
        {
            "name": "DSO Impact Rank",
            "expression": "RANKX(ALL(fact_ar_performance[customer]),[DSO Impact],,DESC,DENSE)",
            "formatString": "0"
        },
        # ── Cash Collection ──────────────────────────────────────────────────
        {
            "name": "Payments Collected",
            "expression": "SUM(fact_ar_performance[payments_collected_eur])",
            "formatString": "#,##0.00 €",
            "description": "Requires payments_collected_eur column in source data"
        },
        {
            "name": "Payments Collected PM",
            "expression": "CALCULATE([Payments Collected],PREVIOUSMONTH(dim_date[date]))",
            "formatString": "#,##0.00 €"
        },
        {
            "name": "Payments MoM Change",
            "expression": "[Payments Collected]-[Payments Collected PM]",
            "formatString": "#,##0.00 €"
        },
        {
            "name": "Payments MoM Change %",
            "expression": "DIVIDE([Payments MoM Change],[Payments Collected PM])",
            "formatString": "0.0%"
        },
        {
            "name": "Collection Rate",
            "expression": "DIVIDE([Payments Collected],[Overdue Amount PM])",
            "formatString": "0.0%"
        },
        {
            "name": "Collection Rate Label",
            "expression": 'FORMAT([Collection Rate],"0.0%")'
        },
    ]

    fact_columns = [
        {"name": "customer",     "dataType": "string",   "sourceColumn": "customer"},
        {"name": "country",      "dataType": "string",   "sourceColumn": "country"},
        {"name": "month_date",   "dataType": "dateTime", "sourceColumn": "month_date",
         "formatString": "MMM-yy"},
        {"name": "dso",          "dataType": "double",   "sourceColumn": "dso"},
        {"name": "ar_eur",       "dataType": "decimal",  "sourceColumn": "ar_eur",
         "formatString": "#,##0.00"},
        {"name": "overdue_eur",  "dataType": "decimal",  "sourceColumn": "overdue_eur",
         "formatString": "#,##0.00"},
        {"name": "overdue_pct",  "dataType": "double",   "sourceColumn": "overdue_pct",
         "formatString": "0.0%"},
        {"name": "gt60_eur",     "dataType": "decimal",  "sourceColumn": "gt60_eur",
         "formatString": "#,##0.00"},
        {"name": "gt60_pct",     "dataType": "double",   "sourceColumn": "gt60_pct",
         "formatString": "0.0%"},
        {"name": "gt90_eur",     "dataType": "decimal",  "sourceColumn": "gt90_eur",
         "formatString": "#,##0.00"},
        {"name": "gt90_pct",     "dataType": "double",   "sourceColumn": "gt90_pct",
         "formatString": "0.0%"},
        {"name": "payments_collected_eur", "dataType": "decimal",
         "sourceColumn": "payments_collected_eur", "formatString": "#,##0.00",
         "isNullable": True},
    ]

    date_columns = [
        {"name": "date",         "dataType": "dateTime", "sourceColumn": "date",
         "formatString": "MMM-yy", "isKey": True},
        {"name": "year",         "dataType": "int64",    "sourceColumn": "year"},
        {"name": "month_num",    "dataType": "int64",    "sourceColumn": "month_num"},
        {"name": "month_name",   "dataType": "string",   "sourceColumn": "month_name"},
        {"name": "month_short",  "dataType": "string",   "sourceColumn": "month_short"},
        {"name": "quarter",      "dataType": "string",   "sourceColumn": "quarter"},
        {"name": "sort_key",     "dataType": "int64",    "sourceColumn": "sort_key"},
    ]

    schema = {
        "name": "AMKAD_OTC",
        "compatibilityLevel": 1550,
        "model": {
            "culture": "en-US",
            "dataAccessOptions": {"legacyRedirects": True, "returnErrorValuesAsNull": True},
            "defaultPowerBIDataSourceVersion": "powerBI_V3",
            "expressions": [
                {
                    "name": "Excel_File_Path",
                    "kind": "m",
                    "expression": M_PARAM
                }
            ],
            "tables": [
                {
                    "name": "fact_ar_performance",
                    "columns": fact_columns,
                    "measures": measures_fact,
                    "partitions": [
                        {
                            "name": "fact_ar_performance",
                            "mode": "import",
                            "source": {
                                "type": "m",
                                "expression": M_FACT
                            }
                        }
                    ]
                },
                {
                    "name": "dim_date",
                    "columns": date_columns,
                    "partitions": [
                        {
                            "name": "dim_date",
                            "mode": "import",
                            "source": {
                                "type": "m",
                                "expression": M_DIM_DATE
                            }
                        }
                    ]
                }
            ],
            "relationships": [
                {
                    "name": "fact_to_dim_date",
                    "fromTable": "fact_ar_performance",
                    "fromColumn": "month_date",
                    "toTable": "dim_date",
                    "toColumn": "date",
                    "crossFilteringBehavior": "oneDirection"
                }
            ],
            "annotations": [
                {"name": "PBIDesktopVersion", "value": "2.130"},
                {"name": "PBI_QueryOrder", "value": '["fact_ar_performance","dim_date"]'}
            ]
        }
    }
    return schema


# ─────────────────────────────────────────────────────────────────────────────
#  REPORT LAYOUT  (4 pages, basic visuals wired to the measures)
# ─────────────────────────────────────────────────────────────────────────────

def kpi_card(name, x, y, w, h, measure_table, measure_name):
    """Returns a single-value KPI card visual config."""
    return {
        "id": abs(hash(name)) % 1000000,
        "x": x, "y": y, "z": 0, "width": w, "height": h,
        "config": json.dumps({
            "name": name,
            "layouts": [{"id": 0, "position": {"x": x, "y": y, "z": 0, "width": w, "height": h}}],
            "singleVisual": {
                "visualType": "card",
                "projections": {
                    "Values": [{"queryRef": f"{measure_table}.{measure_name}", "active": True}]
                },
                "prototypeQuery": {
                    "Select": [{"Measure": {"Expression": {"SourceRef": {"Source": "f"}},
                                            "Property": measure_name}, "Name": f"{measure_table}.{measure_name}"}],
                    "From": [{"Name": "f", "Entity": measure_table, "Type": 0}]
                },
                "vcObjects": {
                    "title": [{"properties": {"text": {"expr": {"Literal": {"Value": f"'{measure_name}'"}}}}}]
                }
            }
        })
    }


def table_visual(name, x, y, w, h, columns):
    """Returns a table visual config."""
    select = []
    src = []
    projections = {"Values": []}
    for i, (tbl, col, col_name) in enumerate(columns):
        alias = chr(ord('a') + i)
        src.append({"Name": alias, "Entity": tbl, "Type": 0})
        ref = {"queryRef": f"{tbl}.{col_name}"}
        projections["Values"].append(ref)
        if col.startswith("["):
            select.append({"Measure": {"Expression": {"SourceRef": {"Source": alias}},
                                        "Property": col_name}, "Name": f"{tbl}.{col_name}"})
        else:
            select.append({"Column": {"Expression": {"SourceRef": {"Source": alias}},
                                       "Property": col_name}, "Name": f"{tbl}.{col_name}"})
    return {
        "id": abs(hash(name)) % 1000000,
        "x": x, "y": y, "z": 0, "width": w, "height": h,
        "config": json.dumps({
            "name": name,
            "singleVisual": {
                "visualType": "tableEx",
                "projections": projections,
                "prototypeQuery": {"Select": select, "From": src}
            }
        })
    }


def line_chart(name, x, y, w, h, x_tbl, x_col, y_tbl, y_measure):
    return {
        "id": abs(hash(name)) % 1000000,
        "x": x, "y": y, "z": 0, "width": w, "height": h,
        "config": json.dumps({
            "name": name,
            "singleVisual": {
                "visualType": "lineChart",
                "projections": {
                    "Category": [{"queryRef": f"{x_tbl}.{x_col}"}],
                    "Y":         [{"queryRef": f"{y_tbl}.{y_measure}", "active": True}]
                },
                "prototypeQuery": {
                    "Select": [
                        {"Column":  {"Expression": {"SourceRef": {"Source": "d"}}, "Property": x_col},
                         "Name": f"{x_tbl}.{x_col}"},
                        {"Measure": {"Expression": {"SourceRef": {"Source": "f"}}, "Property": y_measure},
                         "Name": f"{y_tbl}.{y_measure}"}
                    ],
                    "From": [
                        {"Name": "d", "Entity": x_tbl, "Type": 0},
                        {"Name": "f", "Entity": y_tbl, "Type": 0}
                    ]
                }
            }
        })
    }


def build_page(name, display_name, visuals, ordinal):
    return {
        "name": name,
        "displayName": display_name,
        "displayOption": 1,
        "height": 720,
        "width": 1280,
        "visualContainers": visuals,
        "ordinal": ordinal
    }


def build_report_layout(model_id="AMKAD_OTC"):
    W, H = 1280, 720
    card_w, card_h = 250, 110
    gap = 20

    # ── Page 1: DSO Performance ───────────────────────────────────────────────
    p1_visuals = [
        kpi_card("dso_card",    gap, gap, card_w, card_h, "fact_ar_performance", "DSO"),
        kpi_card("ar_card",     gap*2+card_w, gap, card_w, card_h, "fact_ar_performance", "AR Total"),
        kpi_card("dso_chg",     gap*3+card_w*2, gap, card_w, card_h, "fact_ar_performance", "DSO MoM Change"),
        kpi_card("dso_trend",   gap*4+card_w*3, gap, card_w, card_h, "fact_ar_performance", "DSO Trend Indicator"),
        line_chart("dso_trend_chart", gap, card_h+gap*3, W-gap*2, 260,
                   "dim_date", "month_short", "fact_ar_performance", "DSO"),
        table_visual("dso_top10", gap, card_h+260+gap*5, W-gap*2, 230, [
            ("fact_ar_performance", "customer",       "customer"),
            ("fact_ar_performance", "[DSO]",          "DSO"),
            ("fact_ar_performance", "[DSO MoM Change]","DSO MoM Change"),
            ("fact_ar_performance", "[DSO Impact]",   "DSO Impact"),
        ]),
    ]

    # ── Page 2: Cash Collection ───────────────────────────────────────────────
    p2_visuals = [
        kpi_card("pay_card",  gap, gap, card_w, card_h, "fact_ar_performance", "Payments Collected"),
        kpi_card("pay_chg",   gap*2+card_w, gap, card_w, card_h, "fact_ar_performance", "Payments MoM Change %"),
        kpi_card("coll_rate", gap*3+card_w*2, gap, card_w, card_h, "fact_ar_performance", "Collection Rate Label"),
        line_chart("pay_trend", gap, card_h+gap*3, W-gap*2, 350,
                   "dim_date", "month_short", "fact_ar_performance", "Payments Collected"),
    ]

    # ── Page 3: Overdue & Aging ───────────────────────────────────────────────
    p3_visuals = [
        kpi_card("ovd_pct",  gap, gap, card_w, card_h, "fact_ar_performance", "Overdue % Label"),
        kpi_card("gt60_pct", gap*2+card_w, gap, card_w, card_h, "fact_ar_performance", "GT60 %"),
        kpi_card("gt90_lbl", gap*3+card_w*2, gap, card_w, card_h, "fact_ar_performance", "GT90 % Label"),
        kpi_card("ovd_eur",  gap*4+card_w*3, gap, card_w, card_h, "fact_ar_performance", "Overdue Amount"),
        table_visual("aging_table", gap, card_h+gap*3, W-gap*2, 300, [
            ("fact_ar_performance", "customer",        "customer"),
            ("fact_ar_performance", "[Overdue Amount]","Overdue Amount"),
            ("fact_ar_performance", "[Overdue %]",     "Overdue %"),
            ("fact_ar_performance", "[GT60 Amount]",   "GT60 Amount"),
            ("fact_ar_performance", "[GT90 Amount]",   "GT90 Amount"),
        ]),
    ]

    # ── Page 4: PPT Export ───────────────────────────────────────────────────
    p4_visuals = [
        kpi_card("e_dso",     gap, gap, card_w, card_h, "fact_ar_performance", "DSO"),
        kpi_card("e_dso_chg", gap*2+card_w, gap, card_w, card_h, "fact_ar_performance", "DSO MoM Change"),
        kpi_card("e_ovd",     gap*3+card_w*2, gap, card_w, card_h, "fact_ar_performance", "Overdue % Label"),
        kpi_card("e_gt90",    gap*4+card_w*3, gap, card_w, card_h, "fact_ar_performance", "GT90 % Label"),
        table_visual("export_kpi_table", gap, card_h+gap*3, W//2-gap*2, 300, [
            ("fact_ar_performance", "customer",         "customer"),
            ("fact_ar_performance", "[DSO]",            "DSO"),
            ("fact_ar_performance", "[DSO MoM Change]", "DSO MoM Change"),
            ("fact_ar_performance", "[Overdue %]",      "Overdue %"),
            ("fact_ar_performance", "[GT90 %]",         "GT90 %"),
        ]),
        table_visual("export_top_contributors", W//2+gap, card_h+gap*3, W//2-gap*2, 300, [
            ("fact_ar_performance", "customer",          "customer"),
            ("fact_ar_performance", "[DSO Impact]",      "DSO Impact"),
            ("fact_ar_performance", "[Overdue MoM Change EUR]", "Overdue MoM Change EUR"),
            ("fact_ar_performance", "[GT90 MoM Change EUR]",   "GT90 MoM Change EUR"),
        ]),
    ]

    layout = {
        "id": 0,
        "resourcePackages": [],
        "sections": [
            build_page("dso_page",   "DSO Performance",   p1_visuals, 0),
            build_page("cash_page",  "Cash Collection",   p2_visuals, 1),
            build_page("aging_page", "Overdue & Aging",   p3_visuals, 2),
            build_page("ppt_page",   "PPT Export",        p4_visuals, 3),
        ],
        "config": json.dumps({
            "version": "5.43",
            "themeCollection": {"baseTheme": {"name": "CY24SU02"}},
            "activeSectionIndex": 0,
            "defaultDrillFilterOtherVisuals": True,
            "linguisticSchemaSyncVersion": 0
        }),
        "filters": "[]",
        "layoutOptimization": 0,
        "layoutVersion": 5,
        "themeCollection": {"baseTheme": {"name": "CY24SU02"}}
    }
    return layout


# ─────────────────────────────────────────────────────────────────────────────
#  CONTENT TYPES  &  METADATA
# ─────────────────────────────────────────────────────────────────────────────

CONTENT_TYPES = """<?xml version="1.0" encoding="utf-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="json" ContentType="application/json" />
  <Override PartName="/DataModelSchema" ContentType="application/json" />
  <Override PartName="/DiagramLayout"   ContentType="application/json" />
  <Override PartName="/Report/Layout"   ContentType="application/json" />
  <Override PartName="/Metadata"        ContentType="application/json" />
  <Override PartName="/Version"         ContentType="application/octet-stream" />
  <Override PartName="/SecurityBindings" ContentType="application/octet-stream" />
</Types>"""

METADATA = json.dumps({
    "version": "3.0",
    "createdFromTemplate": False,
    "description": "DHL AMKAD OTC/SPR AR Reporting Template",
    "author": "AMKAD Finance Ops",
    "autoBuiltReport": False
})

DIAGRAM_LAYOUT = json.dumps({
    "version": 1,
    "diagramLayout": {
        "nodeLayouts": {
            "fact_ar_performance": {"x": 400, "y": 200, "width": 280, "height": 400},
            "dim_date":            {"x": 50,  "y": 100, "width": 200, "height": 200}
        }
    }
})

SECURITY_BINDINGS = b""
VERSION = "3.0"


def _utf16(s: str) -> bytes:
    """Power BI .pbit internal text files are UTF-16 LE with BOM."""
    return s.encode("utf-16-le")


def _utf16_bom(s: str) -> bytes:
    return b"\xff\xfe" + s.encode("utf-16-le")


# ─────────────────────────────────────────────────────────────────────────────
#  ASSEMBLE  .pbit
# ─────────────────────────────────────────────────────────────────────────────

def write_pbit(output_path: str):
    schema  = build_model_schema()
    layout  = build_report_layout()

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        # [Content_Types].xml is the only text file kept as UTF-8.
        zf.writestr("[Content_Types].xml",  CONTENT_TYPES)
        # All other Power BI text files must be UTF-16 LE (no BOM).
        zf.writestr("DataModelSchema",      _utf16(json.dumps(schema, ensure_ascii=False, indent=2)))
        zf.writestr("DiagramLayout",        _utf16(DIAGRAM_LAYOUT))
        zf.writestr("Metadata",             _utf16(METADATA))
        zf.writestr("Report/Layout",        _utf16(json.dumps(layout, ensure_ascii=False, indent=2)))
        zf.writestr("SecurityBindings",     SECURITY_BINDINGS)
        zf.writestr("Version",              _utf16(VERSION))

    with open(output_path, "wb") as f:
        f.write(buf.getvalue())

    size_kb = os.path.getsize(output_path) / 1024
    print(f"✓  Written: {output_path}  ({size_kb:.1f} KB)")
    print()
    print("Next steps:")
    print("  1. Open AMKAD_OTC_Report.pbit in Power BI Desktop")
    print("  2. When prompted, set  Excel_File_Path  to your .xlsm file location")
    print("  3. Click Load — all 4 pages will populate automatically")
    print()
    print("  If column names in your Excel differ from the expected ones,")
    print("  edit the Power Query via  Transform Data → Advanced Editor")
    print("  and update the rename mapping (see docs/build_guide.md).")


if __name__ == "__main__":
    out = os.path.join(os.path.dirname(__file__), "AMKAD_OTC_Report.pbit")
    write_pbit(out)
