═══════════════════════════════════════════════════════════════════════════════
  DHL AMKAD OTC REPORT — MANUAL BUILD (10 minutes)
  Use this when the .pbit won't open in your Power BI Desktop.
═══════════════════════════════════════════════════════════════════════════════

You will:
  STEP 1  Connect to your Excel file
  STEP 2  Paste 2 Power Query scripts
  STEP 3  Paste 25 DAX measures
  STEP 4  Build the 4 report pages

────────────────────────────────────────────────────────────────────────────────
STEP 1 — Open a blank report and connect to Excel
────────────────────────────────────────────────────────────────────────────────

1. Open Power BI Desktop → click "Blank report"
2. Home ribbon → Get Data → Excel workbook
3. Browse to:
     C:\Users\jp1lq1\OneDrive - DPDHL\Desktop\BI AMKAD\Monthly_Performance_Report- April 2026 AMKAD - All Customers.xlsm
4. In the Navigator, tick "Raw Data EUR Onboarding"
5. Click "Transform Data" (NOT Load) — this opens Power Query Editor

────────────────────────────────────────────────────────────────────────────────
STEP 2 — Replace the auto-query with the clean version
────────────────────────────────────────────────────────────────────────────────

You're now in Power Query Editor. On the left, you'll see one query
called "Raw Data EUR Onboarding".

  2a) Right-click that query → Rename → type:  fact_ar_performance

  2b) Home ribbon → Advanced Editor → DELETE everything in the window,
      then paste the ENTIRE block below (between the ===== lines):

===== PASTE THIS BLOCK FOR fact_ar_performance =====
let
    Source = Excel.Workbook(
        File.Contents("C:\Users\jp1lq1\OneDrive - DPDHL\Desktop\BI AMKAD\Monthly_Performance_Report- April 2026 AMKAD - All Customers.xlsm"),
        null, true
    ),
    RawSheet = Source{[Item="Raw Data EUR Onboarding", Kind="Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(RawSheet, [PromoteAllScalars=true]),
    RemovedBlanks = Table.SelectRows(
        PromotedHeaders,
        each not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), {"", null}))
    ),
    SelectedColumns = Table.SelectColumns(RemovedBlanks, {
        "Customer", "Country", "Month", "DSO",
        "AR Amount (EUR)", "Overdue Amount (EUR)", "Overdue %",
        ">60 Days (EUR)", ">60 Days %", ">90 Days (EUR)", ">90 Days %"
    }),
    RenamedColumns = Table.RenameColumns(SelectedColumns, {
        {"Customer", "customer"}, {"Country", "country"},
        {"Month", "month_raw"}, {"DSO", "dso"},
        {"AR Amount (EUR)", "ar_eur"},
        {"Overdue Amount (EUR)", "overdue_eur"},
        {"Overdue %", "overdue_pct"},
        {">60 Days (EUR)", "gt60_eur"}, {">60 Days %", "gt60_pct"},
        {">90 Days (EUR)", "gt90_eur"}, {">90 Days %", "gt90_pct"}
    }),
    ParsedMonth = Table.AddColumn(RenamedColumns, "month_date",
        each let raw = Text.Trim(Text.From([month_raw])),
                 parsed = try Date.FromText(raw, [Format="MMM-yy", Culture="en-US"])
                          otherwise try Date.FromText(raw, [Format="MMMM yyyy", Culture="en-US"])
                          otherwise null
             in parsed,
        type date),
    TypedTable = Table.TransformColumnTypes(ParsedMonth, {
        {"customer", type text}, {"country", type text},
        {"dso", type number}, {"ar_eur", Currency.Type},
        {"overdue_eur", Currency.Type}, {"overdue_pct", Percentage.Type},
        {"gt60_eur", Currency.Type}, {"gt60_pct", Percentage.Type},
        {"gt90_eur", Currency.Type}, {"gt90_pct", Percentage.Type}
    }),
    FinalTable = Table.RemoveColumns(TypedTable, {"month_raw"})
in
    FinalTable
===== END =====

      Click Done.

  2c) Create the date dimension:
      Home ribbon → New Source → Blank Query
      Right-click the new query (Query1) → Rename → type:  dim_date
      Home ribbon → Advanced Editor → paste:

===== PASTE THIS BLOCK FOR dim_date =====
let
    MinDate = List.Min(fact_ar_performance[month_date]),
    MaxDate = List.Max(fact_ar_performance[month_date]),
    MonthCount = Number.RoundUp(Duration.Days(MaxDate - MinDate) / 30) + 2,
    DateList = List.Dates(MinDate, MonthCount, #duration(30, 0, 0, 0)),
    DateTable = Table.FromList(DateList, Splitter.SplitByNothing(), {"date"}),
    TypedDates = Table.TransformColumnTypes(DateTable, {{"date", type date}}),
    WithYear = Table.AddColumn(TypedDates, "year", each Date.Year([date]), Int64.Type),
    WithMonth = Table.AddColumn(WithYear, "month_num", each Date.Month([date]), Int64.Type),
    WithMonthNm = Table.AddColumn(WithMonth, "month_name", each Date.ToText([date], "MMMM"), type text),
    WithMonthSh = Table.AddColumn(WithMonthNm, "month_short", each Date.ToText([date], "MMM-yy"), type text),
    WithQuarter = Table.AddColumn(WithMonthSh, "quarter", each "Q" & Text.From(Date.QuarterOfYear([date])), type text),
    WithSortKey = Table.AddColumn(WithQuarter, "sort_key", each Date.Year([date]) * 100 + Date.Month([date]), Int64.Type),
    FinalDimDate = Table.Sort(WithSortKey, {{"date", Order.Ascending}})
in
    FinalDimDate
===== END =====

      Click Done.

  2d) Home ribbon → Close & Apply
      (Power BI loads both tables. Wait for it to finish — should take <30s.)

────────────────────────────────────────────────────────────────────────────────
STEP 3 — Set up the relationship between the two tables
────────────────────────────────────────────────────────────────────────────────

1. Left rail → click the "Model" icon (looks like joined boxes)
2. You'll see both tables. Drag the field:
       fact_ar_performance[month_date]
   onto:
       dim_date[date]
3. A dialog appears. Confirm:
       Cardinality: Many to one (*:1)
       Cross-filter direction: Single
       Make active: yes
4. Click OK.

────────────────────────────────────────────────────────────────────────────────
STEP 4 — Add the DAX measures
────────────────────────────────────────────────────────────────────────────────

Switch to Report view (top icon on the left rail).

For each measure below:
  1. In the Data pane on the right, right-click fact_ar_performance → New measure
  2. In the formula bar at the top, DELETE the default "Measure = " text
  3. Paste the FULL measure (name = expression) from below
  4. Press Enter
  5. Repeat for the next measure

Paste them in this exact order (some measures reference others):

────────────────── GROUP A: BASE AGGREGATES (paste first) ──────────────────

AR Total = SUM(fact_ar_performance[ar_eur])

Overdue Amount = SUM(fact_ar_performance[overdue_eur])

GT60 Amount = SUM(fact_ar_performance[gt60_eur])

GT90 Amount = SUM(fact_ar_performance[gt90_eur])

Overdue % = DIVIDE([Overdue Amount], [AR Total])

GT60 % = DIVIDE([GT60 Amount], [AR Total])

GT90 % = DIVIDE([GT90 Amount], [AR Total])

Overdue Amount PM = CALCULATE([Overdue Amount], PREVIOUSMONTH(dim_date[date]))

Overdue % PM = CALCULATE([Overdue %], PREVIOUSMONTH(dim_date[date]))

GT90 Amount PM = CALCULATE([GT90 Amount], PREVIOUSMONTH(dim_date[date]))

Overdue MoM Change EUR = [Overdue Amount] - [Overdue Amount PM]

Overdue MoM Change Pct = [Overdue %] - [Overdue % PM]

GT90 MoM Change EUR = [GT90 Amount] - [GT90 Amount PM]

Overdue Impact Rank = RANKX(ALL(fact_ar_performance[customer]), [Overdue Amount], , DESC, DENSE)

Overdue % Label = FORMAT([Overdue %], "0.0%") & IF(NOT ISBLANK([Overdue % PM]), " (" & IF([Overdue MoM Change Pct] >= 0, "+", "") & FORMAT([Overdue MoM Change Pct], "0.0pp") & " MoM)", "")

GT90 % Label = FORMAT([GT90 %], "0.0%") & IF(NOT ISBLANK([GT90 Amount PM]), " (" & IF([GT90 MoM Change EUR] >= 0, "+", "") & FORMAT([GT90 MoM Change EUR] / 1000000, "#,##0.0M €") & " MoM)", "")

────────────────── GROUP B: DSO MEASURES ──────────────────

DSO = VAR TotalAR = SUM(fact_ar_performance[ar_eur]) RETURN DIVIDE(SUMX(fact_ar_performance, fact_ar_performance[dso] * fact_ar_performance[ar_eur]), TotalAR)

DSO PM = CALCULATE([DSO], PREVIOUSMONTH(dim_date[date]))

DSO MoM Change = [DSO] - [DSO PM]

DSO MoM Change Label = VAR delta = [DSO MoM Change] RETURN IF(ISBLANK(delta), BLANK(), IF(delta >= 0, "+" & FORMAT(delta, "0.0"), FORMAT(delta, "0.0")))

DSO Trend Indicator = VAR delta = [DSO MoM Change] RETURN SWITCH(TRUE(), delta > 2, "Worsening", delta < -2, "Improving", "Stable")

DSO Impact = [DSO] - [DSO PM]

DSO Impact Rank = RANKX(ALL(fact_ar_performance[customer]), [DSO Impact], , DESC, DENSE)

────────────────── GROUP C: CASH COLLECTION ──────────────────
(Only needed if your Excel has a payments_collected_eur column.
 If not, skip — your Cash Collection page will just be empty.)

Payments Collected = SUM(fact_ar_performance[payments_collected_eur])

Payments Collected PM = CALCULATE([Payments Collected], PREVIOUSMONTH(dim_date[date]))

Payments MoM Change % = DIVIDE([Payments Collected] - [Payments Collected PM], [Payments Collected PM])

Collection Rate = DIVIDE([Payments Collected], [Overdue Amount PM])

Collection Rate Label = FORMAT([Collection Rate], "0.0%")

────────────────────────────────────────────────────────────────────────────────
STEP 5 — Build the report pages
────────────────────────────────────────────────────────────────────────────────

Add 4 pages (right-click the page tab at the bottom → Duplicate page or +).
Rename them: "DSO Performance", "Cash Collection", "Overdue & Aging", "PPT Export".

For each visual:
  1. Click on the page where you want it
  2. In the Visualizations pane (right), click the visual type
  3. From the Data pane, drag the listed fields into the slots

──── Page 1: DSO Performance ────
  KPI Card → field: [DSO]                              (title: "DSO")
  KPI Card → field: [AR Total]                         (title: "AR Total")
  KPI Card → field: [DSO MoM Change]                   (title: "DSO vs Prior Month")
  Line chart → Axis: dim_date[month_short],
               Values: [DSO]                           (title: "DSO Trend")
  Table → Columns: customer, [DSO], [DSO MoM Change], [DSO Impact]
          Visual-level filter: [DSO Impact Rank] is <= 10

──── Page 2: Cash Collection ────
  (Skip if no payments_collected_eur column)
  KPI Card → [Payments Collected]
  KPI Card → [Collection Rate Label]
  Line chart → Axis: dim_date[month_short], Values: [Payments Collected]

──── Page 3: Overdue & Aging ────
  KPI Card → [Overdue % Label]
  KPI Card → [GT60 %]
  KPI Card → [GT90 % Label]
  Clustered bar chart → Axis: customer,
                        Values: [GT60 Amount], [GT90 Amount]
  Table → Columns: customer, [Overdue Amount], [Overdue %], [GT90 Amount]
          Visual-level filter: [Overdue Impact Rank] is <= 10

──── Page 4: PPT Export ────
  This is the clean page for copy-pasting to PowerPoint.
  Re-create your top 2-3 visuals here with:
    - White background
    - Font size >= 12pt
    - No page-level filters
    - Each chart has a self-contained title

────────────────────────────────────────────────────────────────────────────────
STEP 6 — Export visuals to PowerPoint
────────────────────────────────────────────────────────────────────────────────

For each visual you want in your SPR slide:
  1. Hover over the visual → click the "..." menu (top right of the visual)
  2. Copy → Copy as image
  3. Paste directly into your PowerPoint slide

────────────────────────────────────────────────────────────────────────────────
COLUMN NAME TROUBLESHOOTING
────────────────────────────────────────────────────────────────────────────────

If Step 2d throws an error like "the column 'Overdue %' wasn't found":
  Your Excel uses slightly different column names. Open the M code in
  Advanced Editor and edit the SelectedColumns + RenamedColumns blocks
  to match your actual headers exactly.

Common variations to look for:
  "AR Amount (EUR)"      may be   "AR (EUR)"  or  "Total AR"
  "Overdue Amount (EUR)" may be   "Overdue (EUR)"
  ">60 Days (EUR)"       may be   "60+ Days"  or  "Bucket 60"
  ">90 Days (EUR)"       may be   "90+ Days"  or  ">90"
