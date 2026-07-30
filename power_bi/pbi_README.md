# AMKAD AR in Power BI

Port of the AMKAD AR reporting to a Power BI model. Reads the same SharePoint
sources as the Excel/HTML pipeline; changes nothing about it and can run
alongside it.

## The one structural change worth knowing

The Excel `RawData` query reads the **newest** `Daily_Performance_Report`
workbook. That workbook is a copy-forward chain: each day's file is a copy of
yesterday's plus one appended day. When a pipeline run fails, that day is
missing from every file built afterwards and no later refresh brings it back,
which is why the Excel side needs the self-healing repair query.

Power BI reads **every Cust Sol source file directly** and unions them, so the
history is rebuilt from source on each refresh. A failed pipeline run is
invisible: the day appears as soon as its source file exists. No backfill
logic, no gap repair, no dependence on the Excel chain.

## Queries

| Query | Grain | Source | Use it for |
|---|---|---|---|
| `fact_ar_daily` | Date × Country × Customer | Cust Sol daily files | Daily trends, current position |
| `fact_ar_monthly` | Month × Country × Customer | derived from `fact_ar_daily` | Month trends on a consistent daily basis |
| `fact_ar_month_end` | Month × Country × Customer | Debits workbooks | SPR / Account Business Review |
| `dim_date`, `dim_customer`, `dim_basis` | — | derived | Shared slicers |

`fact_ar_monthly` keeps the **last daily snapshot** of each month. That is not
the close — month close shifts by days and postings land after the final
snapshot. For anything reported externally, use `fact_ar_month_end`.

## Setup

1. **Home > Transform data > Manage Parameters > New**
   Name `SharePointSite_Url`, type Text, current value
   `https://dpdhl.sharepoint.com/teams/EXP-USQIA-BS33384_AMKAD`
2. Create the queries in this order — later ones reference earlier ones:
   `fact_ar_daily`, `fact_ar_month_end`, `fact_ar_monthly`, `dim_date`,
   `dim_customer`, `dim_basis`.
3. Sign in with **Organizational account**. Privacy level **Organizational**.
4. Add the `CustKey` step to both fact queries (see the comment block at the
   top of `pbi_04_dimensions.m`).
5. **Close & Apply**, then build the relationships listed in that same comment
   block. All single-direction, one-to-many from dimension to fact.
6. Mark `dim_date` as a date table: right-click it > **Mark as date table** >
   column `date`. Time intelligence gives wrong answers without this.
7. On `dim_date`, set **Sort by column**: `month_name` sorted by `month_sort`,
   `month_short` sorted by `month_sort`. Otherwise chart axes read April,
   August, December, February.

The first refresh opens every source file and is slow — expect a few minutes.
Later refreshes are the same cost, so if the folder grows past a couple of
years, set up incremental refresh on `fact_ar_daily` partitioned by `Date`.

## Measures

The percentage and DSO columns that are formulas in the Excel sheet are
deliberately **not** recreated as columns. Percentages must never be averaged
across rows — a calculated column gives wrong subtotals at every level. They
belong in DAX, where the ratio is recomputed at whatever level you view it.

```dax
Total AR = SUM ( fact_ar_daily[Total AR] )
Overdue  = SUM ( fact_ar_daily[Overdue] )
GT60     = SUM ( fact_ar_daily[GT60] )
GT90     = SUM ( fact_ar_daily[GT90] )
Total UAC = SUM ( fact_ar_daily[Total UAC] )
Payments  = SUM ( fact_ar_daily[Total Payments] )

Overdue % = DIVIDE ( [Overdue], [Total AR] )
GT60 %    = DIVIDE ( [GT60],    [Total AR] )
GT90 %    = DIVIDE ( [GT90],    [Total AR] )
```

The buckets are **nested and cumulative**: GT90 ⊆ GT60 ⊆ Overdue ⊆ Total AR.
Never put them in a stacked column chart or a pie — they overlap, so the stack
double-counts. Use side-by-side bars or separate lines.

### Balance vs flow

AR, Overdue, GT60, GT90 and UAC are **balances**. Summing them across dates is
meaningless: twenty daily snapshots of a €4m book sum to €80m. When a visual
spans more than one date, take the closing position:

```dax
AR (closing) =
CALCULATE (
    [Total AR],
    LASTNONBLANK ( dim_date[date], [Total AR] )
)
```

Gross Sales and Payments are **flows** and do sum correctly over a period.

### Trend direction

A change is not good or bad because of its sign — it depends on the metric.
A fall in Overdue, GT60, GT90, UAC or Total AR is **good**; a rise in Payments
or Gross Sales is **good**. Set conditional formatting per measure accordingly
rather than applying one red/green rule to everything. This was a real bug in
the HTML report: the arrows were coloured by direction instead of by whether
the movement was favourable.

```dax
AR MoM =
VAR Curr = [AR (closing)]
VAR Prev = CALCULATE ( [AR (closing)], DATEADD ( dim_date[date], -1, MONTH ) )
RETURN Curr - Prev

AR MoM Colour =                         -- lower is better
IF ( [AR MoM] <= 0, "#1B8A5A", "#C0392B" )

Payments MoM Colour =                   -- higher is better
IF ( [Payments MoM] >= 0, "#1B8A5A", "#C0392B" )
```

### Basis toggle

To reproduce the report's "current data / EOM data" confirmation label, put
`dim_basis[Basis]` on a slicer and switch the source fact:

```dax
Selected Basis = SELECTEDVALUE ( dim_basis[Basis], "Month-end close" )

AR (selected basis) =
IF (
    [Selected Basis] = "Month-end close",
    SUM ( fact_ar_month_end[Total AR] ),
    [AR (closing)]
)

Basis Label =
VAR B = [Selected Basis]
VAR AsOf =
    IF (
        B = "Month-end close",
        MAX ( fact_ar_month_end[MonthKey] ),
        FORMAT ( MAX ( fact_ar_daily[Date] ), "dd MMM yyyy" )
    )
RETURN B & " — " & AsOf
```

Put `Basis Label` in a card on every page that can switch basis. Silently
mixing the two is the failure mode: a country total built from one basis and a
customer breakdown built from the other will not reconcile, and nothing on
screen says why.

## Country and customer pages

Sort customer tables by `GT90` descending, not by name — the critical accounts
should be at the top without scrolling. For the country-level over-90 figure,
the ratio has to be recomputed at country level, not averaged from the
customer rows:

```dax
Country GT90 % =
DIVIDE (
    CALCULATE ( [GT90],     ALLEXCEPT ( dim_customer, dim_customer[Country] ) ),
    CALCULATE ( [Total AR], ALLEXCEPT ( dim_customer, dim_customer[Country] ) )
)
```

## Known source quirks these queries already handle

- The source sheet has **"Total Payments" twice** (native currency and EUR).
  Columns are read by position, not header name; promoting headers renames one
  of them and silently binds the mapping to the wrong column.
- **Debits header spellings vary** between files — lowercase `overdue` next to
  `Overdue €`, a space after `>` in the day buckets. Power Query column names
  are case- and space-sensitive, and an unmatched name expands to nulls that
  then zero-fill and look like real zeros. Headers are normalised before use.
- **SAP exports payments as negatives.** Flipped to positive to match reporting.
- **Customer names carry trailing spaces.** `CustKey` upper-cases and trims, or
  one account splits into two phantom rows.
- **A customer name can appear in more than one country** and is a different
  account each time. The key is always Country + Customer.
- **Individual files can be locked or unreadable.** Each file read is wrapped
  best-effort, so one bad file costs that day rather than the whole refresh.
