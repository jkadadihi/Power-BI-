# DHL AMKAD Power BI — Build Guide

## Step 1: Load Data (Power Query)

1. Open Power BI Desktop → **Get Data → Excel**
2. Select your `.xlsm` file
3. Open **Transform Data** (Power Query Editor)
4. Click **Advanced Editor** and paste the contents of:
   - `power_query/01_load_ar_performance.m` → rename query to `fact_ar_performance`
   - `power_query/02_dim_date.m` → rename query to `dim_date`
5. Update the file path on line 5 of `01_load_ar_performance.m` to match your local path
6. Click **Close & Apply**

---

## Step 2: Build the Data Model

Follow `data_model/relationships.md` to wire up relationships in Model view.

---

## Step 3: Add DAX Measures

For each `.dax` file in `dax_measures/`:

1. In Report view, click **New Measure** (Home ribbon)
2. Copy-paste each measure block (between the blank lines)
3. Name the measure exactly as written (the name is the first line before `=`)

### Measure load order
1. `overdue_aging_measures.dax` (base aggregates — other measures depend on these)
2. `dso_measures.dax`
3. `cash_collection_measures.dax`

---

## Step 4: Build Report Pages

### Page 1 — DSO Performance
| Visual | Type | Fields |
|---|---|---|
| DSO (current month) | KPI Card | `[DSO]`, `[DSO MoM Change]` |
| AR Total | KPI Card | `[AR Total]` |
| DSO Trend | Line Chart | X: `dim_date[month_short]`, Y: `[DSO]` |
| Top Contributors | Table | `customer`, `[DSO]`, `[DSO MoM Change]`, `[DSO Impact]` |

Filter: add `[DSO Impact Rank] <= 10` as a visual-level filter on the table.

---

### Page 2 — Cash Collection
| Visual | Type | Fields |
|---|---|---|
| Payments Collected | KPI Card | `[Payments Collected]`, `[Payments MoM Change %]` |
| Collection Rate | KPI Card | `[Collection Rate Label]` |
| Trend | Line Chart | X: `dim_date[month_short]`, Y: `[Payments Collected]` |

---

### Page 3 — Overdue & Aging
| Visual | Type | Fields |
|---|---|---|
| Overdue % | KPI Card | `[Overdue % Label]` |
| >60 Days % | KPI Card | `[GT60 %]` |
| >90 Days % | KPI Card | `[GT90 % Label]` |
| Aging Breakdown | Clustered Bar | Customer vs `[GT60 Amount]`, `[GT90 Amount]` |
| Worst Customers | Table | `customer`, `[Overdue Amount]`, `[Overdue %]`, `[GT90 Amount]` |

Filter: add `[Overdue Impact Rank] <= 10` on the table.

---

### Page 4 — PPT Export Page *(copy from here into slides)*

This page contains clean, slide-ready versions of the key visuals:

| Visual | Use in slide |
|---|---|
| KPI Summary Table (`customer`, DSO, Overdue %, >90) | Paste into DSO/Aging slide |
| Top 5 DSO Contributors bar | Paste into DSO slide |
| Top 5 Overdue Contributors bar | Paste into Aging slide |

**Design rules for this page:**
- White background
- No page-level filters visible
- Every chart has a self-contained title + unit label
- Font size ≥ 12pt (legible after paste into PPT at 50% zoom)

---

## Step 5: Export to PowerPoint

**Manual (current):**
1. Right-click any visual → **Copy → Copy as image**
2. Paste into the SPR PowerPoint slide

**Automated (future):**
Use the Power BI REST API `ExportTo` endpoint with `PowerPoint` format.

---

## Column Name Mapping

If your Excel sheet uses different column names, update the rename step
in `power_query/01_load_ar_performance.m` (the `RenamedColumns` block).

| Expected name in script | Likely Excel variant |
|---|---|
| `AR Amount (EUR)` | `AR`, `AR (EUR)`, `Total AR` |
| `Overdue Amount (EUR)` | `Overdue`, `Overdue EUR` |
| `>60 Days (EUR)` | `60+`, `Bucket 60` |
| `>90 Days (EUR)` | `90+`, `Bucket 90`, `>90` |
