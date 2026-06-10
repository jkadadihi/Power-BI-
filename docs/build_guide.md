# DHL AMKAD Power BI - Build Guide

## Quickest path

1. Clone this repo:
   ```
   git clone https://github.com/jkadadihi/Power-BI-
   cd Power-BI-
   ```
2. Run: `python generate_pbit.py`
3. Open `AMKAD_OTC_Report.pbit` in Power BI Desktop
4. When prompted, set `Excel_File_Path` to your file, e.g.:
   `C:\Users\jp1lq1\OneDrive - DPDHL\Desktop\BI AMKAD\Monthly_Performance_Report- April 2026 AMKAD - All Customers.xlsm`
5. Click Load

## Report pages

| Page | Visuals |
|---|---|
| DSO Performance | KPI cards, DSO trend line, top-10 contributor table |
| Cash Collection | KPI cards, payments trend line |
| Overdue & Aging | KPI cards, aging breakdown table |
| PPT Export | Clean KPI + contributor tables - copy/paste into SPR slides |

## Column name mapping

If your Excel column names differ, update the rename step in `power_query/01_load_ar_performance.m`.

| Expected | Common variant |
|---|---|
| AR Amount (EUR) | AR, AR (EUR), Total AR |
| Overdue Amount (EUR) | Overdue, Overdue EUR |
| >60 Days (EUR) | 60+, Bucket 60 |
| >90 Days (EUR) | 90+, Bucket 90, >90 |
