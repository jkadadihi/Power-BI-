# DHL AMKAD Power BI Automation

This repo has two related but separate automation efforts:

- **AMKAD OTC/SPR Power BI report** (this README, root files) - a Power BI
  template that reports on AR/DSO/overdue performance from a monthly Excel
  export.
- **AMCOD month-end automation** (`amcod_automation/`) - Power Query,
  Office Script, and Power Automate building blocks that replace the manual
  copy/paste steps in Marcia Lopez Espinoza's monthly AMCOD close process
  for Global reporting. See `amcod_automation/docs/AMCOD_Automation_Plan.md`
  for the full plan and how each file maps back to the manual process.

## Quick Start

1. Clone this repo to your machine
2. Open terminal in the repo folder
3. Run: `python generate_pbit.py`
4. Open `AMKAD_OTC_Report.pbit` in Power BI Desktop
5. Set `Excel_File_Path` to your `.xlsm` path when prompted
6. Click Load

## Files

| File | Purpose |
|---|---|
| `generate_pbit.py` | Generates the `.pbit` template (run locally) |
| `power_query/01_load_ar_performance.m` | Power Query M - loads Excel sheet |
| `power_query/02_dim_date.m` | Power Query M - date dimension |
| `dax_measures/dso_measures.dax` | DAX - DSO, MoM change, trend |
| `dax_measures/overdue_aging_measures.dax` | DAX - overdue %, aging buckets |
| `dax_measures/cash_collection_measures.dax` | DAX - payments, collection rate |
| `docs/build_guide.md` | Step-by-step build guide |
| `amcod_automation/` | AMCOD month-end process automation (Power Query, Office Scripts, Power Automate) - see plan doc inside |
