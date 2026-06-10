# DHL AMKAD Power BI Automation

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
