# Cust Sol — Power Query Engine (ALTERNATIVE — not the chosen approach)

> **Superseded.** The chosen approach is **clone-forward + append**
> (`docs/CUSTSOL_DAILY_AUTOMATION.md`): the newest daily `.xlsm` already
> contains the full running history in its `RawData` table (that's what the
> downstream HTML query reads), so each day copies yesterday's file forward
> and appends only today's rows. That matches the existing pipeline exactly
> and doesn't depend on retaining every raw Cust_Sol source file.
>
> This Power Query engine is kept only as a fallback for if the raw source
> files are guaranteed to be retained forever and you'd rather rebuild the
> whole history from them each day. If you use it, the output file must still
> be saved as **`.xlsm`** with a table named **`RawData`** and the
> `Daily_Performance_Report- <date>  AMKAD.xlsm` name, or the downstream query
> won't see it.

This is the Power Query version of the daily import. It reuses the pattern
already proven in the HTML pipeline (PQ reading daily SharePoint files) and
needs far less custom code — the whole transform is one M query.

**How the pieces fit:**

```
Cust_Sol Daily File folder  (Gerard drops a file each day)
        │
        ▼
ENGINE workbook  (lives OUTSIDE the folder)
  └─ Power Query "RawDataEUR" — combines the whole folder → full history
        │  (refreshed daily by the existing open-on-schedule flow)
        ▼
Power Automate  → copies the refreshed engine into the Daily folder
                  as "Daily_Performance_Report- <date>  AMKAD.xlsx"
```

The engine holds the single source of truth (full history, rebuilt every
refresh). The dated files in the Daily folder are disposable snapshots — if
one is corrupted, the engine is untouched.

---

## 1. Build the engine query (one time)

1. Open (or create) the **engine workbook** — a single workbook that lives
   *outside* the Cust_Sol Daily File folder.
2. **Data → Get Data → Blank Query → Advanced Editor**.
3. Paste the entire contents of `power_query/custsol_rawdata_eur.m`.
4. Name the query **RawDataEUR**. Done → **Close & Load To… → Table** on a
   sheet (this becomes your Raw Data EUR sheet).
5. First run will prompt for credentials to the SharePoint site — sign in
   with an org account that can read the folder (organizational, not
   anonymous).

The query outputs the 19 data columns (Date, Country, Customer, Go Live,
Payment Term, then the Q:AD block). Columns F:P from the old manual layout
are intentionally left out — not needed. If anything downstream references
the full A:AD shape, add null placeholder columns after the query loads.

### Why it's built this way
- **Whole folder, not "latest file":** Power Query rebuilds (never appends),
  so reading the *whole* folder each refresh is what gives you full history.
  Reading only the newest file would leave the table with just one day.
- **Headers not promoted / mapped by position:** the source has "Total
  Payments" as a header *twice* (native col G, Euro col V). Going by column
  position avoids the rename collision and also survives header renames.

---

## 2. Refresh (already solved)

Set the query to refresh on open: **Query → Properties → uncheck "Enable
background refresh" is optional, CHECK "Refresh data when opening the file."**
Your existing open-on-schedule Power Automate flow (the one that opens the
workbook so PQ runs on start) then refreshes it daily with no one touching it.

To keep the dated snapshots frozen, the *copies* placed in the Daily folder
should NOT refresh on open — but since the copy is made right after the engine
refreshes, it already carries the correct values, and a viewer opening it
later without folder access simply keeps those cached values.

---

## 3. Copy the refreshed engine to a dated file (Power Automate)

A small flow, run after the daily refresh:

| # | Action | Settings |
|---|--------|----------|
| 1 | **Trigger** — pick one: *Recurrence* (daily, a few min after your refresh flow) OR chain onto the end of the existing refresh flow | — |
| 2 | Initialize variable `varTargetFile` (String) | `concat('Daily_Performance_Report- ', formatDateTime(convertFromUtc(utcNow(), 'Eastern Standard Time'), 'MMMM d yyyy'), '  AMKAD.xlsx')` |
| 3 | Initialize variable `varMonthFolder` (String) | `concat('/Shared Documents/BS33384_AMKAD/Steering Board Presentations/2026/2026 Presentations/2026 AMKAD Summaries/Daily/', formatDateTime(convertFromUtc(utcNow(), 'Eastern Standard Time'), 'MMMM yyyy'))` |
| 4 | **Get file content** (SharePoint) | the engine workbook |
| 5 | **Create new folder** (SharePoint) | Path: `varMonthFolder`. *Configure run after step 6 to allow this to fail* (folder exists most days). |
| 6 | **Create file** (SharePoint) | Folder: `varMonthFolder`, Name: `varTargetFile`, Content: output of step 4. Run after step 5 on **Succeeded OR Failed**. |

That's the whole thing. No Office Scripts, no read/append, no duplicate guard —
Power Query owns the data, Power Automate just snapshots the file.

---

## 4. Notes / troubleshooting

- **Slows down over time:** combining the whole folder re-reads every file each
  refresh. Fine for a long while; a year (~250 files) is where you'd start to
  feel it. The clean long-term fix is the Power BI dataflow / database route
  (your "Gold" plan) — same folder, but the union lives in a dataset instead of
  re-reading Excel files daily.
- **Euro columns blank early in the month:** those rows load with blank Euro
  fields (no failure). Nothing to configure.
- **Missing day:** if no file lands, the folder just has one fewer file; the
  refresh still succeeds with whatever's there.
- **Year rollover:** `SiteUrl`/paths and the folder expressions contain `2026`
  — update the two variable expressions and the query's folder when 2027 folders
  are created.
- **The Office Scripts** (`office_scripts/custsol_*`) and their flow artifacts
  remain in the repo as a tested alternative, but this Power Query engine is the
  chosen approach and does not use them.
