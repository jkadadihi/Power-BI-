# Cust Sol — Power Query Engine (CHOSEN approach)

Chosen because it keeps Power Automate to a **single action**. Power Query
does all the data work (mapping + full-history accumulation); Power Automate
only makes the dated file, which is the one thing Power Query physically
cannot do (it outputs into its own workbook, it can't create a separate file).

**How the pieces fit:**

```
Cust_Sol Daily File folder  (Gerard drops a raw source file each day)
        │
        ▼
ENGINE workbook  (lives OUTSIDE the folder, saved as .xlsm)
  └─ Power Query "RawData" — combines the WHOLE folder → full history
        │  (refreshed daily by the existing open-on-schedule flow)
        ▼
Power Automate  → ONE copy of the refreshed engine into the Daily folder
                  as "Daily_Performance_Report- <date>  AMKAD.xlsm"
        │
        ▼
Downstream HTML query reads the newest Daily_Performance_Report .xlsm's
"RawData" table — unchanged, this feeds it exactly what it expects.
```

The engine holds the single source of truth (full history, rebuilt every
refresh from the folder). The dated files are disposable snapshots.

> **The cost of this approach, so it's a conscious choice:** history lives in
> the raw Cust_Sol source files, so those must be **kept permanently** — if old
> ones are deleted, that history disappears from the engine on the next
> refresh. And each refresh re-reads every file in the folder, so it slows as
> files accumulate (fine for ~a year; the long-term fix is the Power BI
> dataflow / database route).

---

## 1. Build the engine query (one time)

1. Open (or create) the **engine workbook** — a single workbook that lives
   *outside* the Cust_Sol Daily File folder. **Save it as `.xlsm`**
   (Excel Macro-Enabled Workbook) — the downstream query only looks for
   `.xlsm` files.
2. **Data → Get Data → Blank Query → Advanced Editor**.
3. Paste the entire contents of `power_query/custsol_rawdata_eur.m`.
4. Name the query **RawData**. Done → **Close & Load To… → Table** on a sheet.
5. **Critical:** click a cell in the loaded table → **Table Design** → set the
   **Table Name** to exactly **`RawData`**. The downstream HTML query reads
   `Excel.Workbook(...){[Item="RawData", Kind="Table"]}` — it finds the table
   by that literal name, so if the ListObject is called `Table1` or anything
   else, the pipeline won't see the data.
6. First run will prompt for credentials to the SharePoint site — sign in
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

This is the **only** Power Automate in the whole design. The essential work is
one action — **Create file** (a renamed copy). The rest is just building the
filename and making sure the month folder exists.

**Trigger — chain it onto your existing refresh flow (recommended).** Add these
actions to the *end* of the open-on-schedule flow that already refreshes the
engine. That guarantees the copy happens *after* the refresh, with no second
trigger and no timing guesswork. (Alternatively, a separate *Recurrence* trigger
set a few minutes after the refresh works too — just less certain on timing.)

| # | Action | Settings |
|---|--------|----------|
| 1 | Initialize variable `varTargetFile` (String) | `concat('Daily_Performance_Report- ', formatDateTime(convertFromUtc(utcNow(), 'Eastern Standard Time'), 'MMMM d yyyy'), '  AMKAD.xlsm')` |
| 2 | Initialize variable `varMonthFolder` (String) | `concat('/Shared Documents/BS33384_AMKAD/Steering Board Presentations/2026/2026 Presentations/2026 AMKAD Summaries/Daily/', formatDateTime(convertFromUtc(utcNow(), 'Eastern Standard Time'), 'MMMM yyyy'))` |
| 3 | **Get file content** (SharePoint) | the engine `.xlsm` workbook |
| 4 | **Create new folder** (SharePoint) | Path: `varMonthFolder`. *Configure the next step to run even if this fails* — the month folder already exists on every day except the 1st. |
| 5 | **Create file** (SharePoint) | Folder: `varMonthFolder`, File Name: `varTargetFile`, File Content: output of step 3. Set **Configure run after** on step 5 to run on both **Succeeded** and **Failed** of step 4. |

**Why "Get file content + Create file" instead of the "Copy file" action:**
SharePoint's Copy file keeps the original's name — it can't rename the copy to
today's date. Get-content + Create-file is what lets the copy be named
`Daily_Performance_Report- <date>  AMKAD.xlsm`.

That's the whole thing. No Office Scripts, no read/append, no duplicate guard —
Power Query owns the data, Power Automate just snapshots the file once a day.

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
- **Frozen snapshots:** the dated copies contain the engine's live query, but
  the downstream reads their *saved* values (Excel.Workbook reads stored data,
  it doesn't re-run queries), so a copy always reflects the engine's state at
  copy time. To be extra safe you can set the engine query to not refresh on
  open in the copies, but it isn't required for the pipeline.
- **The Office Scripts** (`office_scripts/custsol_*`), the clone-forward flow
  artifacts, and `docs/CUSTSOL_DAILY_AUTOMATION.md` remain in the repo as a
  tested alternative (more Power Automate, but history doesn't depend on
  retaining source files). This Power Query engine is the chosen approach.
