# Cust Sol Daily Import — Build & Run Guide

Automates the daily manual process: take the customer solutions report that
lands in the **Cust_Sol Daily File** folder, clone the most recent
**Daily_Performance_Report** workbook into a new file dated today, and append
the day's rows to its **Raw Data EUR** table. History carries forward across
days and month folders. No alerts are sent by design.

**Status: the calculation logic is already tested.** Both Office Scripts pass
an offline test suite (33 checks) covering the full column mapping (A:E and
Q:AD), the duplicated "Total Payments" header, blank-row skipping, the
missing-Euro-data case, and the same-day re-run duplicate guard. Run it any
time with `tests/test_custsol.sh`. What is *not* pre-testable from outside
DHL's tenant is the Power Automate wiring — that's what this guide walks
through.

---

## 1. One-time setup: upload the two Office Scripts

1. Open any workbook in **Excel Online** (scripts are stored per-account, not
   per-file).
2. **Automate** tab → **New Script**, paste in the contents of
   `office_scripts/custsol_read_daily.ts`, rename the script
   **custsol read daily**, save.
3. Repeat with `office_scripts/custsol_append_rawdata_eur.ts`, named
   **custsol append rawdata eur**.

The writer expects the Raw Data EUR table to be named **`RawData`**
(confirmed) and to have the exact headers listed at the bottom of this doc.

---

## 2. Build the flow — three options, try in this order

### Option A — Import the package (fastest if it works)

1. Power Automate → **My flows** → **Import** → **Import Package (Legacy)**.
2. Upload `docs/power_automate/custsol_flow_package.zip`.
3. On the import screen, set the flow's *Import setup* to **Create as new**,
   and select your existing SharePoint and Excel Online (Business)
   connections for the two connector rows.
4. Import, then open the flow and do the **post-import fix-ups** (section 3).

If the import errors out (Microsoft's package format is picky), fall back to
Option B — do not spend more than a few minutes fighting it.

### Option B — Paste from clipboard (reliable fallback)

1. Create a blank **Automated cloud flow** with trigger
   **When a file is created (properties only)** (SharePoint).
2. Configure the trigger:
   - Site: `https://dpdhl.sharepoint.com/teams/EXP-USQIA-BS33384_AMKAD`
   - Library: `Shared Documents`
   - Folder: `.../2026 AMKAD Summaries/Daily/Cust_Sol Daily File`
3. Copy the entire contents of
   `docs/power_automate/custsol_flow_clipboard.json`.
4. In the designer: **+ New step** → **My clipboard** tab → press
   **Ctrl+V** → the pasted "AMKAD_Daily_CustSol_Import" scope appears →
   select it.
5. Do the **post-import fix-ups** (section 3).

### Option C — Manual build (nothing else worked)

Single linear chain, no branches:

| # | Action | Key settings |
|---|--------|--------------|
| 1 | **Trigger:** When a file is created (properties only) | Folder: `Daily/Cust_Sol Daily File` |
| 2 | Initialize variable `varMonthFolder` (String) | expression below |
| 3 | Initialize variable `varTargetFileName` (String) | expression below |
| 4 | Get files (properties only) | Folder: `.../Daily`, **Include Nested Items: Yes** |
| 5 | Filter array | From: `value` of step 4; where **Name contains** `Daily_Performance_Report` |
| 6 | Compose `LatestFile` | `last(sort(body('Filter_array'), 'Modified'))` |
| 7 | Get file content | Id: `outputs('Compose')?['{Identifier}']` |
| 8 | Create new folder | Path: `varMonthFolder` |
| 9 | Create file | Folder: `varMonthFolder`, Name: `varTargetFileName`, Content: output of 7. **Configure run after: run on both Succeeded and Failed of step 8** (the month folder already exists most days). |
| 10 | Run script — `custsol read daily` | File: **the trigger's** file Identifier (the source report) |
| 11 | Run script — `custsol append rawdata eur` | File: `Id` from step 9's output (today's new file). Parameter `sourceJson`: the `result` output of step 10. |

`varMonthFolder`:

```
concat('/Shared Documents/BS33384_AMKAD/Steering Board Presentations/2026/2026 Presentations/2026 AMKAD Summaries/Daily/', formatDateTime(convertFromUtc(utcNow(), 'Eastern Standard Time'), 'MMMM yyyy'))
```

`varTargetFileName`:

```
concat('Daily_Performance_Report- ', formatDateTime(convertFromUtc(utcNow(), 'Eastern Standard Time'), 'MMMM d yyyy'), '  AMKAD.xlsm')
```

(`convertFromUtc` keeps the date correct for files that land late in the
US evening, when UTC has already rolled to tomorrow.)

**Why this shape:** listing `Daily/` with nested items and taking the newest
`Daily_Performance_Report` by Modified date handles the month rollover
automatically — on August 1 the newest file is simply July 31's, no
"is this month's folder empty" branching needed. Get-content + Create-file
replaces SharePoint's Copy file action because Copy file cannot rename the
copy. Sorting is by **Modified**, never by filename — the
"July 13 2026"-style names don't sort chronologically as text.

---

## 3. Post-import fix-ups (Options A and B)

Connector dropdowns don't survive import as pickable values. Open each of
these actions and re-select from the pickers (the values to pick are already
in this doc):

1. **Trigger** (Option A only): site / library / the Cust_Sol Daily File folder.
2. **Get files (properties only)**: site, library, the `Daily` folder, and
   confirm **Include Nested Items = Yes**.
3. **Get file content / Create new folder / Create file**: re-pick the site;
   leave the expression-driven fields (folder, name, content) as imported.
4. **Both Run script actions**: pick Location (the SharePoint site), Document
   Library (`Shared Documents`), leave File as the imported expression, and
   select the script by name — the imported `scriptId` is a placeholder and
   **must** be replaced.

---

## 4. First live run

1. Save the flow, then drop a copy of a real daily report into
   `Cust_Sol Daily File` (or use **Test → Manually** with a recent trigger).
2. Check the run: every step green, and the last step's output shows
   `rowsAppended` equal to the number of data rows in the source file, with
   `skippedDuplicates: 0`.
3. Open today's new `Daily_Performance_Report- … AMKAD.xlsm`: new rows at the
   bottom of Raw Data EUR, columns F:P filled in by the table's own formulas.

---

## 5. Things to know / troubleshooting

- **Re-running on the same day**: the `Create file` step fails if today's
  output file already exists — delete today's file first, or point the last
  Run script step at the existing file manually. Even if the append runs
  twice, the script's duplicate guard skips rows already present
  (`skippedDuplicates` will be non-zero instead of double-counting).
- **Euro columns blank** (first 1–2 days after month close): the run does NOT
  fail. Rows are appended with the Euro fields empty and the final step's
  output shows `missingEuroRows` / `needsReview: true`. No alert is sent —
  check the run history if you care that day.
- **Source report never arrives**: the flow simply doesn't run and no file is
  created that day. This is by design (no alerting was wanted). If that
  changes, add a scheduled flow that checks the folder for today's file.
- **Year rollover (Jan 2027)**: the trigger folder and the two path
  expressions contain `2026` — update both the trigger's folder and the
  `varMonthFolder` expression when the 2027 folder structure is created.
- **"Table 'RawData' not found"**: someone renamed the table. Fix the name in
  Excel (Table Design → Table Name) or update `RAW_DATA_EUR_TABLE_NAME` in
  the writer script.
- **"Expected column … not found"**: a header was renamed in either file.
  The scripts match columns by header text on purpose (so column *reordering*
  never breaks anything) — renames need a matching one-word fix in the script.
- **Retesting after any script change**: run `tests/test_custsol.sh` — it
  recompiles both scripts and re-runs all 33 checks in seconds.

---

## 6. Expected headers (what the scripts match on)

**Source report** (single sheet):
`Report Date, countrycode, currencycode, Go Live Date, Customer Reporting
Name, Payment Terms, Total Payments, Total UAC, TDSO, TTLAR, Overdue,
GT60days, Gross Sales, GT90Days, GT60Days Euro, Gross Sales Euro, Company
Code, Total AREuro, Overdue AREuro, GT90Days Euro, Total UAC Euro, Total
Payments` — note "Total Payments" appears twice; the first is the native
amount, the last is the Euro amount, and the scripts handle that explicitly.

**Raw Data EUR table (`RawData`)** — columns written by the script:
`Date, Country, Customer, Go Live, Payment Term (days)` (A:E) and
`Total AR € (Live), Gross Sales € (Live), Overdue € (Live), > 60 days €
(Live), > 90 days € (Live), Total AR, Overdue, > 60 days, Gross Sales,
> 90 days, Total UAC € (Live), Total UAC, Total Payments € (Live), Total
Payments` (Q:AD). Columns F:P are never written — the table auto-fills its
own formulas there.
