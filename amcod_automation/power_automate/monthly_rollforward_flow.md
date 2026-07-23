# Power Automate flow: Monthly Roll-Forward + Debits Append

Automates the AMCOD month roll-forward (Stage 1) and the Raw Data EU update
(Stage 3) as one flow, using the Design B model:

- each monthly workbook already carries the full 2023 -> prior-month history
- the flow duplicates last month's file into the new month's folder (history
  rides along), points the query at the new month, refreshes, and appends
  the new month's Debits block to RawData **as values** so it becomes
  permanent history for next time.

Old Debits source files are never needed - only the current month's Debits
file has to exist in the new month's folder.

## Pieces this flow ties together

| Piece | File in this repo |
|---|---|
| Pull current month's Debits | `power_query/01_debits_to_raw_data_eu.m` (query `Debits_CurrentMonth`, loaded to table `Staging_CurrentMonth`) |
| Append to RawData as values | `office_scripts/AppendCurrentMonthToRawData.ts` |

## Trigger

**Manual** (a "Run" button) for the first few months while you build trust,
then switch to **Recurrence** timed a day or two after Global's month-close
deadline once you're confident all countries have closed and the Debits file
is present.

## Steps

1. **Trigger** (manual button, or recurrence).
2. **Compose - PriorMonthName** and **Compose - NewMonthName**
   - Derive both from `utcNow()` (or from a flow input), e.g. `June 2026` and
     `July 2026`, in the same "MMMM yyyy" form as the folder names.
   - Also compose the **new month's folder path** and **quarter folder** so
     the copy lands in the right place
     (`.../<Year> AMKAD Summary/Qtr <n>/<NewMonthName>/`).
3. **SharePoint - Get file content** (or **Find files in folder**) for last
   month's Monthly workbook (`Monthly_Performance_Report... <PriorMonthName>`).
4. **SharePoint - Create new folder** for `<NewMonthName>` (if it doesn't
   already exist).
5. **SharePoint - Create file** - write last month's workbook content into
   the new month's folder under the new name. *(This is the duplicate - the
   whole history comes with it.)*
6. **Excel Online (Business) - Update a row / Set parameter**
   - Set the workbook's `ReportMonth_FolderName` parameter to `<NewMonthName>`
     so the copied query pulls the NEW month, not the month it was copied
     from. If setting a Power Query parameter directly isn't available in your
     Excel connector, store the month in a named cell the query reads instead,
     and have this step write that cell.
7. **Excel Online (Business) - Run script: refresh**
   - Refresh `Debits_CurrentMonth` so `Staging_CurrentMonth` holds the new
     month's rows. (Use a small "refresh all" Office Script if the connector's
     native refresh doesn't cover Power Query.)
8. **Excel Online (Business) - Run script: `AppendCurrentMonthToRawData`**
   - Appends the staging rows to the bottom of RawData as values. The script
     is re-run safe: if the month is already present it skips, so an accidental
     double-run won't duplicate rows.
9. **Excel Online (Business) - Run script (optional): clear staging** so the
   staging table doesn't linger with this month's rows.
10. **(Optional) Notify** - post a Teams message / email: "July 2026 file
    created and Raw Data EU updated - N rows appended," using the script's
    returned `appendedRows` / `month`.

## Order matters

Refresh (step 7) must run **before** append (step 8): the append reads
whatever is currently in the staging table. The dedupe guard in the script
protects against re-running append, but it can't protect against appending a
stale staging table - so always refresh first.

## What this flow deliberately does NOT do

- It does not touch DSO/TDSO/% or the local-currency columns - those are
  RawData formulas that auto-fill on the new rows.
- It does not pull Amazon-excluded or payment-sign-fixed data - those are
  separate downstream steps (`02_remove_amazon.m`,
  `03_convert_negative_payments.m`) that read RawData after this append.
- It does not submit anything to Cockpit or Global - out of scope until the
  upstream stages are trusted running unattended.
