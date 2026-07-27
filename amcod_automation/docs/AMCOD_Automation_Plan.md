# AMCOD Month-End Automation - Plan

## Why this exists

Marcia Lopez Espinoza runs the AMCOD month-end reporting process for Global
reporting almost entirely by hand: copy/paste between workbooks, manual
sign-flips on payment data, manual Amazon exclusion, manual DSO/TDSO/SPR
cross-checks, and manual BRM re-keying. The process works, but it lives in
one person's head and one person's calendar - if Marcia is out sick during
close week, the process stalls.

The goal is **not** to rebuild AMCOD or replace the workbooks. It's to strip
the repetitive, deterministic, error-prone steps out of Marcia's hands and
into Power Query / Office Scripts / Power Automate, so:

- the manual copy/paste between Debits -> Raw Data EU disappears
- Amazon exclusion and payment sign-flips can't be gotten wrong or forgotten
- DSO/TDSO/SPR reconciliation is checked automatically instead of by eye
- a missed exchange-rate upload gets flagged *before* it blanks out
  downstream reporting, instead of being discovered after the fact
- BRM re-keying becomes a load step instead of a copy/paste step

## Current flow (as demonstrated in the "[EOM #s due]" meeting)

```
SAP / SQL Sources
  -> Debits File (post month-end close)
  -> Raw Data EU tab (manual copy/paste)          <- Stage 3, biggest manual cost
  -> Monthly Workbook (pivots, DSO/SPR, Gross Sales)
  -> AMCOD Workbook (BRM metrics pasted in)        <- Stage 9, second biggest manual cost
  -> Cockpit (imports prior month, adds current)
  -> Global Submission Email
```

Supporting inputs: SIP file (downloaded monthly), BRM file (Gross Sales, Open
AR, 60+, 90+, DSO), and a separate exchange-rate maintenance step that, if
missed, blanks daily reporting files.

## Automation priorities and what each deliverable does

| # | Manual stage | Deliverable in this repo | Status |
|---|---|---|---|
| 1 | Roll-forward + Debits -> Raw Data EU (Stages 1 & 3) | `power_query/01_debits_to_raw_data_eu.m` + `office_scripts/AppendCurrentMonthToRawData.ts` + `power_automate/monthly_rollforward_flow.md` | Built |
| 2 | Remove Amazon (Stage 5) | `power_query/02_remove_amazon.m` | Built |
| 3 | Negative payment cleanup (Stage 4) | `power_query/03_convert_negative_payments.m` | Built |
| 4 | DSO/TDSO/SPR validation (Stage 7) | `office_scripts/ValidationDashboard.ts` | Built |
| 5 | Exchange-rate monitoring | `office_scripts/ExchangeRateMonitor.ts` + `power_automate/exchange_rate_monitor_flow.md` | Built |
| 6 | BRM integration (Stage 9) | `power_query/04_brm_integration.m` | Built |

Priorities 1-3 are implemented as Power Query `.m` scripts because that step
is Excel-native, deterministic, and exactly what Power Query is for: pull
raw source data, transform it the same way every time, refresh with one
click instead of one keyboard's worth of copy/paste.

Priorities 4-5 are implemented as Office Scripts because they need to *run
themselves* on a schedule (via Power Automate) rather than wait for Marcia to
open the workbook - a missed exchange-rate upload is a problem specifically
because nobody noticed until it was too late.

Priority 6 (BRM) is Power Query again, since it is the same shape of problem
as Priority 1: a structured export, copied by hand, into a fixed set of
report cells.

## How to adopt these

### Priority 1 - the Design B model (validated against the real workbook)

Priority 1 is now three pieces working together
(`power_query/01_debits_to_raw_data_eu.m` +
`office_scripts/AppendCurrentMonthToRawData.ts` +
`power_automate/monthly_rollforward_flow.md`), built around how Marcia
actually keeps the files:

- **Each monthly workbook holds the full 2023 -> current history**, carried
  forward every time the file is duplicated for a new month. The history is
  already inside the file, so nothing has to rebuild it and the old Debits
  source files are never needed - only the current month's Debits file has
  to exist.
- **The query pulls the current month only**, into a staging table
  `Staging_CurrentMonth`. It does not recurse history and does not overwrite
  RawData. The month is auto-derived as *last* month (a month closes during
  the following month), so there is no parameter to advance - just refresh.
  Only `SharePointSite_Url` is a parameter.
- **The Office Script appends the staging rows to the bottom of RawData as
  values**, so they become permanent history the next duplicate carries
  forward. It is re-run safe (skips a month already present) and re-asserts
  the RawData formula columns (DSO, TDSO Gap, %) onto the new rows so they
  calculate.
- **The Power Automate flow** duplicates last month's file into the new
  month's folder, sets the month parameter, refreshes, and runs the append -
  automating Stages 1 and 3 in one pass.

Confirmed facts baked into these files:

- `VW_AMKAD_Source_Debits` is a loaded query **table** inside the Debits
  workbook, not a worksheet tab. Filtering on `[Kind] = "Sheet"` matches
  nothing - match on `[Item]` alone.
- The Debits file keeps the **same name every month**; only its folder
  changes (`.../Qtr 2/June 2026/`), so the month is derived from the folder
  path, not the filename.
- Debits carries **both currency sets**: G-M in EUR (-> RawData's
  `€ (Live)` columns) and N-T in local currency (-> RawData's bare columns).
  Both are needed. Pulling only EUR leaves the local columns blank *and*
  silently breaks DSO.
- Debits spells the local-currency column **`overdue` in lowercase** while
  the EUR one is `Overdue €`. Power Query column names are case-sensitive,
  so this needs an explicit normalize step or it yields nulls.
- RawData's header strings are not what you would guess: `Onboard Dt` (not
  "Onboard Date"), `Payment Term (days)`, `Gross Sales € (Live)`, and a
  **space after `>`** in the day-bucket columns. The append script therefore
  matches headers on a normalized key rather than exact text.
- **Payments arrive negative from SAP** (-1923) and are reported positive
  (1923) - Stage 4 "Payment Cleanup" is folded into the query.
- DSO is a RawData formula
  (`Total AR / Gross Sales * VLOOKUP(Month, CD3Mth, 3, 0)`) that reads the
  **local-currency** columns. It is never written by the query or script; it
  computes on the appended rows once those columns are populated. A DSO of 0
  on a new row means the local-currency columns came through blank.

Validated end-to-end against real June 2026 data: every column of the
appended row matches the manually-produced row, including DSO.

Remaining before the first fully automated run:
- Whether the Excel Online connector can refresh a Power Query connection
  directly, or whether the flow needs a small refresh script alongside the
  append script.

`02_remove_amazon.m`, `03_convert_negative_payments.m`, and
`04_brm_integration.m` (Priorities 2, 3, 6) are still built against assumed
column names - validate those the same way (walk the script against a real
BRM export and a real Raw Data No Amazon tab with Marcia) before relying on
them.

## Phased rollout

**Phase 1 (this change).** Ship the Power Query transforms for Debits ->
Raw Data EU, Amazon removal, and payment sign-flip, plus the validation and
exchange-rate Office Scripts, as standalone artifacts Marcia (or IT) can
wire into the existing Monthly Workbook.

**Phase 2.** Wire `power_automate/exchange_rate_monitor_flow.md` into an
actual Power Automate flow with real recipients, and point the Debits
Power Query at the real monthly folder structure so refresh is one click at
month-end instead of a manual file-open.

**Phase 3.** Extend the BRM Power Query into the AMCOD workbook itself
(replacing Stage 9's copy/paste), then look at Cockpit submission
(Stage 12-13) once the upstream stages are trusted running unattended.

## What's deliberately out of scope for this change

- Rebuilding Cockpit or the Global submission email - too far downstream to
  automate safely before the upstream stages are proven.
- FTE maintenance (Stage 10) and monthly archive tab copy (Stage 11) - low
  transformation complexity, more of a template/process fix than a data
  automation, best handled after Phase 1 is validated in production.
- Gross Sales 3-month-average legacy tabs - Marcia flagged these as possibly
  obsolete; recommend confirming with Global before automating a calculation
  that may not be needed anymore.
