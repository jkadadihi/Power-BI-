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
| 1 | Debits Workbook -> Raw Data EU (Stage 3) | `power_query/01_debits_to_raw_data_eu.m` | Built |
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

`power_query/01_debits_to_raw_data_eu.m` (Priority 1) has been validated
against the real structure:

- Source is `AMKAD_KPI_ViewRefreshable_Debits Only...` on SharePoint, inside
  a dated month folder (`.../2026 AMKAD Summary/Qtr 2/June 2026/`) that gets
  recreated every month per Stage 1. The query points at the stable parent
  folder and lets `SharePoint.Files` recurse through every Qtr/Month
  subfolder, so a new month needs zero query edits - only the year segment
  in `DebitsRoot_RelativePath` needs updating, once a year.
- Real Debits columns (A-T) and the real Raw Data EU / Monthly file target
  columns (confirmed with Marcia's actual screenshots) are mapped 1:1 in
  the script's header comment. Only columns with a genuine Debits source
  are populated (Month, Country, Customer, Onboard Date, Payment days,
  DSO, and the 7 EUR/"Live" metrics) - TDSO, TDSO Gap, the %-columns, and
  the local-currency columns are intentionally left untouched since they
  are not sourced from Debits.
- One remaining unknown: the exact sheet name inside the Debits workbook
  (the script currently just takes the first worksheet). Confirm and set
  explicitly once known.
- Raw Data EU appears to mix manually-formatted/formula columns (TDSO Gap,
  %-columns) with columns that would come from this query. Since Power
  Query owns every column of whatever table it's loaded into, load this
  query's output to a separate staging table (e.g. "Debits Live Feed")
  rather than directly overwriting Raw Data EU, and have Raw Data EU's
  EUR/"Live" columns reference the staging table by formula (or confirm
  with Marcia/IT that Raw Data EU can be restructured to be fully
  query-owned instead).

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
