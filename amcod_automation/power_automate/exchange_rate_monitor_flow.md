# Power Automate flow: Exchange Rate Monitor

Wraps `office_scripts/ExchangeRateMonitor.ts` so the missed-deadline scenario
Marcia described - rates not loaded, daily reporting files going blank,
having to manually recalculate afterward - gets caught automatically instead
of after the fact.

## Trigger

**Recurrence** - daily, starting a few business days before the exchange-
rate deadline (Marcia/Global know the exact cutoff; configure the
recurrence to start covering that window, e.g. the 20th through end of
month).

## Steps

1. **Recurrence trigger** (daily, business hours only).
2. **Compose** - build the current month column label to check, e.g.
   `formatDateTime(utcNow(), 'MMM-yy')`. Passing this in explicitly (rather
   than relying on the script's own fallback) keeps the check correct
   regardless of when in the day the flow runs.
3. **Excel Online (Business) - Run script**
   - Location/Document library/File: the Monthly Workbook (or a dedicated
     Exchange Rate workbook, if rates are maintained separately).
   - Script: `ExchangeRateMonitor.ts`.
   - Parameter `monthColumnOverride`: output of the Compose step.
4. **Condition** - `length(outputs('Run_script')?['body/missingCurrencies']) is greater than 0`.
   - **If yes:** Post an adaptive card / message to Marcia (and backup
     contact) via Teams, and/or send an email, listing the missing
     currencies and the deadline. Include the workbook link so the fix is
     one click away.
   - **If no:** end the run silently - no message, no noise. (A daily
     "all good" ping trains people to ignore the flow; only alert on an
     actual gap.)

## Suggested notification copy

> Exchange rates for **{monthColumn}** are missing for: **{missingCurrencies}**.
> Deadline: {deadline}. Daily reporting will go blank for these currencies
> until rates are loaded. [Open workbook]({workbook link})

## Escalation

If the condition is still true within 1 business day of the deadline,
consider a second branch that also notifies a backup/manager - this is the
scenario that caused the manual-recalculation rework Marcia mentioned, and
it's exactly the kind of single-point-of-failure risk this automation
exists to remove.

## Notes

- This flow only *detects* missing rates; it does not load them. Loading
  the rates is still a manual/BRM-sourced step for now - automating the
  load itself is a good Phase 2/3 follow-up once the source of the rate
  feed (manual entry vs. an exportable file) is confirmed.
- Keep the recurrence inside business hours so a Teams ping doesn't land at
  2am and get missed before the deadline anyway.
