# Cust Sol Daily Import — the actual plan

One sentence: **your existing AMKAD input file's query gets 8 extra steps
that pull in today's new rows, and Power Automate saves a copy of that same
file under today's date.** No Office Script. No second "engine" file. No
new mechanism — this builds on the query you already have running.

---

## Why this works (the one fact that matters)

Power Query always **replaces** a table's contents on refresh — it can never
add rows to what's already sitting there. So a query cannot "keep growing"
a table in place.

But your **existing** query doesn't have that problem, because it doesn't
grow anything in place either — it finds the newest `Daily_Performance_Report`
file each time and **reads that file's `RawData` table fresh**. Since each
daily file already contains full history (because it was itself built from
the *previous* day's file the same way), reading "the newest file" gives you
"all history so far" — every time, safely.

So all we're doing is extending that same idea by one more step: after your
query finds full history-through-yesterday, also read today's new source
file and stack those rows on top. The result is history-through-today. Save
that as today's file, and tomorrow the cycle repeats automatically.

```
Your existing query:
  find newest Daily_Performance_Report file → read its RawData table
                                                = history through YESTERDAY

              + (added) find newest Cust_Sol source file → map its columns
                                                = TODAY's new rows

  stack them, drop exact duplicates
                                                = history through TODAY
                                                       │
                                                       ▼
                                    Power Automate saves this file as
                              "Daily_Performance_Report- <today>  AMKAD.xlsm"
```

---

## 1. Update the query (5 minutes, in the file you already have)

1. Open your **current AMKAD input file** — the one whose query you pasted
   earlier (the one that already finds the latest `Daily_Performance_Report`).
2. **Data → Queries & Connections** → right-click the **RawData** query →
   **Edit**.
3. **Home → Advanced Editor**.
4. Select all, delete, and paste in the full contents of
   `power_query/custsol_rawdata.m`.
   - The first ~40 lines are **word-for-word your existing query** — nothing
     about how it finds the latest file changed.
   - Everything after the `// ===== NEW =====` comment is the addition: find
     today's Cust_Sol source file, map its 19 columns to match, stack on top,
     drop any exact repeat.
5. **Done**. Click **Close & Load**.
6. Refresh once manually (**Data → Refresh All**) and check: the table now
   has today's rows at the bottom in addition to everything it had before.

That's the entire data-side change. Everything else about the file — its
formulas, other sheets, other queries — is untouched.

---

## 2. The one Power Automate step: save today's copy

This is the **only** automation piece, and it's unavoidable — nothing inside
Excel/Power Query can create a brand-new file under a new name; only
something outside Excel can do that.

**Add this to the end of your existing open-on-schedule flow** (the one that
already opens this file daily so the query refreshes):

| # | Action | Settings |
|---|--------|----------|
| 1 | Initialize variable `varTargetFile` (String) | `concat('Daily_Performance_Report- ', formatDateTime(convertFromUtc(utcNow(), 'Eastern Standard Time'), 'MMMM d yyyy'), '  AMKAD.xlsm')` |
| 2 | Initialize variable `varMonthFolder` (String) | `concat('/Shared Documents/BS33384_AMKAD/Steering Board Presentations/2026/2026 Presentations/2026 AMKAD Summaries/Daily/', formatDateTime(convertFromUtc(utcNow(), 'Eastern Standard Time'), 'MMMM yyyy'))` |
| 3 | **Get file content** (SharePoint) | your AMKAD input file (the one you just refreshed) |
| 4 | **Create new folder** (SharePoint) | Path: `varMonthFolder` |
| 5 | **Create file** (SharePoint) | Folder: `varMonthFolder`, Name: `varTargetFile`, Content: output of step 3. **Configure run after** step 5 to run whether step 4 succeeds or fails (the month folder already exists most days — only the 1st of the month needs it created). |

**Why "Get content + Create file" instead of "Copy file":** SharePoint's
Copy file action can't rename the copy — Get content + Create file is what
lets the new copy be named with today's date.

That's it. Five actions, all mechanical, none of them touch the data logic —
that's entirely Power Query's job now.

---

## 3. Things to know

- **Re-running the same day:** `Table.Distinct` on Date+Country+Customer
  means if the query or the flow runs twice in one day, nothing doubles up —
  the second run's identical rows are dropped.
- **Euro columns blank early in the month:** those rows load with blank Euro
  fields; nothing fails. No alert is sent, per your call earlier.
- **Source file missing that day:** `TodayRaw` resolves to `null`, `TodayRows`
  becomes empty, and the query still succeeds — it just carries forward
  yesterday's history with nothing new added, so no gap-day crash.
- **Table name:** must stay literally **`RawData`** — that's what your
  existing downstream logic (and this query) looks for.
- **Year rollover (Jan 2027):** the folder path in step 2 above, and the
  `"2026 AMKAD Summaries/Daily"` filter inside the query, both hardcode 2026
  — update both when the 2027 folders exist.

---

## Retired alternatives

Two earlier designs are archived in `docs/archive/` — a separate "engine
workbook" (Power Query) version and an Office Script append version. Both
were superseded by this one because it needed neither: no new file, no
custom script, just an extension of the query you already had.
