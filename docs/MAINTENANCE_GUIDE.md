# AMKAD AR Report — Maintenance Guide

*A practical, no-prior-knowledge-required guide to keeping this automation running.*

If you are the person who now looks after this report, **read the first two
sections once**, then use the **Troubleshooting Playbook** as your day-to-day
reference. When something looks wrong, your **first stop is always the
Data Quality tab** of the report — it usually tells you what happened.

---

## 1. What this automation actually does

It turns a daily Excel file into an interactive HTML dashboard, automatically.

```
  Daily Excel file on SharePoint
  (Daily_Performance_Report- <Month Day Year> AMKAD.xlsm)
            │
            │  Power Query picks the LATEST file and reads its "RawData" table
            ▼
  RawData table in the working workbook
            │
            │  Office Script reads RawData, does the maths, builds JSON + HTML tables
            ▼
  JSON result ("result" object)
            │
            │  Power Automate injects the JSON into the HTML template
            ▼
  Finished HTML report  →  emailed / saved / shared
```

**The golden rule:** the report can only ever show what is in `RawData`.
If a number looks wrong, work **backwards** along that arrow chain — start at the
report, then RawData, then the source Excel file — until you find where the
number first goes wrong.

### Key idea: "latest snapshot per month" (MTD)

The daily Excel file's `RawData` table does **not** hold just today's numbers —
it holds a **running history**: many daily extracts for every month going back
to 2024. If you simply added all the June rows together you would count June
20–30 times over and every figure would be ~20× too big.

So the Office Script does this for each month:

1. Look at all the dated extracts in that month.
2. Pick the **latest date that is a real, full extract with actual money in it**.
3. Use **only** those rows. Everything from earlier days in the month is skipped.

This gives a true "as of the latest snapshot" (Month-to-Date) view.
You can see exactly what it picked in **Data Quality → Snapshot Audit**.

---

## 2. The pieces and where they live

Everything is in the Git repo (branch `claude/amkad-ar-mtd-filtering-ir3fmb`,
or whatever it has been merged into):

| File | What it is | Where it runs |
|---|---|---|
| `power_query/…` (the M query) | Picks the latest daily file, loads its `RawData` | Excel → Power Query (Data → Queries) |
| `office_scripts/amkad_ar_report.ts` | All the calculations + builds the tables/JSON | Excel Online → Automate → Code Editor |
| `power_automate/html_template.html` | The dashboard layout + charts + interactivity | Power Automate → the Compose/HTML action |
| `docs/MAINTENANCE_GUIDE.md` | This guide | — |

> **Deploying = copy the file's contents into the right tool and save/run.**
> There is no "build" step. See Section 4 for the exact clicks per piece.

---

## 3. Routine maintenance calendar

| When | Do this |
|---|---|
| **Every run / whenever numbers are reviewed** | Open the **Data Quality tab**. Confirm "Latest snapshot date used" is the date you expect, and that "Rows used" is a sensible number (thousands, not tens). |
| **At the start of each new year** | The Power Query path is hard-coded to `2026 AMKAD Summaries/Daily`. Update `2026` to the new year (see 4C). This is the single most common thing that will "suddenly break" the whole report. |
| **If the daily file naming changes** | Update the filename parsing in Power Query (see 4C). |
| **If a new column is added to RawData / a column is renamed** | Check the field mapping the Office Script uses (see 4B, "Field mapping"). |
| **If the report layout needs a change** | Edit `html_template.html` and re-paste into Power Automate (4A). |

---

## 4. How to update each component

### 4A. The HTML dashboard (`power_automate/html_template.html`)

Change this when you want to change **how the report looks or behaves**
(colours, wording, tabs, charts).

1. Edit `power_automate/html_template.html` in the repo.
2. Copy the **entire** file contents.
3. In Power Automate, open the flow → find the action that holds the HTML
   (a Compose or "HTML template" action) → paste over its contents → **Save**.
4. Run the flow once and open the output to check.

**Do not touch anything that looks like** `@{outputs('Run_script')?['body/result']?['…']}`.
Those are the placeholders Power Automate fills with real numbers. Deleting or
renaming one will blank out that value. Everything else is normal HTML/CSS/JS.

> **Why it must be self-contained:** the report often opens inside SharePoint or
> Outlook preview, which **blocks anything loaded from the internet** (external
> fonts, chart libraries, images). That is why all styling and charts are written
> inline. Never add a `<script src="https://…">` or external stylesheet — it will
> silently fail in those viewers.

### 4B. The Office Script (`office_scripts/amkad_ar_report.ts`)

Change this when the **calculations, tables, or the data it reads** need to change.

1. Edit `office_scripts/amkad_ar_report.ts` in the repo.
2. Copy the **entire** file.
3. Excel Online → **Automate** tab → open the AMKAD script → **Code Editor** →
   select all → paste → **Save** → **Run** once to confirm no errors.
4. The flow calls this same script, so Power Automate picks up the change
   automatically on its next run.

**Field mapping (important):** the script matches RawData columns by name
(e.g. `Total AR € (Live)`, `Overdue € (Live)`). If someone **renames a column**
in the source, the script may read blanks (→ zeros). Look near the top of the
script for the config/mapping section and update the expected column name to
match the new header. Symptom of a broken mapping: **one metric goes to €0 for
every month** (not just one month).

### 4C. The Power Query (loads the latest daily file)

Change this at **year rollover**, or if the **file naming pattern changes**.

- **Year rollover:** update `"2026 AMKAD Summaries/Daily"` to the new year.
- **Filename format:** the files are named
  `Daily_Performance_Report- June 30 2026 AMKAD.xlsm`. The query parses the
  `Month Day Year` from the name to find the newest one. If the format changes
  (different word order, no `AMKAD` suffix, etc.), update the parsing step.
- **To check which file it is using:** temporarily change the query's final line
  from `RawDataTable` to `Ranked` and look at the top row's `Name` and `FileDate`.
  Change it back when done.

> Selecting the latest file by the **date written in its name** is deliberate —
> it is more reliable than SharePoint's "Date created", which resets whenever a
> file is copied or re-uploaded and can make an old/template file look "newest".

### 4D. The automated email (`power_automate/email_body.html`)

The flow ends with a **Send an email (V2)** action. Its Body is **not** the
full interactive report — email clients (especially Outlook) strip `<script>`
and most CSS, so the tabs/charts/dark-mode would render broken. Instead the
email is a short, email-safe KPI summary with a button linking to the real
report file.

**To update the email:**

1. Edit `power_automate/email_body.html` in the repo.
2. Copy the **entire** file.
3. In Power Automate, open **Send an email (V2)** → click into **Body** →
   click the **`</>` (code view)** button → select all → paste → **Save**.

**How the link works:** the button and the plain-text fallback link both use
an expression, not the dynamic-content picker:

```
@{outputs('Create_sharing_link_for_a_file_or_folder')?['body/link/webUrl']}
```

That pulls the URL straight from the **"Create sharing link for a file or
folder"** action's output. If that action is ever renamed, this expression
must be updated to match (spaces in the action name become underscores).

**If the link doesn't work:**
- Run the flow once, open **run history → "Create sharing link…" → Outputs**,
  and confirm `link.webUrl` actually has a value. If it's empty, the problem
  is that action, not the email.
- If recipients get "access denied" when they click it, check that action's
  **Scope** setting — "Organization" lets anyone at DHL open it; "Specific
  people" restricts it.
- Power Automate's code-view editor **strips some HTML on save** (comments,
  some attributes). Always re-paste the **whole file** rather than editing a
  fragment in place, so nothing drifts out of sync with what's in the repo.

---

## 5. Troubleshooting Playbook

**Always start at the report's Data Quality tab.** Then match your symptom below.

### Symptom: every number is ~20× too big / totals look insane
- **Cause:** the latest-snapshot filter isn't working, so multiple daily extracts
  in a month are being summed.
- **Check:** Data Quality → "Rows used" will be very large; Snapshot Audit shows
  many dates being combined.
- **Fix:** make sure the **current** Office Script is deployed (this was the
  original bug that the snapshot logic fixed).

### Symptom: one month is blank / all 0.0% (the "April" case)
- **Cause:** that month has **no real money in the source** — the extract rows
  exist but the `Total AR € (Live)` column is blank/zero for that month.
- **Check:** Data Quality → **Snapshot Audit** → find the month. If every date
  shows `€0`, the data genuinely isn't there.
- **Fix:** this is a **source-data problem, not a report bug**. Open the daily
  `.xlsm`, load/inspect `RawData`, filter to that month, and look at the live
  EUR columns. If they're empty there too, the fix is upstream — whatever
  generates the daily file didn't write that month's values.
- **Note:** the current script automatically **hides** a fully-empty month so it
  doesn't show as a dead column. If you still see a dead column, the script
  isn't the latest version — redeploy it (4B).

### Symptom: one metric (e.g. UAC) is €0 for EVERY month
- **Cause:** a **column was renamed** in RawData, so the script can't find it.
- **Fix:** update the field mapping in the Office Script (4B).

### Symptom: the whole report failed to refresh / Power Query errors
- **Cause (most common):** year rollover — the folder path still says the old year.
- **Other causes:** the naming pattern changed; no files matched; SharePoint
  permissions/sign-in expired.
- **Fix:** update the path/parsing (4C); confirm you can open the SharePoint
  folder yourself; re-authenticate the SharePoint connection.

### Symptom: charts are blank and there's a red "issue running scripts" banner
- **Cause:** you're viewing inside a sandboxed preview (SharePoint/Outlook) that
  restricts scripts, **or** someone added an external resource to the HTML.
- **Fix:** open the report in a real browser tab to confirm it works there. If it
  only breaks in preview, check that no `https://` script/stylesheet/font/image
  was added to the template (4A, "self-contained").

### Symptom: tables have no column headers / a section is empty
- **Cause:** an Office Script table builder is missing its `<thead>` or its
  underlying data list is empty.
- **Fix:** confirm the latest script is deployed. If a specific list is empty
  (e.g. Action Required), check whether any rows actually meet the thresholds in
  the source for that month.

### Symptom: "latest month" is wrong / an old month is showing as newest
- **Cause:** Power Query selected the wrong daily file (often a copied/re-uploaded
  file with a misleading timestamp).
- **Fix:** use the "which file is it using?" check in 4C; confirm the newest file
  in the folder is named with the correct date.

### Symptom: the "Open the report" button/link in the email doesn't show or doesn't work
- **Cause (most common):** Power Automate's code-view editor stripped some HTML on
  save (this happens even when nothing looks wrong at a glance).
- **Cause:** the sharing-link action produced no URL, or was renamed so the email's
  expression no longer points at it.
- **Fix:** re-paste the **entire** `power_automate/email_body.html` file into the
  email Body (don't hand-edit a fragment). Then check run history →
  **"Create sharing link…" → Outputs** to confirm `link.webUrl` has a real value.
  See 4D for the full explanation.

---

## 6. Before you change anything — a safety checklist

- [ ] Copy the current working version somewhere before overwriting (the repo Git
      history is your safety net — commit before and after changes).
- [ ] Change **one** component at a time, then re-run and check the report.
- [ ] After any change, open the **Data Quality tab** and sanity-check
      "Latest snapshot date used" and "Rows used".
- [ ] Keep the `@{outputs('Run_script')…}` placeholders intact in the HTML.
- [ ] Never add internet-loaded resources to the HTML.

---

## 7. Glossary

| Term | Plain meaning |
|---|---|
| **RawData** | The big table inside the daily Excel file that holds all the numbers, for all months, for many days. Everything is derived from this. |
| **Snapshot / extract** | One day's capture of the data. A month has many. |
| **MTD (Month-to-Date)** | Using only the latest snapshot in a month, so figures are a point-in-time balance, not a sum of every day. |
| **Office Script** | Microsoft's TypeScript-based automation that runs inside Excel Online. Does the calculations here. |
| **Power Automate** | Microsoft's workflow tool. Runs the script and drops the results into the HTML. |
| **Power Query** | Excel's data-loading tool. Fetches the latest daily file from SharePoint. |
| **Placeholder** (`@{outputs(...)}`) | A slot in the HTML that Power Automate replaces with a real value/table at run time. |
| **UAC** | Unapplied Cash — money received but not yet matched to an invoice. |
| **Live EUR columns** | The `… € (Live)` fields; these are the only monetary values the report uses (DSO/TDSO are excluded on purpose). |

---

## 8. Escalation / hand-off notes

*(Fill these in for the next person.)*

- **Report owner / business contact:** _______________________
- **SharePoint site owner (for access issues):** _______________________
- **Who produces the daily `.xlsm` file (for source-data gaps like empty months):** _______________________
- **Where the flow lives (Power Automate environment + flow name):** _______________________
- **Git repo / branch:** `jkadadihi/Power-BI-` — branch `claude/amkad-ar-mtd-filtering-ir3fmb`

---

*When in doubt: the report shows what's in RawData; RawData comes from the latest
daily file; the Data Quality tab tells you what the script actually used. Follow
that chain and you'll find almost any problem.*
