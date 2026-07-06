#!/usr/bin/env python3
"""Generate the AMKAD AR Maintenance Guide as a styled Word document."""
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

RED = RGBColor(0xC4, 0x04, 0x0D)
DHL_RED = RGBColor(0xD4, 0x05, 0x11)
INK = RGBColor(0x19, 0x15, 0x12)
SOFT = RGBColor(0x5F, 0x58, 0x52)
MUTED = RGBColor(0x8A, 0x83, 0x7B)
GOOD = RGBColor(0x0F, 0x7A, 0x2A)
WARN = RGBColor(0x9A, 0x62, 0x08)
CRIT = RGBColor(0xC0, 0x2A, 0x2A)
INFO = RGBColor(0x2A, 0x6F, 0xB0)

doc = Document()

# ---- base styles ----
normal = doc.styles["Normal"]
normal.font.name = "Calibri"
normal.font.size = Pt(10.5)
normal.font.color.rgb = INK
normal.paragraph_format.space_after = Pt(7)
normal.paragraph_format.line_spacing = 1.12

for lvl, size in [("Heading 1", 17), ("Heading 2", 13), ("Heading 3", 11.5)]:
    st = doc.styles[lvl]
    st.font.name = "Calibri"
    st.font.size = Pt(size)
    st.font.bold = True
    st.font.color.rgb = INK

def shade(cell_or_para, hex_color):
    """Apply background shading to a table cell or paragraph."""
    el = cell_or_para._tc if hasattr(cell_or_para, "_tc") else cell_or_para._p
    pPr = el.get_or_add_tcPr() if hasattr(cell_or_para, "_tc") else el.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hex_color)
    pPr.append(shd)

def left_border(cell, hex_color, size="24"):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    lb = OxmlElement("w:left")
    lb.set(qn("w:val"), "single"); lb.set(qn("w:sz"), size); lb.set(qn("w:space"), "0"); lb.set(qn("w:color"), hex_color)
    borders.append(lb); tcPr.append(borders)

def eyebrow(text):
    p = doc.add_paragraph()
    r = p.add_run(text.upper())
    r.font.size = Pt(8); r.font.bold = True; r.font.color.rgb = RED
    rPr = r._element.get_or_add_rPr(); sp = OxmlElement("w:spacing"); sp.set(qn("w:val"), "40"); rPr.append(sp)
    p.paragraph_format.space_after = Pt(2)
    return p

def body(text, color=INK, size=10.5, italic=False, after=7):
    p = doc.add_paragraph()
    add_rich(p, text, color, size, italic)
    p.paragraph_format.space_after = Pt(after)
    return p

def add_rich(p, text, color=INK, size=10.5, italic=False):
    """Very small **bold** and `code` parser."""
    import re
    parts = re.split(r"(\*\*.*?\*\*|`.*?`)", text)
    for part in parts:
        if part.startswith("**") and part.endswith("**"):
            r = p.add_run(part[2:-2]); r.bold = True
        elif part.startswith("`") and part.endswith("`"):
            r = p.add_run(part[1:-1]); r.font.name = "Consolas"; r.font.size = Pt(9.5); r.font.color.rgb = RGBColor(0x3A,0x33,0x2D)
        else:
            r = p.add_run(part)
        r.font.size = Pt(size); r.font.color.rgb = color; r.italic = italic

def callout(title, text, accent=DHL_RED):
    t = doc.add_table(rows=1, cols=1); t.alignment = WD_TABLE_ALIGNMENT.LEFT
    cell = t.cell(0, 0)
    shade(cell, "FBF7EA" if accent == RGBColor(0xFF,0xCC,0x00) else "FBF3F3")
    left_border(cell, "%02X%02X%02X" % (accent[0], accent[1], accent[2]))
    p = cell.paragraphs[0]
    rt = p.add_run(title + "  "); rt.bold = True; rt.font.color.rgb = INK; rt.font.size = Pt(10.5)
    add_rich(p, text, INK, 10.5)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def styled_table(headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.LEFT
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        shade(hdr[i], "F2EFEA")
        pr = hdr[i].paragraphs[0]; run = pr.add_run(h.upper())
        run.bold = True; run.font.size = Pt(8.5); run.font.color.rgb = SOFT
    for row in rows:
        cells = t.add_row().cells
        for i, val in enumerate(row):
            cells[i].paragraphs[0].paragraph_format.space_after = Pt(2)
            add_rich(cells[i].paragraphs[0], val, INK, 9.5)
    if widths:
        for i, w in enumerate(widths):
            for r in t.rows:
                r.cells[i].width = Inches(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t

def tcard(sev_label, sev_color, sev_fill, symptom, cause, fix, check=None):
    t = doc.add_table(rows=1, cols=1)
    cell = t.cell(0, 0)
    left_border(cell, "%02X%02X%02X" % (sev_color[0], sev_color[1], sev_color[2]), size="36")
    shade(cell, "FFFFFF")
    # severity + symptom
    p = cell.paragraphs[0]
    s = p.add_run(sev_label.upper() + "   "); s.bold = True; s.font.size = Pt(8); s.font.color.rgb = sev_color
    sym = p.add_run(symptom); sym.bold = True; sym.font.size = Pt(11); sym.font.color.rgb = INK
    def line(label, txt):
        pp = cell.add_paragraph(); pp.paragraph_format.space_after = Pt(2)
        lr = pp.add_run(label + "  "); lr.bold = True; lr.font.size = Pt(8.5); lr.font.color.rgb = MUTED
        add_rich(pp, txt, SOFT, 9.8)
    if check: line("CHECK", check)
    line("CAUSE", cause)
    line("FIX", fix)
    doc.add_paragraph().paragraph_format.space_after = Pt(3)

# ================= CONTENT =================
eyebrow("DHL Express · AMKAD Finance")
h = doc.add_heading("AMKAD AR Report — Maintenance Guide", level=0)
for r in h.runs: r.font.color.rgb = INK
sub = body("Everything you need to keep the accounts-receivable dashboard running, written for "
           "whoever inherits it next. Skim sections 1–2 once, then live in the Troubleshooting playbook.",
           SOFT, 11)
meta = body("Repo: **jkadadihi/Power-BI-**   ·   Branch: **claude/amkad-ar-mtd-filtering-ir3fmb**   ·   "
            "First stop for any problem: **the Data Quality tab**", SOFT, 9.5)

# Section 1
eyebrow("Section 01"); doc.add_heading("What this automation does", level=1)
body("It turns a daily Excel file into an interactive HTML dashboard — automatically, with no manual "
     "steps in between. Data flows one direction:")
styled_table(["Step", "Component", "What happens"],
    [["1 · Source", "Daily Excel file", "On SharePoint, named `…- June 30 2026 AMKAD.xlsm`"],
     ["2 · Load", "Power Query", "Picks the latest file, reads its **RawData** table"],
     ["3 · Calculate", "Office Script", "Does the maths, builds tables + a JSON result"],
     ["4 · Assemble", "Power Automate", "Drops the JSON into the HTML template"],
     ["5 · Deliver", "HTML report", "Emailed, saved, or shared"]],
    widths=[1.1, 1.5, 3.8])
callout("The golden rule.",
        "The report can only show what is in **RawData**. When a number looks wrong, work **backwards** "
        "along that chain — report → RawData → the source Excel file — until you find where the number "
        "first goes wrong.", DHL_RED)

# Section 2
eyebrow("Section 02"); doc.add_heading('The one concept that matters: "latest snapshot per month"', level=1)
body("The daily file's `RawData` table doesn't hold just today's numbers — it holds a **running history**: "
     "many daily captures for every month, going back to 2024. If you simply added all of June's rows "
     "together you'd count June twenty-plus times over, and every figure would be roughly **20× too big**.")
body("So for each month the script keeps **only the latest date that is a real, full extract with actual "
     "money in it**, and ignores every earlier day. That produces a true point-in-time (Month-to-Date) view. "
     "You can always see exactly which date it chose in **Data Quality → Snapshot Audit**.")

# Section 3
eyebrow("Section 03"); doc.add_heading("The pieces & where they live", level=1)
body("Deploying just means copying a file's contents into the right tool and saving. There is no build step.")
styled_table(["File in the repo", "What it is", "Where it runs"],
    [["`power_query/…`", "Picks the latest daily file, loads RawData", "Excel → Power Query"],
     ["`office_scripts/amkad_ar_report.ts`", "All calculations + builds the tables & JSON", "Excel Online → Automate"],
     ["`office_scripts/refresh_trigger.ts`", "No-op script that forces the RawData refresh (Section 5E)", "Excel Online → Automate"],
     ["`power_automate/html_template.html`", "Dashboard layout, charts, interactivity", "Power Automate → HTML action"],
     ["`power_automate/email_body.html`", "The automated email body (Section 5D)", "Power Automate → Send an email"],
     ["`docs/MAINTENANCE_GUIDE.md`", "The Markdown version of this guide", "—"]],
    widths=[2.5, 2.4, 1.8])

# Section 4
eyebrow("Section 04"); doc.add_heading("Routine maintenance calendar", level=1)
styled_table(["When", "Do this"],
    [["Every review", "Open the **Data Quality tab**. Confirm 'Latest snapshot date used' is what you expect and 'Rows used' is in the thousands, not tens."],
     ["Start of each year", "The Power Query path is hard-coded to `2026 AMKAD Summaries/Daily`. Change 2026 to the new year. This is the #1 thing that will suddenly 'break' the report."],
     ["If file names change", "Update the filename date-parsing in Power Query."],
     ["If a RawData column is renamed", "Update the field mapping in the Office Script."],
     ["If the layout needs a change", "Edit the HTML template and re-paste into Power Automate."],
     ["Occasionally", "Confirm the 'AMKAD – Trigger Daily Refresh' flow (5E) is still enabled and still scheduled before the main report flow, with a healthy gap between them."]],
    widths=[1.7, 5.1])

# Section 5
eyebrow("Section 05"); doc.add_heading("How to update each component", level=1)

doc.add_heading("A · The dashboard look & behaviour", level=3)
body("File: `power_automate/html_template.html`", SOFT, 9.5, after=3)
for step in ["Edit the file in the repo and copy the entire contents.",
             "In Power Automate, open the HTML/Compose action, paste over it, Save.",
             "Run the flow once and open the output to check."]:
    doc.add_paragraph(step, style="List Number")
callout("Two hard rules.",
        "Never touch the `@{outputs('Run_script')…}` placeholders — they're where real values get injected. "
        "And never add anything loaded from the internet (external fonts, chart libraries, images): "
        "SharePoint/Outlook preview blocks it, so the page must stay fully self-contained.",
        RGBColor(0xFF,0xCC,0x00))
callout("The Simulator tab is calibrated, not guessed.",
        "The conservative/moderate/aggressive aging roll-forward rates are computed client-side from the "
        "report's own historical trend data (see `calibrateAgingRates()`), not hardcoded round numbers. The "
        "Simulator tab displays the sample size and computed rates directly, so it stays auditable.",
        DHL_RED)

doc.add_heading("B · The calculations & tables", level=3)
body("File: `office_scripts/amkad_ar_report.ts`", SOFT, 9.5, after=3)
for step in ["Edit the file and copy the entire contents.",
             "Excel Online → Automate → open the AMKAD script → Code Editor → paste → Save → Run once.",
             "The flow uses this same script, so it picks up the change on its next run."]:
    doc.add_paragraph(step, style="List Number")
body("**Field mapping:** the script matches columns by name (e.g. `Total AR € (Live)`). If a column is "
     "renamed in the source, that metric reads blank → **€0 for every month**. Fix by updating the "
     "expected name in the script's mapping section.", SOFT, 9.8)

doc.add_heading("C · Which daily file is read", level=3)
body("File: `power_query/…`", SOFT, 9.5, after=3)
for step in ["Year rollover: change the folder year in the path.",
             "Naming change: files are 'Daily_Performance_Report- June 30 2026 AMKAD.xlsm' — the query reads the Month Day Year from the name. Update the parse if the format changes.",
             "To see which file it picked: temporarily set the query's last line to `Ranked` and read the top row's Name & FileDate. Revert after."]:
    doc.add_paragraph(step, style="List Number")

doc.add_heading("D · The automated email", level=3)
body("File: `power_automate/email_body.html`", SOFT, 9.5, after=3)
body("The email is a short, email-safe KPI summary with a button linking to the full report — not the "
     "full interactive dashboard, since Outlook strips the scripts/CSS that make it work.", SOFT, 9.8)
for step in ["Edit the file and copy the entire contents.",
             "In Power Automate, open Send an email (V2) → Body → code view (</>) → select all → paste → Save."]:
    doc.add_paragraph(step, style="List Number")
body("The report link is an **expression**, not the dynamic-content picker: "
     "`@{outputs('Create_sharing_link_for_a_file_or_folder')?['body/link/webUrl']}`. If that action is ever "
     "renamed, update the expression to match. Power Automate's code-view editor strips some HTML on save — "
     "always re-paste the whole file rather than patching a fragment.", SOFT, 9.8)

doc.add_heading("E · Automatic daily refresh", level=3)
body("File: `office_scripts/refresh_trigger.ts`", SOFT, 9.5, after=3)
body("Office Scripts cannot trigger a Power Query refresh from inside a script — there is no supported API "
     "for it. The fix is two separate scheduled Power Automate flows, not one:", SOFT, 9.8)
styled_table(["Flow", "Trigger", "What it does"],
    [["AMKAD – Trigger Daily Refresh (new)", "Recurrence, runs first", "Runs refresh_trigger.ts — a no-op script whose only job is to open the workbook, which (with 'Refresh data when opening the file' enabled) triggers the refresh."],
     ["AMKAD AR Report (existing)", "Recurrence, runs after, with a gap", "Runs the real report script, builds the HTML, sends the email."]],
    widths=[2.2, 1.6, 3.9])
body("The time gap between the two flows (start with 20–30 minutes) is what guarantees the refresh finishes "
     "before the report reads the data — the two flows don't share a session, so there's no race condition. "
     "To verify it's working: after a scheduled run, check the Data Quality tab's 'Latest snapshot date used'.", SOFT, 9.8)

# Section 6
eyebrow("Section 06"); doc.add_heading("Troubleshooting playbook", level=1)
body("Always start at the report's **Data Quality tab**, then match your symptom.", SOFT)
tcard("Critical", CRIT, "FBEAEA", "Every number is ~20× too big",
      "The latest-snapshot filter isn't applied, so multiple daily extracts get summed.",
      "**Redeploy the current Office Script** — the snapshot logic lives there.",
      check="'Rows used' is huge; Snapshot Audit shows many dates combined.")
tcard("Data gap", WARN, "FBF1DF", "One month is blank / all 0.0% (the 'April' case)",
      "That month has **no money in the source** — the rows exist but the live EUR column is blank/zero.",
      "It's a **source-data** problem, not the report. Fix the daily file upstream. The current script already hides a fully-empty month automatically.",
      check="Snapshot Audit → that month shows €0 on every date.")
tcard("Mapping", WARN, "FBF1DF", "One metric (e.g. UAC) is €0 for every month",
      "A RawData column was renamed, so the script can't find it.",
      "Update the field mapping in the Office Script (Section 05).")
tcard("Refresh", CRIT, "FBEAEA", "The whole report failed / Power Query errors",
      "Most often **year rollover** — the path still has the old year. Also: naming change, no files matched, or SharePoint sign-in expired.",
      "Update the path/parse; confirm you can open the folder; re-authenticate SharePoint.")
tcard("Display", INFO, "EAF0F7", "Charts blank + red 'issue running scripts' banner",
      "A sandboxed preview (SharePoint/Outlook) restricting scripts, or an external resource was added to the HTML.",
      "Open it in a real browser tab to confirm it works. Make sure no https:// script/font/image was added to the template.")
tcard("Selection", INFO, "EAF0F7", "'Latest month' is wrong / an old month shows as newest",
      "Power Query picked the wrong daily file (often a copied/re-uploaded one).",
      "Use the 'which file is it using?' check (Section 05) and confirm the newest file is named with the correct date.")
tcard("Stale data", WARN, "FBF1DF", "Data still looks a day (or more) stale despite the refresh setup",
      "The 'AMKAD – Trigger Daily Refresh' flow (5E) didn't run, was disabled, or the gap before the report flow was too short.",
      "Re-enable/reschedule the trigger flow; widen the gap between the two flows if refresh is taking longer than expected.",
      check="Power Automate run history for the trigger flow — did it run, and when, relative to the report flow?")
tcard("Email link", INFO, "EAF0F7", "The 'Open the report' button/link in the email doesn't show or work",
      "Power Automate's code-view editor stripped some HTML on save, or the sharing-link action produced no URL / was renamed.",
      "Re-paste the entire email_body.html file (5D). Check run history → 'Create sharing link…' → Outputs to confirm link.webUrl has a real value.")

# Section 7
eyebrow("Section 07"); doc.add_heading("Before you change anything", level=1)
for item in ["Commit to Git before and after your change — the history is your undo button.",
             "Change one component at a time, then re-run and check the report.",
             "After any change, open the Data Quality tab and sanity-check the snapshot date and rows used.",
             "Keep every @{outputs('Run_script')…} placeholder intact in the HTML.",
             "Never add internet-loaded resources to the HTML."]:
    p = doc.add_paragraph(style="List Bullet")
    r = p.add_run("☐  " + item); r.font.size = Pt(10.5)

# Section 8
eyebrow("Section 08"); doc.add_heading("Hand-off & escalation", level=1)
body("Fill these in for whoever comes after you.", SOFT)
styled_table(["Role / question", "Name & contact"],
    [["Report owner / business contact", " "],
     ["SharePoint site owner (for access issues)", " "],
     ["Who produces the daily .xlsm (for source-data gaps like empty months)", " "],
     ["Power Automate environment + flow name", " "]],
    widths=[4.3, 2.5])

# footer note
doc.add_paragraph()
f = doc.add_paragraph()
fr = f.add_run("When in doubt: the report shows what's in RawData · RawData comes from the latest daily file · "
               "the Data Quality tab tells you what the script actually used. Follow that chain and you'll "
               "find almost any problem.")
fr.italic = True; fr.font.size = Pt(9); fr.font.color.rgb = MUTED

import sys
out = sys.argv[1] if len(sys.argv) > 1 else "AMKAD_AR_Maintenance_Guide.docx"
doc.save(out)
print("saved", out)
