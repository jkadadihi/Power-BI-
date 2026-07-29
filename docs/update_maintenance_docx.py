#!/usr/bin/env python3
"""
Update the AMKAD HTML Maintenance Guide (.docx) IN PLACE.

The Word file carries embedded screenshots that the Markdown source and
build_docx.py know nothing about, so regenerating from Markdown would silently
drop them. This edits the existing document instead: it clones the styling of
elements already in the file (table rows, callout boxes, section labels) so new
content matches, and leaves every image untouched.

Adds coverage for work done after the guide was written:
  - Stage 1, the daily file build (Cust Sol import) that creates the daily file
  - The SPR tab: per-country commentary sheets and the MonthEndAR month-end pull
  - Output files now named from the source file's date, not today's
  - Troubleshooting entries for a missed day, a blank SPR tab, and a missing
    month-end table

Usage: python3 docs/update_maintenance_docx.py <in.docx> <out.docx>
"""
import copy
import sys

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph


# ---------- helpers that clone existing formatting ----------

def set_cell_text(cell, text, bold=None):
    """Replace a cell's text while keeping its first run's formatting."""
    paras = cell.paragraphs
    first = paras[0]
    for extra in paras[1:]:
        extra._element.getparent().remove(extra._element)
    if first.runs:
        keep = first.runs[0]
        for r in first.runs[1:]:
            r._element.getparent().remove(r._element)
        keep.text = text
        if bold is not None:
            keep.font.bold = bold
    else:
        run = first.add_run(text)
        if bold is not None:
            run.font.bold = bold


def clone_row(table, source_index, values):
    """Append a row copied from an existing one, so borders/shading match."""
    src = table.rows[source_index]._tr
    new = copy.deepcopy(src)
    table._tbl.append(new)
    row = table.rows[-1]
    for cell, val in zip(row.cells, values):
        set_cell_text(cell, val)
    return row


def clone_callout(doc, template_table, text, before):
    """Copy a 1x1 troubleshooting callout box and drop in new text."""
    new = copy.deepcopy(template_table._tbl)
    before.addprevious(new)
    tbl = Table(new, doc)
    set_cell_text(tbl.rows[0].cells[0], text)
    return tbl


def clone_para(doc, template_para, text, before):
    """Copy a paragraph's styling and insert it before `before`."""
    new = copy.deepcopy(template_para._p)
    before.addprevious(new)
    p = Paragraph(new, doc)
    if p.runs:
        keep = p.runs[0]
        for r in p.runs[1:]:
            r._element.getparent().remove(r._element)
        keep.text = text
    else:
        p.add_run(text)
    return p


def find_para(doc, predicate):
    for p in doc.paragraphs:
        if predicate(p.text.strip()):
            return p
    return None


def main(src_path, out_path):
    doc = Document(src_path)

    # Style templates lifted from the document itself.
    tmpl_section = find_para(doc, lambda t: t == "SECTION 05")
    tmpl_h3 = next(p for p in doc.paragraphs if p.style.name == "Heading 3")
    tmpl_body = find_para(doc, lambda t: t.startswith("File: power_automate/html_template"))
    tmpl_list = next(p for p in doc.paragraphs if p.style.name == "List Number")
    tmpl_callout = doc.tables[7]          # a troubleshooting box
    tmpl_h1 = next(p for p in doc.paragraphs if p.style.name == "Heading 1")

    # ---------- 1 · flow table: show that the daily file is now built too ----------
    t1 = doc.tables[0]
    set_cell_text(t1.rows[1].cells[0], "1 · Build")
    set_cell_text(t1.rows[1].cells[1], "Cust Sol import flow")
    set_cell_text(t1.rows[1].cells[2],
                  "Reads the raw source file and builds that day's "
                  "Daily_Performance_Report .xlsm, carrying history forward (Section 5F)")
    for vals in [
        ("2 · Load", "Power Query", "Picks the latest daily file, reads its RawData table"),
        ("3 · Calculate", "Office Script",
         "Does the maths, builds tables + a JSON result. Also reads the SPR sheets and MonthEndAR (5G)"),
        ("4 · Assemble", "Power Automate", "Drops the JSON into the HTML template"),
        ("5 · Deliver", "HTML report", "Emailed, saved, or shared"),
    ]:
        pass  # rows 2..5 already carry these; only their wording is refreshed below
    set_cell_text(t1.rows[3].cells[2],
                  "Does the maths, builds tables + a JSON result. Also reads the "
                  "SPR sheets and MonthEndAR (Section 5G)")

    # ---------- 2 · files table ----------
    t3 = doc.tables[2]
    set_cell_text(t3.rows[1].cells[0], "power_query/custsol_rawdata.m")
    for vals in [
        ("office_scripts/custsol_read_daily.ts",
         "Reads the raw source file, returns its rows as JSON (5F)",
         "Excel Online → Automate"),
        ("office_scripts/custsol_append_rawdata_eur.ts",
         "Appends those rows to the new daily file's RawData (5F)",
         "Excel Online → Automate"),
        ("power_query/spr_month_end_ar.m",
         "Pulls the latest closed month from Debits into MonthEndAR (5G)",
         "Excel → Power Query"),
        ("SPR_<CC> sheets in the input workbook",
         "Collector commentary: SPRIssues_<CC> and SPRForecast_<CC> (5G)",
         "Excel (typed by collectors)"),
    ]:
        clone_row(t3, 1, vals)

    # ---------- 3 · maintenance calendar ----------
    t4 = doc.tables[3]
    for vals in [
        ("Each month, before the review",
         "Collectors fill the SPR sheets for their country. The SPR tab's country strip "
         "shows a coverage count (e.g. '7 of 9') so gaps are visible beforehand."),
        ("After each month close",
         "Refresh the MonthEndAR query so the SPR tab shows the new closed month. "
         "It holds one month only and is replaced on each refresh."),
        ("If a daily run fails",
         "Re-trigger the flow on that day's source file. Output files are named from "
         "the SOURCE file's date, so backfilling produces the correct name (5F)."),
    ]:
        clone_row(t4, 1, vals)

    # ---------- 4 · new sections 5F and 5G, before SECTION 06 ----------
    anchor = find_para(doc, lambda t: t == "SECTION 06")._p

    def h3(text):
        clone_para(doc, tmpl_h3, text, anchor)

    def body(text):
        clone_para(doc, tmpl_body, text, anchor)

    def item(text):
        clone_para(doc, tmpl_list, text, anchor)

    def callout(text):
        clone_callout(doc, tmpl_callout, text, anchor)

    h3("F · The daily file build (Stage 1)")
    body("Where it lives: Excel Online → Automate → the 'custsol read daily' and "
         "'custsol append rawdata eur' scripts.")
    body("This flow creates the Daily_Performance_Report .xlsm that everything else reads. "
         "It replaced a manual routine of copying the emailed report into a template, "
         "calculating, copying two column ranges out, refreshing and saving a new file.")
    item("Trigger: a file is created in …/Daily/Cust_Sol Daily File.")
    item("Two Initialize variable actions build the output file name and the month folder.")
    item("Get files → Filter array → Compose picks the newest existing daily file.")
    item("Get file content + Create file copy it forward under the new name. This is what "
         "carries the full history into the new file.")
    item("Run script 'custsol read daily' on the SOURCE file returns its rows as JSON.")
    item("Run script 'custsol append rawdata eur' on the NEW file appends those rows to RawData.")
    body("Why two scripts: an Office Script can only touch the workbook it runs on, so it "
         "cannot read the source and write the daily file in one pass. Power Automate carries "
         "the JSON between them.")
    callout("Naming.  The output file is named from the SOURCE file's date (its MMDDYYYY "
            "prefix), not today's date. That is what makes backfilling work: re-trigger the "
            "flow on an old source file and the output gets that day's correct name. It also "
            "means a misnamed source file fails the run rather than silently producing a "
            "wrongly dated file — keep flow-failure notifications on.")
    body("Built-in safety: the append skips rows already present (matched on Date + Country + "
         "Customer), so re-running the same day cannot double-count — check skippedDuplicates "
         "in the last action's output. Rows with blank EUR columns (normal for a day or two "
         "after close) are appended anyway and counted in missingEuroRows. Columns are matched "
         "by header name, so reordering columns in the source is harmless; renaming one is not.")

    h3("G · The SPR tab (commentary + month-end figures)")
    body("Files: the SPR_<CC> sheets in the input workbook + power_query/spr_month_end_ar.m")
    body("The SPR tab shows an Account Business Review per customer: Main Issues, Actions, "
         "cash forecasting and the AR figures. It has two inputs.")
    body("1 · Commentary. Each country has a sheet (SPR_US, SPR_MX, …) with two tables: "
         "SPRIssues_<CC> (Month, Country, Customer, Issue, Action — one row per issue) and "
         "SPRForecast_<CC> (Month, Country, Customer + the five cash figures — one row per "
         "customer). Month is YYYY-MM and history is kept by month: to start a new cycle, copy "
         "the rows down and change the Month.")
    callout("Adding customers and countries.  The report finds these tables by NAME PREFIX, so "
            "a new country needs no code change — create a sheet with tables named "
            "SPRIssues_XX and SPRForecast_XX and it appears. A new customer needs nothing at "
            "all: once it shows in the AR data it is listed, tagged 'No commentary' until "
            "someone writes some. Names are matched ignoring case and extra spaces, so 'nokia' "
            "and 'Nokia ' still attach; genuinely different names (Nokia vs Nokia Corp) stay "
            "separate accounts by design.")
    body("2 · Month-end figures. SPRs are reported on month-end numbers, but the daily pipeline "
         "only holds daily snapshots, and month close can shift by days — so the last daily "
         "snapshot of a month is NOT the close. spr_month_end_ar.m pulls the latest closed "
         "month from the Debits source into a MonthEndAR table in the input workbook.")
    item("Requires a text parameter SharePointSite_Url.")
    item("The query name AND the loaded table name must both be exactly MonthEndAR.")
    item("It holds one month — the latest close — and each refresh replaces it.")
    body("What the tab shows: each account lists Financial Performance twice — month-end close "
         "above, daily/current below — each labelled with the table it was read from. The "
         "country strip totals ONE declared basis, switchable between month-end and current, "
         "and never blends the two.")

    # ---------- 5 · troubleshooting entries, before SECTION 07 ----------
    anchor7 = find_para(doc, lambda t: t == "SECTION 07")._p
    for text in [
        "MISSED DAY   A daily run failed, so there is no file for that date.  "
        "Most gaps are harmless: the report picks the latest valid snapshot per month, so a "
        "mid-month hole does not change any month-end figure. The one that matters is a gap on "
        "the last business day(s) of a month, because the month then resolves to an earlier "
        "date and its 'close' is quietly wrong. Fix: the source file is still in Cust_Sol "
        "Daily File — re-trigger the flow on it and the output is named for that day.",

        "SPR BLANK   The SPR tab shows no accounts, or commentary does not appear.  "
        "Check, in order: the country has a sheet with tables named SPRIssues_<CC> / "
        "SPRForecast_<CC> (prefix must match exactly); the Month cell is YYYY-MM; and the "
        "Customer spelling matches the AR data. Case and extra spaces are tolerated, different "
        "names are not. The coverage count on the country strip ('7 of 9') is the quickest way "
        "to spot a customer whose commentary did not attach.",

        "NO MONTH-END TABLE   An account shows only the daily figures.  "
        "MonthEndAR has no row for that customer. Confirm the MonthEndAR query refreshed after "
        "the close and that the table is named exactly MonthEndAR. Note the tab shows the most "
        "recent close available even when the commentary is dated a later, still-open month — "
        "it names both periods when they differ, which is expected mid-month.",
    ]:
        clone_callout(doc, tmpl_callout, text, anchor7)

    # ---------- 6 · remove every repo reference ----------
    # The people who inherit this will not have the Git repo — it was only ever
    # our working copy, on a personal account. Every component is therefore
    # named by WHERE THEY OPEN IT, not by a file path they cannot reach.
    renames = {
        "FILE IN THE REPO": "COMPONENT",
        "power_query/custsol_rawdata.m": "Power Query query: RawData",
        "office_scripts/amkad_ar_report.ts": "Office Script: AMKAD AR report",
        "office_scripts/refresh_trigger.ts": "Office Script: refresh trigger",
        "power_automate/html_template.html": "HTML template (report flow)",
        "power_automate/email_body.html": "Email body (Send an email V2)",
        "office_scripts/custsol_read_daily.ts": "Office Script: custsol read daily",
        "office_scripts/custsol_append_rawdata_eur.ts": "Office Script: custsol append rawdata eur",
        "office_scripts/custsol_append_rawdata_eu": "Office Script: custsol append rawdata eur",
        "power_query/spr_month_end_ar.m": "Power Query query: MonthEndAR",
        "The Markdown version of this guide": "This guide",
        "docs/MAINTENANCE_GUIDE.md": "This document",
        "File: power_automate/html_template.html":
            "Where it lives: the HTML / Compose action inside the report flow.",
        "File: office_scripts/amkad_ar_report.ts":
            "Where it lives: Excel Online → Automate → the AMKAD AR report script.",
        "File: power_query/…":
            "Where it lives: the input workbook → Data → Queries.",
        "File: power_automate/email_body.html":
            "Where it lives: Power Automate → Send an email (V2) → Body.",
        "File: office_scripts/refresh_trigger.ts":
            "Where it lives: Excel Online → Automate → the refresh-trigger script.",
        "Files: office_scripts/custsol_read_daily.ts + custsol_append_rawdata_eur.ts":
            "Where it lives: Excel Online → Automate → the 'custsol read daily' and "
            "'custsol append rawdata eur' scripts.",
        "Files: the SPR_<CC> sheets in the input workbook + power_query/spr_month_end_ar.m":
            "Where it lives: the SPR_<CC> sheets in the input workbook, plus the "
            "MonthEndAR query in that same workbook.",
        "Runs refresh_trigger.ts — a no-op script":
            "Runs the refresh-trigger script — a no-op",
        "Update the filename date-parsing in Power Query":
            "Update the filename date-parsing in the RawData query",
    }

    def rewrite(text):
        for old, new in renames.items():
            if old in text:
                text = text.replace(old, new)
        return text

    def rewrite_para(p):
        original = p.text
        updated = rewrite(original)
        if updated == original or not p.runs:
            return
        keep = p.runs[0]
        for r in p.runs[1:]:
            r._element.getparent().remove(r._element)
        keep.text = updated

    for p in doc.paragraphs:
        rewrite_para(p)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    rewrite_para(p)

    # Where the master copies actually live has to be stated, since the repo is
    # not available to whoever inherits this.
    after_files = doc.tables[2]._tbl
    nxt = after_files.getnext()
    clone_callout(
        doc, tmpl_callout,
        "Master copies.  Keep the current text of every script, query and template in the "
        "shared SharePoint folder next to the workbooks, and update it whenever you change "
        "one. That folder is the only source of truth for this automation — there is no "
        "build step and no other copy to fall back on.",
        nxt if nxt is not None else doc.paragraphs[-1]._p)

    doc.save(out_path)
    print("Saved", out_path)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
