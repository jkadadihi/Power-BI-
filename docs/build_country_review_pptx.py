#!/usr/bin/env python3
"""
Build the AMKAD Country Performance Review deck by cloning the REAL master
template (docs/templates/AMKAD_Country_Review_Template.pptx) and only
editing data-driven text/table cells — every shape, font, color, and logo
in the original is preserved exactly.

How it works:
  1. Open the master template as-is.
  2. Edit the title slide's country/month text.
  3. Use the Texas Instruments slide pair (the template's slides 6-7) as the
     pattern for a customer cover+detail slide. Duplicate that pair once per
     customer in the data, editing only:
       - customer name / date text
       - the TTL AR / Past Due / Over 90 summary line
       - the "FINANCIALS" table's Total AR / Current / Overdue / >90 cells
     Everything else on those two slides (headers, box positions, fonts,
     the logo picture, box colors) is left completely untouched.
  4. Delete the original 9 hardcoded customer pairs (their commentary and
     logos belong to last cycle's real customers, not this run).
  5. Reorder so Questions / Thank you stay at the very end.

Known limitation: customer LOGOS are static images baked into the template.
There's no way to auto-source a different company's logo from AR figures,
so every generated customer slide currently shows the template's original
logo image. Swap it manually per customer if that matters, or tell me and
I'll wire in a logo-lookup step if you have a folder of logo files.

Left blank on purpose (needs a person, not derivable from AR data):
  - AMKAD POC name, Net terms / Go-Live date
  - MAIN ISSUES / CURRENT ACTIONS / SUPPORT NEEDED (already blank in the
    template's own structure)
  - PAYMENTS / Cash Forecasting detail text and table values

Usage:
    python3 build_country_review_pptx.py <data.json> <output.pptx> [template.pptx]
"""
import copy
import json
import sys
from pptx import Presentation

TEMPLATE_PATH_DEFAULT = "docs/templates/AMKAD_Country_Review_Template.pptx"

TITLE_SLIDE_IDX = 0
COVER_TEMPLATE_IDX = 5   # "slide 6" — Texas Instruments cover
DETAIL_TEMPLATE_IDX = 6  # "slide 7" — Texas Instruments detail
FIRST_CUSTOMER_IDX = 5
LAST_CUSTOMER_IDX = 22   # inclusive — the 9 original customer pairs (slides 6-23)


def euro(v):
    try:
        return f"€ {round(v):,}".replace(",", ",")
    except Exception:
        return "€ 0"


# ---------------- slide-level plumbing (python-pptx has no native support) ----------------

R_NS_URI = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _remap_rids(element, rid_map):
    """python-pptx's _add_relationship always mints a fresh rId — it can't
    preserve the original one. So the copied shapes' r:embed/r:id/r:link
    attributes (e.g. a picture's <a:blip r:embed="rId5">) must be rewritten
    to point at the NEW id, or the embedded image reference breaks."""
    for el in element.iter():
        for attr_name in list(el.attrib.keys()):
            if attr_name.startswith("{" + R_NS_URI + "}"):
                old_rid = el.attrib[attr_name]
                if old_rid in rid_map:
                    el.attrib[attr_name] = rid_map[old_rid]


def duplicate_slide(prs, index):
    """Append a copy of slide `index` to the end of the deck, including its
    images/relationships, and return the new slide."""
    source = prs.slides[index]
    dest = prs.slides.add_slide(source.slide_layout)

    for shp in list(dest.shapes):
        shp._element.getparent().remove(shp._element)

    rid_map = {}
    for rel_id, rel in source.part.rels.items():
        if "slideLayout" in rel.reltype or "notesSlide" in rel.reltype:
            continue
        if rel.is_external:
            new_rid = dest.part.rels._add_relationship(rel.reltype, rel.target_ref, is_external=True)
        else:
            new_rid = dest.part.rels._add_relationship(rel.reltype, rel.target_part, is_external=False)
        rid_map[rel_id] = new_rid

    for shp in source.shapes:
        el = copy.deepcopy(shp._element)
        _remap_rids(el, rid_map)
        dest.shapes._spTree.append(el)

    return dest


def delete_slide(prs, index):
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    rId = slides[index].rId
    prs.part.drop_rel(rId)
    xml_slides.remove(slides[index])


def move_slide(prs, old_index, new_index):
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    xml_slides.remove(slides[old_index])
    xml_slides.insert(new_index, slides[old_index])


# ---------------- text/table editing (structure-preserving) ----------------

def set_run(shape, para_idx, run_idx, text):
    try:
        shape.text_frame.paragraphs[para_idx].runs[run_idx].text = text
    except IndexError:
        pass


def clear_all_runs(shape):
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            run.text = ""


def shape_by_name(slide, name):
    for shp in slide.shapes:
        if shp.name == name:
            return shp
    return None


def fill_cover_slide(slide, customer):
    name_box = shape_by_name(slide, "Text Placeholder 1")
    if name_box is not None:
        set_run(name_box, 0, 0, f"{customer} — Net terms: TBD          Go Live: TBD")

    poc_box = shape_by_name(slide, "Rounded Rectangle 12")
    if poc_box is not None:
        set_run(poc_box, 1, 0, "— fill in —")


def fill_detail_slide(slide, customer, country_code, agg):
    title_box = shape_by_name(slide, "TextBox 1")
    if title_box is not None:
        set_run(title_box, 0, 0, f"{customer}  {country_code} – ")

    summary_box = shape_by_name(slide, "TextBox 12")
    if summary_box is not None:
        set_run(summary_box, 3, 0, f"{customer} {country_code}- TTL AR: {euro(agg['totalAR'])} ")
        set_run(summary_box, 4, 0, f"       Past Due: {euro(agg['overdue'])}")
        set_run(summary_box, 5, 0, f"       Over 90: {euro(agg['gt90'])}")

    # PAYMENTS free-text detail — genuinely manual/forecast content, blank it.
    payments_box = shape_by_name(slide, "TextBox 10")
    if payments_box is not None:
        clear_all_runs(payments_box)

    # FINANCIALS table — fill what we can compute precisely; leave the finer
    # 0-30/31-60/61-90 aging split as "—" since our data only has a 0-60
    # aggregate, not that granularity, and the ask was to look exactly like
    # the template rather than relabel its columns.
    for shp in slide.shapes:
        if shp.has_table and shp.name == "Table 3":
            t = shp.table
            current = agg["totalAR"] - agg["overdue"]
            values = ["Total", euro(agg["totalAR"]), euro(current), euro(agg["overdue"]), "—", "—", "—", euro(agg["gt90"])]
            for ci, val in enumerate(values):
                if ci < len(t.columns):
                    t.cell(1, ci).text = val
        elif shp.has_table and shp.name == "Table 14":
            t = shp.table
            for ri in range(1, len(t.rows)):
                t.cell(ri, 1).text = "— fill in —"


def fill_title_slide(prs, country_name, report_month):
    slide = prs.slides[TITLE_SLIDE_IDX]
    box = shape_by_name(slide, "Text Placeholder 4")
    if box is not None:
        set_run(box, 0, 0, country_name)
        set_run(box, 1, 0, report_month)


def build(data, out_path, template_path):
    prs = Presentation(template_path)

    country_name = data.get("CountryName", "Country")
    country_code = data.get("CountryCode", "")
    report_month = data.get("ReportMonth", "")
    countries_data = data.get("CountryDrilldownJson", [])
    customers = data.get("CustomerDrilldownJson", [])

    fill_title_slide(prs, country_name, report_month)

    by_customer = {}
    for row in customers:
        cust = row.get("customer", "Unknown")
        by_customer.setdefault(cust, []).append(row)

    new_slide_indices = []
    for cust, rows in by_customer.items():
        agg = {
            "totalAR": sum(r.get("totalAR", 0) for r in rows),
            "overdue": sum(r.get("overdue", 0) for r in rows),
            "gt60": sum(r.get("gt60", 0) for r in rows),
            "gt90": sum(r.get("gt90", 0) for r in rows),
        }
        if agg["totalAR"] <= 0:
            continue

        cover = duplicate_slide(prs, COVER_TEMPLATE_IDX)
        fill_cover_slide(cover, cust)
        new_slide_indices.append(len(prs.slides.__iter__.__self__._sldIdLst) - 1)

        detail = duplicate_slide(prs, DETAIL_TEMPLATE_IDX)
        fill_detail_slide(detail, cust, country_code, agg)
        new_slide_indices.append(len(prs.slides.__iter__.__self__._sldIdLst) - 1)

    # Remove the original 9 hardcoded customer pairs (last cycle's real data).
    for idx in range(LAST_CUSTOMER_IDX, FIRST_CUSTOMER_IDX - 1, -1):
        delete_slide(prs, idx)

    # Keep "Questions" / "Thank you" at the very end, after the new customers.
    total = len(prs.slides.__iter__.__self__._sldIdLst)
    thank_you_idx = FIRST_CUSTOMER_IDX + 1  # right after Questions, pre-shift
    questions_idx = FIRST_CUSTOMER_IDX
    move_slide(prs, thank_you_idx, total - 1)
    move_slide(prs, questions_idx, total - 2)

    prs.save(out_path)
    print("Saved", out_path, "-", len(prs.slides.__iter__.__self__._sldIdLst), "slides")


if __name__ == "__main__":
    data_path = sys.argv[1] if len(sys.argv) > 1 else "sample_data.json"
    out_path = sys.argv[2] if len(sys.argv) > 2 else "AMKAD_Country_Review.pptx"
    template_path = sys.argv[3] if len(sys.argv) > 3 else TEMPLATE_PATH_DEFAULT
    with open(data_path) as f:
        data = json.load(f)
    build(data, out_path, template_path)
