#!/usr/bin/env python3
"""
AMKAD Internship Close-Out — comprehensive deck.

Focus: the AR reporting HTML automation and the daily data pipeline that now
feeds it, plus the SOA automation, the Power BI / database groundwork, and the
real engineering challenges overcome. Leads with a transparent, conservative
hours-saved model so the impact is defensible in front of a CFO.

Numbers are built from the intern's own figures:
  - Daily Cust_Sol file:  15 min/day x 5 days -> 1.25 hr/wk -> ~63 hr/yr
  - SOA automation:       6 weekly SOAs (Arrow x5 countries, Baker x1),
                          30-60 min each -> 3-6 hr/wk -> ~150-300 hr/yr
  - Combined:             ~4-7 hr/wk -> ~18-31 hr/mo -> ~210-360 hr/yr
                          (~5-9 full 40-hour work-weeks per year)
Conservative anchor used as the headline: ~210 hr/yr.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ---------- palette ----------
RED = RGBColor(0xD4, 0x05, 0x11)
RED_DARK = RGBColor(0xA5, 0x04, 0x0D)
YELLOW = RGBColor(0xFF, 0xCC, 0x00)
INK = RGBColor(0x1A, 0x1D, 0x21)
SOFT = RGBColor(0x5B, 0x64, 0x70)
MUTED = RGBColor(0x8A, 0x83, 0x7B)
BG = RGBColor(0xF7, 0xF6, 0xF3)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LINE = RGBColor(0xE2, 0xE0, 0xDA)
GOOD = RGBColor(0x0F, 0x8A, 0x2E)
INKCARD = RGBColor(0x23, 0x27, 0x2C)
FONT = "Calibri"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
SW, SH = prs.slide_width, prs.slide_height


def add_slide(bg=BG):
    s = prs.slides.add_slide(BLANK)
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    r.fill.solid(); r.fill.fore_color.rgb = bg
    r.line.fill.background(); r.shadow.inherit = False
    s.shapes._spTree.remove(r._element)
    s.shapes._spTree.insert(2, r._element)
    return s


def rect(s, l, t, w, h, color, line=False):
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    r.fill.solid(); r.fill.fore_color.rgb = color
    if line:
        r.line.color.rgb = LINE; r.line.width = Pt(0.75)
    else:
        r.line.fill.background()
    r.shadow.inherit = False
    return r


def rrect(s, l, t, w, h, color, line=False, radius=0.06):
    r = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    try:
        r.adjustments[0] = radius
    except Exception:
        pass
    r.fill.solid(); r.fill.fore_color.rgb = color
    if line:
        r.line.color.rgb = LINE; r.line.width = Pt(0.75)
    else:
        r.line.fill.background()
    r.shadow.inherit = False
    return r


def textbox(s, l, t, w, h, valign=MSO_ANCHOR.TOP):
    tb = s.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = valign
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    return tb, tf


def para(tf, text, size=14, color=INK, bold=False, italic=False, align=PP_ALIGN.LEFT,
         space_after=6, first=False, bullet=False, font=FONT):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.space_after = Pt(space_after)
    if bullet:
        text = "•  " + text
    r = p.add_run(); r.text = text
    r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
    r.font.color.rgb = color; r.font.name = font
    return p


def eyebrow(s, text, top=Inches(0.5), left=Inches(0.7)):
    tb, tf = textbox(s, left, top, Inches(11), Inches(0.4))
    para(tf, text.upper(), size=12.5, color=RED, bold=True, first=True)
    rect(s, left, top + Inches(0.35), Inches(0.4), Pt(3.2), YELLOW)


def slide_title(s, kicker, title, subtitle=None):
    eyebrow(s, kicker)
    tb, tf = textbox(s, Inches(0.7), Inches(0.9), Inches(11.9), Inches(1.0))
    para(tf, title, size=29, color=INK, bold=True, first=True)
    if subtitle:
        tb2, tf2 = textbox(s, Inches(0.7), Inches(1.62), Inches(11.9), Inches(0.6))
        para(tf2, subtitle, size=14.5, color=SOFT, first=True)
    rect(s, Inches(0.7), Inches(2.12), Inches(11.9), Pt(1.4), LINE)


def footer(s, n):
    tb, tf = textbox(s, Inches(0.7), Inches(7.15), Inches(9), Inches(0.3))
    para(tf, "AMKAD  ·  AR Reporting & Collections Automation  ·  Internship Close-Out",
         size=9, color=MUTED, first=True)
    tb2, tf2 = textbox(s, Inches(12.3), Inches(7.15), Inches(0.7), Inches(0.3))
    para(tf2, str(n), size=9, color=MUTED, first=True, align=PP_ALIGN.RIGHT)


def logo_placeholder(s, l, t, w=Inches(1.9), h=Inches(0.75)):
    box = rrect(s, l, t, w, h, WHITE, line=True, radius=0.08)
    box.line.color.rgb = MUTED
    ln = box.line._get_or_add_ln()
    from pptx.oxml.ns import qn
    d = ln.makeelement(qn('a:prstDash'), {'val': 'dash'}); ln.append(d)
    tb, tf = textbox(s, l, t, w, h, valign=MSO_ANCHOR.MIDDLE)
    para(tf, "paste DHL logo", size=9, color=MUTED, italic=True, first=True, align=PP_ALIGN.CENTER)


def bullets(s, items, left=Inches(0.7), top=Inches(2.45), width=Inches(11.9),
            size=15, gap=9, color=INK):
    tb, tf = textbox(s, left, top, width, Inches(4.4))
    first = True
    for item in items:
        para(tf, item, size=size, color=color, bullet=True, first=first, space_after=gap)
        first = False
    return tb


def stat_tile(s, l, t, w, h, value, label, color=RED, vsize=30):
    rrect(s, l, t, w, h, WHITE, line=True, radius=0.10)
    rect(s, l, t, Pt(4), h, color)
    tb, tf = textbox(s, l + Inches(0.24), t + Inches(0.14), w - Inches(0.42), h - Inches(0.28),
                     valign=MSO_ANCHOR.MIDDLE)
    para(tf, value, size=vsize, color=INK, bold=True, first=True, space_after=2)
    para(tf, label, size=11.5, color=SOFT, space_after=0)


def card(s, l, t, w, h, title, lines, accent=RED, title_size=14, body_size=11.5):
    rrect(s, l, t, w, h, WHITE, line=True, radius=0.07)
    rect(s, l, t, w, Pt(4), accent)
    tb, tf = textbox(s, l + Inches(0.24), t + Inches(0.22), w - Inches(0.44), h - Inches(0.4))
    para(tf, title, size=title_size, color=INK, bold=True, first=True, space_after=7)
    for ln in lines:
        para(tf, ln, size=body_size, color=SOFT, bullet=True, space_after=5)


# =====================================================================
# 1 — TITLE
# =====================================================================
s = add_slide(INK)
rect(s, 0, 0, SW, Inches(0.18), RED)
rect(s, 0, Inches(0.18), SW, Pt(3), YELLOW)
logo_placeholder(s, Inches(10.9), Inches(0.55))
tb, tf = textbox(s, Inches(0.8), Inches(2.1), Inches(11), Inches(0.5))
para(tf, "DHL EXPRESS AMERICAS  ·  KEY ACCOUNT DESK (AMKAD)", size=13,
     color=YELLOW, bold=True, first=True)
tb, tf = textbox(s, Inches(0.8), Inches(2.7), Inches(11.7), Inches(2.0))
para(tf, "Automating AR Reporting & Collections", size=44, color=WHITE, bold=True, first=True,
     space_after=6)
para(tf, "From manual, invisible processes to a live dashboard and self-running data pipelines",
     size=17, color=RGBColor(0xC9, 0xCE, 0xD4))
tb, tf = textbox(s, Inches(0.8), Inches(5.7), Inches(11), Inches(1.0))
para(tf, "Jean Kadadihi", size=17, color=WHITE, bold=True, first=True, space_after=2)
para(tf, "Finance Intern, Controlling  ·  Internship Close-Out Presentation", size=13,
     color=RGBColor(0xB4, 0xBA, 0xC1))


# =====================================================================
# 2 — IMPACT HEADLINE (hook)
# =====================================================================
s = add_slide()
eyebrow(s, "The Bottom Line")
tb, tf = textbox(s, Inches(0.7), Inches(0.9), Inches(11.9), Inches(1.0))
para(tf, "What this internship delivered", size=29, color=INK, bold=True, first=True)
rect(s, Inches(0.7), Inches(1.75), Inches(11.9), Pt(1.4), LINE)

stat_tile(s, Inches(0.7), Inches(2.15), Inches(3.75), Inches(1.7),
          "~210 hrs", "Collector time recovered per year (conservative) — up to ~360 hrs", vsize=32)
stat_tile(s, Inches(4.79), Inches(2.15), Inches(3.75), Inches(1.7),
          "6 / week", "Recurring SOAs generated & emailed automatically (Arrow ×5, Baker ×1)", vsize=32)
stat_tile(s, Inches(8.88), Inches(2.15), Inches(3.72), Inches(1.7),
          "1 new", "Live AR reporting capability that did not exist before", vsize=32)

tb, tf = textbox(s, Inches(0.7), Inches(4.2), Inches(11.9), Inches(1.0))
para(tf, "≈ 5 to 9 full 40-hour work-weeks of collector time returned to actual collections — every year.",
     size=18, color=RED_DARK, bold=True, first=True)
bullets(s, [
    "Built net-new AR visibility (aging, UAC, trends, what-if simulators) that leadership never had before.",
    "Automated the daily data file that feeds it — no one has to rebuild it by hand anymore.",
    "Removed a single-person dependency and the error risk that came with manual copy/paste.",
], top=Inches(4.85), size=14, gap=8)
footer(s, 2)


# =====================================================================
# 3 — AGENDA
# =====================================================================
s = add_slide()
slide_title(s, "Agenda", "What I'll walk through")
items = [
    ("Where things stood", "The manual, invisible starting point"),
    ("What I built", "Three connected automations"),
    ("The impact", "Hours saved — with the math shown"),
    ("The hard parts", "Power BI, data, and production challenges"),
    ("What's next", "How this scales beyond me"),
]
x = Inches(0.7)
for i, (h, d) in enumerate(items):
    t = Inches(2.5) + i * Inches(0.92)
    badge = s.shapes.add_shape(MSO_SHAPE.OVAL, x, t, Inches(0.55), Inches(0.55))
    badge.fill.solid(); badge.fill.fore_color.rgb = RED; badge.line.fill.background()
    badge.shadow.inherit = False
    bp = badge.text_frame.paragraphs[0]; bp.alignment = PP_ALIGN.CENTER
    br = bp.add_run(); br.text = str(i + 1); br.font.size = Pt(18); br.font.bold = True
    br.font.color.rgb = WHITE; br.font.name = FONT
    tb, tf = textbox(s, x + Inches(0.8), t, Inches(10.5), Inches(0.6), valign=MSO_ANCHOR.MIDDLE)
    para(tf, h, size=17, color=INK, bold=True, first=True, space_after=1)
    para(tf, d, size=12.5, color=SOFT)
footer(s, 3)


# =====================================================================
# 4 — SITUATION
# =====================================================================
s = add_slide()
slide_title(s, "Where Things Stood", "Three manual processes, three sources of risk")
card(s, Inches(0.7), Inches(2.45), Inches(3.85), Inches(3.9), "AR Reporting",
     ["No consolidated daily view of receivables",
      "Aging, overdue, UAC lived across spreadsheets",
      "No trend, drill-down, or 'what-if' capability",
      "Leadership lacked a single place to look"], accent=RED)
card(s, Inches(4.74), Inches(2.45), Inches(3.85), Inches(3.9), "Daily Data File",
     ["A collector rebuilt it every business day",
      "Copy → paste → calculate → save-as, by hand",
      "~15 min/day, and only one person knew the steps",
      "Manual copy/paste → real error risk"], accent=RED)
card(s, Inches(8.78), Inches(2.45), Inches(3.82), Inches(3.9), "Statements (SOA)",
     ["Collectors generated SOAs in START by hand",
      "Then formatted, emailed, attached, CC'd",
      "Repeated weekly, per customer & country",
      "Missed sends & delays pulled focus off collections"], accent=RED)
footer(s, 4)


# =====================================================================
# 5 — WHAT I BUILT (overview)
# =====================================================================
s = add_slide()
slide_title(s, "What I Built", "Three connected automations — one AR workflow, digitized")
card(s, Inches(0.7), Inches(2.45), Inches(3.85), Inches(3.9), "1 · AR Dashboard",
     ["Live HTML report of the whole book",
      "Aging, UAC, gross sales, payments, trends",
      "Drill-downs + what-if payment simulators",
      "New visibility that didn't exist before"], accent=RED)
card(s, Inches(4.74), Inches(2.45), Inches(3.85), Inches(3.9), "2 · Daily Pipeline",
     ["Feeds the dashboard automatically",
      "Source file lands → data extracted & appended",
      "New dated file built with full history",
      "Runs hands-off, every business day"], accent=RED_DARK)
card(s, Inches(8.78), Inches(2.45), Inches(3.82), Inches(3.9), "3 · SOA Automation",
     ["6 weekly statements auto-generated",
      "Auto-emailed with attachments & CCs",
      "Configurable framework, not one-offs",
      "Collectors handle exceptions only"], accent=INK)
footer(s, 5)


# =====================================================================
# 6 — PILLAR 1: DASHBOARD
# =====================================================================
s = add_slide()
slide_title(s, "Pillar 1 · The AR Dashboard",
            "A live view of the receivables book — built from nothing")
bullets(s, [
    "Portfolio KPIs: Total AR, Overdue %, GT60 / GT90 aging, UAC, gross sales, payments.",
    "Drill-downs both ways — country ↔ customer — to see what's driving every number.",
    "Trends over time and same-month year-over-year comparison.",
    "What-if payment simulators: portfolio-level and customer-level, with manual allocation across aging bands and UAC.",
    "Editable risk thresholds and an Action-Required list mapped to the collections escalation process.",
], top=Inches(2.45), size=15, gap=11)
box = rrect(s, Inches(0.7), Inches(5.95), Inches(11.9), Inches(0.85), RGBColor(0xFC, 0xEF, 0xEF),
            line=True, radius=0.12)
tb, tf = textbox(s, Inches(1.0), Inches(6.05), Inches(11.3), Inches(0.65), valign=MSO_ANCHOR.MIDDLE)
para(tf, "Why it matters:  this reporting did not exist before — I created the visibility, "
         "not just automated an existing report.", size=13.5, color=RED_DARK, bold=True, first=True)
footer(s, 6)


# =====================================================================
# 7 — PILLAR 2: PIPELINE
# =====================================================================
s = add_slide()
slide_title(s, "Pillar 2 · The Daily Data Pipeline",
            "The manual daily file, now self-running")
tb, tf = textbox(s, Inches(0.7), Inches(2.4), Inches(5.7), Inches(0.4))
para(tf, "BEFORE  —  ~15 min/day, one person", size=13, color=RED, bold=True, first=True)
bullets(s, [
    "Open the emailed report, copy the data",
    "Paste into a template, calculate",
    "Copy specific columns into the workbook",
    "Refresh, then Save-As a new dated file",
], top=Inches(2.85), left=Inches(0.7), width=Inches(5.6), size=13, gap=7, color=SOFT)

tb, tf = textbox(s, Inches(6.85), Inches(2.4), Inches(5.7), Inches(0.4))
para(tf, "AFTER  —  hands-off", size=13, color=GOOD, bold=True, first=True)
bullets(s, [
    "File lands in the shared folder → flow fires",
    "Office Script reads & maps today's rows",
    "New dated file built, carrying full history",
    "Today's rows appended → dashboard updates",
], top=Inches(2.85), left=Inches(6.85), width=Inches(5.6), size=13, gap=7, color=SOFT)

box = rrect(s, Inches(0.7), Inches(5.55), Inches(11.9), Inches(1.25), INKCARD, radius=0.08)
tb, tf = textbox(s, Inches(1.0), Inches(5.72), Inches(11.3), Inches(0.95), valign=MSO_ANCHOR.MIDDLE)
para(tf, "Built with:  Power Query (M)  ·  Office Scripts (TypeScript)  ·  Power Automate  ·  SharePoint",
     size=13.5, color=YELLOW, bold=True, first=True, space_after=5)
para(tf, "Includes duplicate protection and graceful handling of missing / incomplete data (e.g. month-start EUR gaps).",
     size=12.5, color=RGBColor(0xCF, 0xD4, 0xDA))
footer(s, 7)


# =====================================================================
# 8 — PILLAR 3: SOA
# =====================================================================
s = add_slide()
slide_title(s, "Pillar 3 · SOA Automation",
            "Recurring statements generated & delivered without a collector touching them")
bullets(s, [
    "Automated 6 recurring weekly Statements of Account:  Arrow Electronics (5 countries) + Baker Hughes (1).",
    "Auto-generated from START templates with the correct account filters, then auto-emailed — attachments included, internal stakeholders CC'd.",
    "Built as a configurable framework: new customers onboard by updating configuration (accounts, recipients, schedule, template), not by rebuilding the solution.",
    "Exception-based by design — collectors step in only for exceptions, and spend their time on collections instead of routine report generation.",
], top=Inches(2.45), size=14.5, gap=11)
box = rrect(s, Inches(0.7), Inches(5.95), Inches(11.9), Inches(0.85), RGBColor(0xFC, 0xEF, 0xEF),
            line=True, radius=0.12)
tb, tf = textbox(s, Inches(1.0), Inches(6.05), Inches(11.3), Inches(0.65), valign=MSO_ANCHOR.MIDDLE)
para(tf, "Honest scope:  2 customers live today (6 SOAs/week) on a framework designed to scale to the rest — "
         "the hard part (the reusable engine) is done.", size=13, color=RED_DARK, bold=True, first=True)
footer(s, 8)


# =====================================================================
# 9 — THE NUMBERS (money slide)
# =====================================================================
s = add_slide()
slide_title(s, "The Impact", "Hours saved — with the math shown, conservatively")

rows = [
    ("", "Per week", "Per month", "Per year", False),
    ("Daily data pipeline  (15 min × 5 days)", "1.25 hrs", "~5.4 hrs", "~63 hrs", False),
    ("SOA automation  (6 SOAs × 30–60 min)", "3–6 hrs", "13–26 hrs", "150–300 hrs", False),
    ("Combined", "~4–7 hrs", "~18–31 hrs", "~210–360 hrs", True),
]
tw = Inches(11.9)
c0, c1, c2, c3 = Inches(5.0), Inches(2.3), Inches(2.3), Inches(2.3)
x0 = Inches(0.7)
y = Inches(2.45)
rh = Inches(0.72)
for i, (label, a, b, c, total) in enumerate(rows):
    ry = y + i * rh
    if i == 0:
        rect(s, x0, ry, tw, rh, INK)
        cells = [(c0, label, PP_ALIGN.LEFT), (c1, a, PP_ALIGN.CENTER),
                 (c2, b, PP_ALIGN.CENTER), (c3, c, PP_ALIGN.CENTER)]
        cx = x0
        for cw, txt, al in cells:
            tb, tf = textbox(s, cx + Inches(0.15), ry, cw - Inches(0.3), rh, valign=MSO_ANCHOR.MIDDLE)
            para(tf, txt, size=13, color=YELLOW, bold=True, first=True, align=al)
            cx += cw
    else:
        fill = RGBColor(0xFC, 0xEF, 0xEF) if total else (WHITE if i % 2 else RGBColor(0xF1, 0xEF, 0xEA))
        rect(s, x0, ry, tw, rh, fill, line=True)
        if total:
            rect(s, x0, ry, Pt(4), rh, RED)
        cells = [(c0, label, PP_ALIGN.LEFT), (c1, a, PP_ALIGN.CENTER),
                 (c2, b, PP_ALIGN.CENTER), (c3, c, PP_ALIGN.CENTER)]
        cx = x0
        for j, (cw, txt, al) in enumerate(cells):
            tb, tf = textbox(s, cx + Inches(0.15), ry, cw - Inches(0.3), rh, valign=MSO_ANCHOR.MIDDLE)
            para(tf, txt, size=13.5 if total else 13,
                 color=RED_DARK if total else INK,
                 bold=(total or j == 0), first=True, align=al)
            cx += cw

tb, tf = textbox(s, Inches(0.7), Inches(5.55), Inches(11.9), Inches(1.3))
para(tf, "≈ 5 to 9 full 40-hour work-weeks of collector time recovered per year.",
     size=16, color=RED_DARK, bold=True, first=True, space_after=8)
para(tf, "Assumptions (deliberately conservative):  5 business days/week; SOA at 30–60 min each, 6/week; "
         "figures EXCLUDE error-correction, rework, and downstream analysis time — which would push the number higher, not lower.",
     size=11.5, color=MUTED, italic=True)
footer(s, 9)


# =====================================================================
# 10 — THE HARD PARTS (credibility)
# =====================================================================
s = add_slide()
slide_title(s, "The Hard Parts", "What it actually took — and what I learned")
card(s, Inches(0.7), Inches(2.4), Inches(3.85), Inches(2.05), "Power BI & the data",
     ["Extensive trial-and-error building the model",
      "Fragmented, inconsistent source data",
      "Had to consolidate before anything was trustworthy"], accent=RED, title_size=13.5)
card(s, Inches(4.74), Inches(2.4), Inches(3.85), Inches(2.05), "Hosting blocked",
     ["Azure Functions blocked by permissions (RBAC)",
      "Re-architected the whole recurring-run approach",
      "Landed on GitHub Actions — zero-friction, no cost"], accent=RED, title_size=13.5)
card(s, Inches(8.78), Inches(2.4), Inches(3.82), Inches(2.05), "The right data pattern",
     ["Power Query replaces, it can't append",
      "Designed a carry-history-forward pipeline",
      "Solved it without breaking the live report"], accent=RED, title_size=13.5)
card(s, Inches(0.7), Inches(4.6), Inches(3.85), Inches(2.05), "A production incident",
     ["A live query edit went wrong",
      "Recovered cleanly via version history",
      "Adopted a test-first, never-in-prod discipline"], accent=INK, title_size=13.5)
card(s, Inches(4.74), Inches(4.6), Inches(3.85), Inches(2.05), "Connector quirks",
     ["Dozens of field-name / identifier mismatches",
      "Pagination, throttling, locked-file retries",
      "Debugged each to a clean, green run"], accent=INK, title_size=13.5)
card(s, Inches(8.78), Inches(4.6), Inches(3.82), Inches(2.05), "The takeaway",
     ["Shipped despite every obstacle",
      "Learned to de-risk before touching live data",
      "Left it documented and maintainable"], accent=INK, title_size=13.5)
footer(s, 10)


# =====================================================================
# 11 — SKILLS
# =====================================================================
s = add_slide()
slide_title(s, "What It Took", "Skills demonstrated, end to end")
groups = [
    ("Automation & Data", ["Power Query (M)", "Office Scripts (TypeScript)",
                            "Power Automate", "SharePoint / GitHub Actions"]),
    ("Analysis & Modeling", ["AR aging & cash-application logic", "Data consolidation & modeling",
                             "Power BI", "MTD / snapshot filtering"]),
    ("Product & Front-End", ["HTML / CSS / JavaScript", "SVG charting & dashboards",
                             "UX for non-technical users", "Clear documentation"]),
    ("Ways of Working", ["Reverse-engineering undocumented processes", "Building from ambiguity",
                         "Production troubleshooting", "Shipping, not just proposing"]),
]
for i, (title, items) in enumerate(groups):
    col = i % 2
    row = i // 2
    l = Inches(0.7) + col * Inches(6.1)
    t = Inches(2.5) + row * Inches(2.25)
    rrect(s, l, t, Inches(5.85), Inches(2.0), WHITE, line=True, radius=0.07)
    rect(s, l, t, Pt(4), Inches(2.0), RED)
    tb, tf = textbox(s, l + Inches(0.25), t + Inches(0.18), Inches(5.4), Inches(1.7))
    para(tf, title, size=14, color=RED_DARK, bold=True, first=True, space_after=7)
    for it in items:
        para(tf, it, size=12.5, color=INK, bullet=True, space_after=4)
footer(s, 11)


# =====================================================================
# 12 — WHAT'S NEXT
# =====================================================================
s = add_slide()
slide_title(s, "What's Next", "How this scales beyond me")
bullets(s, [
    "Expand the SOA framework to the remaining customers and countries — the reusable engine is already built.",
    "Migrate the data layer to a centralized database (SQL / Data Factory) — the scalable backbone for everything above.",
    "Publish leadership-facing Power BI dashboards on top of the same consolidated data.",
    "Hardening: standardize the monthly folder naming and add optional alerting for missing / late source files.",
], top=Inches(2.5), size=15.5, gap=13)
footer(s, 12)


# =====================================================================
# 13 — WHY ME (close)
# =====================================================================
s = add_slide(INK)
rect(s, 0, 0, SW, Inches(0.16), RED)
rect(s, 0, Inches(0.16), SW, Pt(3), YELLOW)
tb, tf = textbox(s, Inches(0.8), Inches(0.75), Inches(11), Inches(0.5))
para(tf, "WHY ME", size=13, color=YELLOW, bold=True, first=True)
tb, tf = textbox(s, Inches(0.8), Inches(1.35), Inches(11.7), Inches(1.0))
para(tf, "I don't just spot automation opportunities — I ship them.", size=30, color=WHITE,
     bold=True, first=True)
pts = [
    "Built net-new capability that didn't exist — and put it into production.",
    "Delivered quantifiable time back to the team (~210–360 hrs/yr) and cut real error & key-person risk.",
    "Worked through ambiguity, a hosting blocker, a production incident, and dozens of technical dead-ends — and still shipped.",
    "Left everything documented, tested, and running without me.",
]
tb, tf = textbox(s, Inches(0.8), Inches(2.7), Inches(11.7), Inches(3.2))
first = True
for p in pts:
    para(tf, p, size=16, color=RGBColor(0xE7, 0xEA, 0xED), bullet=True, first=first, space_after=15)
    first = False
box = rrect(s, Inches(0.8), Inches(6.05), Inches(11.7), Inches(0.85), RED, radius=0.12)
tb, tf = textbox(s, Inches(1.1), Inches(6.14), Inches(11.1), Inches(0.65), valign=MSO_ANCHOR.MIDDLE)
para(tf, "Give me the next problem — I'll ship the solution.", size=17, color=WHITE, bold=True,
     first=True)


# =====================================================================
# 14 — THANK YOU
# =====================================================================
s = add_slide()
rect(s, 0, Inches(3.5), SW, Pt(3), RED)
tb, tf = textbox(s, Inches(0.8), Inches(2.5), Inches(11.7), Inches(1.0), valign=MSO_ANCHOR.BOTTOM)
para(tf, "Thank you", size=40, color=INK, bold=True, first=True)
tb, tf = textbox(s, Inches(0.8), Inches(3.75), Inches(11.7), Inches(1.2))
para(tf, "Jean Kadadihi", size=18, color=INK, bold=True, first=True, space_after=3)
para(tf, "Finance Intern, Controlling  ·  DHL Express Americas — Key Account Desk (AMKAD)",
     size=13.5, color=SOFT, space_after=2)
para(tf, "Questions welcome.", size=13.5, color=RED_DARK, italic=True)
logo_placeholder(s, Inches(10.9), Inches(0.6))


out = "docs/AMKAD_Internship_Presentation.pptx"
prs.save(out)
print("Saved", out, "-", len(prs.slides._sldIdLst), "slides")
