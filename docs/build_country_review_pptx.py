#!/usr/bin/env python3
"""
Build the AMKAD Country Performance Review deck (the "Key Accounts Desk"
format) from the same JSON the Office Script already produces for the HTML
report.

Usage:
    python3 build_country_review_pptx.py <data.json> <output.pptx>

Input JSON shape (subset of what amkad_ar_report.ts already returns):
{
  "CountryName": "Mexico",
  "ReportMonth": "May 2026",
  "TrendDataJson": [...],              // from the script's TrendDataJson
  "CountryDrilldownJson": [...],       // from the script's CountryDrilldownJson
  "CustomerDrilldownJson": [...]       // from the script's CustomerDrilldownJson
}

Sections left intentionally blank (require human judgement, not derivable
from AR figures): "Main Issues", "Current Actions / Support Needed",
"Cash Forecasting / Payments Summary", AMKAD POC name, Go-Live date.
"""
import json
import sys
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

RED = RGBColor(0xD4, 0x05, 0x11)
YELLOW = RGBColor(0xFF, 0xCC, 0x00)
INK = RGBColor(0x1A, 0x1D, 0x21)
SOFT = RGBColor(0x5B, 0x64, 0x70)
MUTED = RGBColor(0x8A, 0x83, 0x7B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BG = RGBColor(0xF7, 0xF6, 0xF3)
LINE = RGBColor(0xE2, 0xE0, 0xDA)
PLACEHOLDER_FILL = RGBColor(0xFB, 0xF7, 0xEA)
FONT = "Calibri"


def euro(v):
    try:
        return "€" + format(round(v), ",")
    except Exception:
        return "€0"


def pct(v):
    try:
        return f"{v * 100:.1f}%"
    except Exception:
        return "0.0%"


class Deck:
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)
        self.blank = self.prs.slide_layouts[6]
        self.sw, self.sh = self.prs.slide_width, self.prs.slide_height

    def add_slide(self, bg=BG):
        s = self.prs.slides.add_slide(self.blank)
        r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, self.sw, self.sh)
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


def textbox(s, l, t, w, h, valign=MSO_ANCHOR.TOP):
    tb = s.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = valign
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    return tb, tf


def para(tf, text, size=14, color=INK, bold=False, italic=False, align=PP_ALIGN.LEFT,
         space_after=6, first=False, font=FONT):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.space_after = Pt(space_after)
    r = p.add_run(); r.text = text
    r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
    r.font.color.rgb = color; r.font.name = font
    return p


def placeholder_box(s, l, t, w, h, label):
    """Marked box for content that needs a human (commentary, forecasts, POC names)."""
    box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    try:
        box.adjustments[0] = 0.08
    except Exception:
        pass
    box.fill.solid(); box.fill.fore_color.rgb = PLACEHOLDER_FILL
    box.line.color.rgb = YELLOW; box.line.width = Pt(1)
    box.shadow.inherit = False
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = Pt(8); tf.margin_right = Pt(8); tf.margin_top = Pt(6)
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = label
    r.font.size = Pt(9); r.font.italic = True; r.font.color.rgb = MUTED; r.font.name = FONT
    return box


def data_table(s, l, t, w, h, headers, rows):
    n_rows = len(rows) + 1
    n_cols = len(headers)
    gtable = s.shapes.add_table(n_rows, n_cols, l, t, w, h)
    table = gtable.table
    for i, h_ in enumerate(headers):
        cell = table.cell(0, i)
        cell.fill.solid(); cell.fill.fore_color.rgb = RED
        tf = cell.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = h_
        r.font.size = Pt(10.5); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = FONT
    for ri, row in enumerate(rows, start=1):
        for ci, val in enumerate(row):
            cell = table.cell(ri, ci)
            cell.fill.solid(); cell.fill.fore_color.rgb = WHITE if ri % 2 else RGBColor(0xF6, 0xF8, 0xFB)
            tf = cell.text_frame
            p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER if ci > 0 else PP_ALIGN.LEFT
            r = p.add_run(); r.text = str(val)
            r.font.size = Pt(10.5); r.font.color.rgb = INK; r.font.name = FONT
            r.font.bold = (ci == 0)
    return gtable


def build(data, out_path):
    deck = Deck()
    country_name = data.get("CountryName", "Country")
    report_month = data.get("ReportMonth", "")
    trend = data.get("TrendDataJson", [])
    countries = data.get("CountryDrilldownJson", [])
    customers = data.get("CustomerDrilldownJson", [])

    # ---------------- Slide 1: Title ----------------
    s = deck.add_slide(bg=INK)
    rect(s, 0, 0, deck.sw, Pt(10), YELLOW)
    rect(s, 0, Pt(10), deck.sw, Pt(5), RED)
    tb, tf = textbox(s, Inches(0.9), Inches(2.7), Inches(11), Inches(0.5))
    para(tf, "KEY ACCOUNTS DESK", size=14, color=YELLOW, bold=True, first=True)
    tb, tf = textbox(s, Inches(0.9), Inches(3.15), Inches(11.5), Inches(1.1))
    para(tf, "Country Performance Review", size=36, color=WHITE, bold=True, first=True)
    tb, tf = textbox(s, Inches(0.9), Inches(4.3), Inches(11), Inches(0.8))
    para(tf, country_name, size=20, color=WHITE, bold=True, first=True, space_after=2)
    para(tf, report_month, size=15, color=RGBColor(0xC8, 0xC8, 0xC8))

    # ---------------- Slide 2: Agenda ----------------
    s = deck.add_slide()
    tb, tf = textbox(s, Inches(0.7), Inches(0.6), Inches(8), Inches(0.7))
    para(tf, "Agenda", size=28, color=INK, bold=True, first=True)
    items = [("1", "OVERALL COUNTRY PERFORMANCE"), ("2", "CUSTOMER PERFORMANCE"), ("3", "SPECIAL TOPICS")]
    top = Inches(1.9)
    for i, (num, label) in enumerate(items):
        y = top + i * Inches(1.3)
        badge = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.9), y, Inches(0.6), Inches(0.6))
        badge.fill.solid(); badge.fill.fore_color.rgb = RED; badge.line.fill.background(); badge.shadow.inherit = False
        btf = badge.text_frame; btf.margin_left = 0
        bp = btf.paragraphs[0]; bp.alignment = PP_ALIGN.CENTER
        br = bp.add_run(); br.text = num; br.font.size = Pt(18); br.font.bold = True; br.font.color.rgb = WHITE; br.font.name = FONT
        bar = rect(s, Inches(1.75), y + Inches(0.12), Inches(9.5), Inches(0.36), RGBColor(0xF2, 0xEF, 0xEA))
        tb2, tf2 = textbox(s, Inches(1.95), y, Inches(9), Inches(0.6), valign=MSO_ANCHOR.MIDDLE)
        para(tf2, label, size=15, color=INK, bold=True, first=True)

    # ---------------- Section divider ----------------
    def divider(title):
        s = deck.add_slide(bg=INK)
        rect(s, 0, 0, deck.sw, Pt(10), YELLOW)
        rect(s, 0, Pt(10), deck.sw, Pt(5), RED)
        tb, tf = textbox(s, Inches(0.9), Inches(3.3), Inches(11), Inches(1))
        para(tf, title, size=32, color=WHITE, bold=True, first=True)
        return s

    divider("Country Performance")

    # ---------------- Country performance data slide ----------------
    s = deck.add_slide()
    tb, tf = textbox(s, Inches(0.7), Inches(0.5), Inches(10), Inches(0.6))
    para(tf, "Overall Country Performance — " + country_name, size=22, color=INK, bold=True, first=True)

    country_row = next((c for c in countries if c.get("country", "").lower() in country_name.lower()
                         or country_name.lower() in c.get("country", "").lower()), None)
    if country_row:
        stats = [
            ("Total AR", euro(country_row.get("totalAR", 0))),
            ("Overdue", euro(country_row.get("overdue", 0)) + "  (" + pct(country_row.get("overduePct", 0)) + ")"),
            ("GT60", euro(country_row.get("gt60", 0)) + "  (" + pct(country_row.get("gt60Pct", 0)) + ")"),
            ("GT90", euro(country_row.get("gt90", 0)) + "  (" + pct(country_row.get("gt90Pct", 0)) + ")"),
        ]
        x = Inches(0.7)
        for label, val in stats:
            card = rect(s, x, Inches(1.3), Inches(2.9), Inches(1.2), WHITE, line=True)
            tb2, tf2 = textbox(s, x + Inches(0.15), Inches(1.42), Inches(2.6), Inches(1))
            para(tf2, label.upper(), size=10.5, color=SOFT, bold=True, first=True, space_after=4)
            para(tf2, val, size=15, color=INK, bold=True)
            x += Inches(3.05)

    if trend:
        months = [d.get("month", "") for d in trend][-12:]
        ar_vals = [d.get("totalAR", 0) for d in trend][-12:]
        chart_data = CategoryChartData()
        chart_data.categories = months
        chart_data.add_series("Total AR", ar_vals)
        s.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Inches(0.7), Inches(2.8), Inches(11.9), Inches(4.2), chart_data)

    divider("Customer Performance")

    # ---------------- Per-customer slide pairs ----------------
    by_customer = {}
    for row in customers:
        cust = row.get("customer", "Unknown")
        by_customer.setdefault(cust, []).append(row)

    for cust, rows in by_customer.items():
        agg = {
            "totalAR": sum(r.get("totalAR", 0) for r in rows),
            "overdue": sum(r.get("overdue", 0) for r in rows),
            "gt60": sum(r.get("gt60", 0) for r in rows),
            "gt90": sum(r.get("gt90", 0) for r in rows),
        }
        if agg["totalAR"] <= 0:
            continue

        # Cover slide
        s = deck.add_slide()
        tb, tf = textbox(s, Inches(0.7), Inches(0.5), Inches(10), Inches(0.7))
        para(tf, cust, size=24, color=INK, bold=True, first=True)
        placeholder_box(s, Inches(0.7), Inches(1.4), Inches(4), Inches(0.6), "Net terms / Go-Live date — fill in")
        placeholder_box(s, Inches(0.7), Inches(2.2), Inches(4), Inches(0.7), "AMKAD POC: — fill in")

        # Detail slide
        s = deck.add_slide()
        tb, tf = textbox(s, Inches(0.7), Inches(0.45), Inches(11.5), Inches(0.6))
        para(tf, cust.upper() + " – ACCOUNT BUSINESS REVIEW", size=18, color=INK, bold=True, first=True)
        tb2, tf2 = textbox(s, Inches(0.7), Inches(0.95), Inches(11.5), Inches(0.4))
        para(tf2, "TTL AR: " + euro(agg["totalAR"]) + "   |   Past Due: " + euro(agg["overdue"]) +
             "   |   Over 90: " + euro(agg["gt90"]), size=13, color=SOFT, first=True)

        col_w = Inches(3.85)
        placeholder_box(s, Inches(0.7), Inches(1.5), col_w, Inches(2.6), "MAIN ISSUES — fill in")
        placeholder_box(s, Inches(4.65), Inches(1.5), col_w, Inches(2.6), "CURRENT ACTIONS / SUPPORT NEEDED — fill in")
        placeholder_box(s, Inches(8.6), Inches(1.5), col_w, Inches(2.6), "CASH FORECASTING / PAYMENTS — fill in")

        current = agg["totalAR"] - agg["overdue"]
        bucket_0_60 = max(agg["overdue"] - agg["gt60"], 0)
        bucket_60_90 = max(agg["gt60"] - agg["gt90"], 0)
        data_table(s, Inches(0.7), Inches(4.35), Inches(11.9), Inches(1.0),
                   ["FINANCIALS", "Total AR", "Current", "Overdue", "0-60", "60-90", ">90"],
                   [["", euro(agg["totalAR"]), euro(current), euro(agg["overdue"]),
                     euro(bucket_0_60), euro(bucket_60_90), euro(agg["gt90"])]])

    deck.prs.save(out_path)
    print("Saved", out_path, "-", len(deck.prs.slides.__iter__.__self__._sldIdLst), "slides")


if __name__ == "__main__":
    data_path = sys.argv[1] if len(sys.argv) > 1 else "sample_data.json"
    out_path = sys.argv[2] if len(sys.argv) > 2 else "AMKAD_Country_Review.pptx"
    with open(data_path) as f:
        data = json.load(f)
    build(data, out_path)
