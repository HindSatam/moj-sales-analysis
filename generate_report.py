"""Part 3 - quarterly sales market report (PowerPoint).

Usage:
    python generate_report.py --period 2020Q1

Reads cleansed_data.csv (made by clean_data.py) and writes output/sales_report_<period>.pptx.
"""
import argparse
import re
import sys
from datetime import date
from pathlib import Path

import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

BASE = Path(__file__).resolve().parent
MIN_DEALS = 10        # a change vs the previous quarter is only shown if both quarters have this many deals
TOP_DISTRICTS = 10
TOP_CATEGORIES = 8    # remaining categories are grouped as "Other"
CATEGORY_COL = "new_use"
FONT = "Arial"

NAVY = RGBColor(0x01, 0x2F, 0x55)     # colours taken from the rega logo
TEAL = RGBColor(0x00, 0x80, 0x8A)
GREY = RGBColor(0x5B, 0x67, 0x70)
LIGHT = RGBColor(0xF2, 0xF5, 0xF7)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DOWN = RGBColor(0xB4, 0x53, 0x1A)     # orange for decreases (with an arrow, so never colour alone)


# ----------------------------------------------------------------------------- numbers
def aggregate(df, col):
    """Deals, total value, total area and average price per m2 for each value of `col`."""
    g = df.groupby(col).agg(deals=("value", "size"), total_value=("value", "sum"), total_area=("AREA", "sum"))
    g["avg_price"] = g["total_value"] / g["total_area"]     # value-weighted price per m2
    return g


def pct_change(cur, prev, column):
    """% change vs the previous quarter; NaN when either quarter has fewer than MIN_DEALS deals."""
    if prev is None:
        return pd.Series(float("nan"), index=cur.index)
    p = prev.reindex(cur.index)
    enough = (cur["deals"] >= MIN_DEALS) & (p["deals"] >= MIN_DEALS)
    return (cur[column] / p[column] - 1).where(enough)


def fmt_money(v):
    if v >= 1e9:
        return f"SAR {v / 1e9:,.2f} B"
    if v >= 1e6:
        return f"SAR {v / 1e6:,.1f} M"
    return f"SAR {v:,.0f}"


def fmt_change(x):
    """(text, colour) for a % change. Arrow + colour, so it never relies on colour alone."""
    if pd.isna(x):
        return "n/a", GREY
    return (f"▲ {x * 100:.1f}%", TEAL) if x >= 0 else (f"▼ {abs(x) * 100:.1f}%", DOWN)


def label(period):
    return f"{period[:4]} {period[4:]}"       # 2020Q1 -> 2020 Q1


def prepare(df, col, keep=None):
    """Fill missing names with 'Unspecified'; if `keep` is given, group everything else as 'Other'."""
    names = df[col].fillna("Unspecified")
    if keep is not None:
        names = names.where(names.isin(keep), "Other")
    return df.assign(name=names)


# ----------------------------------------------------------------------------- slide helpers
def add_text(slide, text, x, y, w, h, size=14, bold=False, color=NAVY, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    box.text_frame.word_wrap = True
    p = box.text_frame.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name, run.font.size, run.font.bold, run.font.color.rgb = FONT, Pt(size), bold, color
    return box


def new_slide(prs, title, logo):
    slide = prs.slides.add_slide(prs.slide_layouts[6])      # blank layout
    add_text(slide, title, 0.6, 0.45, 9.4, 0.9, size=28, bold=True)
    if logo.exists():
        slide.shapes.add_picture(str(logo), Inches(10.5), Inches(0.4), width=Inches(2.25))
    add_text(slide, "Source: Ministry of Justice sale deeds, cleansed data", 0.6, 7.0, 10, 0.3, size=10, color=GREY)
    return slide


def add_table(slide, headers, rows, col_widths, x, y, row_h=0.42, header_h=0.65):
    """rows = list of lists; a cell is either text or (text, colour)."""
    shape = slide.shapes.add_table(len(rows) + 1, len(headers), Inches(x), Inches(y),
                                   Inches(sum(col_widths)), Inches(header_h + row_h * len(rows)))
    table = shape.table
    for i, w in enumerate(col_widths):
        table.columns[i].width = Inches(w)
    table.rows[0].height = Inches(header_h)
    for r in range(1, len(rows) + 1):
        table.rows[r].height = Inches(row_h)

    def fill(cell, text, colour, bold, background, align):
        cell.fill.solid()
        cell.fill.fore_color.rgb = background
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = cell.text_frame.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = text
        run.font.name, run.font.size, run.font.bold, run.font.color.rgb = FONT, Pt(12), bold, colour

    for c, header in enumerate(headers):
        fill(table.cell(0, c), header, WHITE, True, NAVY, PP_ALIGN.LEFT if c == 0 else PP_ALIGN.RIGHT)
    for r, row in enumerate(rows, start=1):
        for c, value in enumerate(row):
            text, colour = value if isinstance(value, tuple) else (value, NAVY)
            fill(table.cell(r, c), text, colour, False, LIGHT if r % 2 == 0 else WHITE,
                 PP_ALIGN.LEFT if c == 0 else PP_ALIGN.RIGHT)


def add_bar_chart(slide, names, values_m, x, y, w, h):
    data = CategoryChartData()
    data.categories = names
    data.add_series("Total sales value (SAR million)", values_m)
    chart = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(x), Inches(y), Inches(w), Inches(h), data).chart
    chart.has_legend = False
    chart.has_title = False
    chart.font.name, chart.font.size = FONT, Pt(12)
    plot = chart.plots[0]
    plot.gap_width = 60
    plot.series[0].format.fill.solid()
    plot.series[0].format.fill.fore_color.rgb = TEAL
    plot.has_data_labels = True
    plot.data_labels.number_format = "#,##0.0"
    plot.data_labels.number_format_is_linked = False
    plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    chart.category_axis.reverse_order = True          # largest bar at the top
    chart.category_axis.format.line.fill.background()
    chart.value_axis.visible = False
    chart.value_axis.has_major_gridlines = False


# ----------------------------------------------------------------------------- slides
def title_slide(prs, logo, period):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    if logo.exists():
        slide.shapes.add_picture(str(logo), Inches(0.8), Inches(0.8), width=Inches(4.2))
    add_text(slide, "Quarterly Sales Market Report", 0.8, 3.0, 11.5, 1.0, size=44, bold=True)
    add_text(slide, label(period), 0.8, 4.1, 11.5, 0.7, size=28, color=TEAL, bold=True)
    add_text(slide, f"Property sale transactions, Ministry of Justice   |   Prepared {date.today():%d %B %Y}",
             0.8, 5.0, 11.5, 0.5, size=16, color=GREY)


def overview_slide(prs, logo, period, prev_period, this, prev):
    slide = new_slide(prs, f"{label(period)} at a glance", logo)
    tot = aggregate(this.assign(all="all"), "all").iloc[0]
    prev_tot = aggregate(prev.assign(all="all"), "all").iloc[0] if prev is not None else None

    def change(column):
        return float("nan") if prev_tot is None else tot[column] / prev_tot[column] - 1

    sentence = (f"{tot['deals']:,.0f} property records were sold for a total of {fmt_money(tot['total_value'])}, "
                f"at an average of SAR {tot['avg_price']:,.0f} per m².")
    add_text(slide, sentence, 0.6, 1.5, 12.1, 0.9, size=18, color=GREY)

    cards = [("Average price per m²", f"SAR {tot['avg_price']:,.0f}", change("avg_price")),
             ("Total sales value", fmt_money(tot["total_value"]), change("total_value")),
             ("Transactions", f"{tot['deals']:,.0f}", change("deals"))]
    for i, (name, value, delta) in enumerate(cards):
        x = 0.6 + i * 4.1
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(2.7), Inches(3.8), Inches(3.2))
        card.fill.solid()
        card.fill.fore_color.rgb = LIGHT
        card.line.fill.background()
        card.shadow.inherit = False
        add_text(slide, name, x + 0.3, 3.0, 3.2, 0.5, size=16, color=GREY)
        add_text(slide, value, x + 0.3, 3.6, 3.3, 0.9, size=32, bold=True)
        text, colour = fmt_change(delta)
        add_text(slide, text, x + 0.3, 4.7, 3.2, 0.5, size=20, bold=True, color=colour)
        add_text(slide, f"vs {label(prev_period)}" if prev is not None else "no previous-quarter data",
                 x + 0.3, 5.2, 3.2, 0.4, size=14, color=GREY)


def breakdown_slides(prs, logo, period, prev_period, noun, this, prev, top_n=None):
    """A bar chart slide and a details table slide for districts or categories."""
    cur = aggregate(this, "name")
    prv = aggregate(prev, "name") if prev is not None else None
    cur["price_chg"] = pct_change(cur, prv, "avg_price")
    cur["value_chg"] = pct_change(cur, prv, "total_value")
    rows = cur.sort_values("total_value", ascending=False)
    if top_n:
        rows = rows.head(top_n)

    slide = new_slide(prs, f"Total sales value by {noun}", logo)
    add_text(slide, f"{label(period)}, SAR million" + (f" - top {top_n} by value" if top_n else "")
             + ". Deal counts are on the next slide.", 0.6, 1.4, 12, 0.4, size=14, color=GREY)
    add_bar_chart(slide, list(rows.index), list((rows["total_value"] / 1e6).round(1)), 0.6, 1.9, 12.1, 5.0)

    slide = new_slide(prs, f"Prices and deals by {noun}", logo)
    add_text(slide, f"{label(period)}. Change is against {label(prev_period)}; n/a means fewer than {MIN_DEALS} deals "
             f"in either quarter or no previous data.", 0.6, 1.4, 12.1, 0.4, size=14, color=GREY)
    table_rows = [[name, f"{r.deals:,.0f}", f"{r.avg_price:,.0f}", fmt_change(r.price_chg),
                   f"{r.total_value / 1e6:,.1f}", fmt_change(r.value_chg)] for name, r in rows.iterrows()]
    add_table(slide, [noun.capitalize(), "Deals", "Avg price per m² (SAR)", "Change", "Total value (SAR M)", "Change"],
              table_rows, [3.6, 1.3, 2.3, 1.5, 2.3, 1.5], 0.6, 1.95, row_h=0.4 if len(table_rows) > 8 else 0.5)


def notes_slide(prs, logo, prev_period):
    slide = new_slide(prs, "How to read this report", logo)
    lines = [
        "Data: property sale deeds from the Ministry of Justice, after cleansing (duplicates, invalid dates, "
        "missing or zero prices and extreme outliers removed).",
        "Deals: number of property records sold. One sale deed can include more than one property.",
        "Total sales value: area (m²) multiplied by price per m², summed over the deals.",
        "Average price per m²: total sales value divided by total area sold.",
        f"Change: this quarter compared with {label(prev_period)}. Shown only when both quarters have at least "
        f"{MIN_DEALS} deals.",
        "Unspecified: the source did not record a district or category for those deals.",
        f"Categories: the top {TOP_CATEGORIES} by value are shown; the rest are grouped as Other.",
    ]
    box = slide.shapes.add_textbox(Inches(0.6), Inches(1.6), Inches(12.1), Inches(5.0))
    box.text_frame.word_wrap = True
    for i, line in enumerate(lines):
        p = box.text_frame.paragraphs[0] if i == 0 else box.text_frame.add_paragraph()
        p.space_after = Pt(10)
        run = p.add_run()
        run.text = "•  " + line
        run.font.name, run.font.size, run.font.color.rgb = FONT, Pt(16), NAVY


# ----------------------------------------------------------------------------- main
def main():
    parser = argparse.ArgumentParser(description="Generate the quarterly sales market report.")
    parser.add_argument("--period", required=True, help="quarter to report on, e.g. 2020Q1")
    parser.add_argument("--input", default=str(BASE / "cleansed_data.csv"))
    parser.add_argument("--output", help="default: output/sales_report_<period>.pptx")
    parser.add_argument("--logo", default=str(BASE / "assets" / "rega_logo.png"))
    args = parser.parse_args()

    period = args.period.upper()
    if not re.fullmatch(r"\d{4}Q[1-4]", period):
        sys.exit(f"Period must look like 2020Q1 (year, Q, quarter 1-4); got '{args.period}'.")

    df = pd.read_csv(args.input, encoding="utf-8-sig")
    quarters = sorted(df["quarter"].unique())
    if period not in quarters:
        sys.exit(f"No data for {period}. Quarters available: {', '.join(quarters)}")

    prev_period = str(pd.Period(period, freq="Q") - 1)
    this_raw, prev_raw = df[df["quarter"] == period], df[df["quarter"] == prev_period]
    prev_raw = prev_raw if len(prev_raw) else None
    if prev_raw is None:
        print(f"Note: no data for the previous quarter ({prev_period}); changes will show n/a.")

    # districts: top 10 by value; categories: top 8 by value, the rest grouped as "Other"
    districts = lambda d: prepare(d, "NEIGHBORHOOD_NAME")
    keep = prepare(this_raw, CATEGORY_COL).groupby("name")["value"].sum().nlargest(TOP_CATEGORIES).index
    categories = lambda d: prepare(d, CATEGORY_COL, keep)

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    logo = Path(args.logo)
    if not logo.exists():
        print(f"Warning: logo not found at {logo}; the report will have no logo.")

    title_slide(prs, logo, period)
    overview_slide(prs, logo, period, prev_period, this_raw, prev_raw)
    breakdown_slides(prs, logo, period, prev_period, "district", districts(this_raw),
                     districts(prev_raw) if prev_raw is not None else None, TOP_DISTRICTS)
    breakdown_slides(prs, logo, period, prev_period, "category", categories(this_raw),
                     categories(prev_raw) if prev_raw is not None else None)
    notes_slide(prs, logo, prev_period)

    output = Path(args.output) if args.output else BASE / "output" / f"sales_report_{period}.pptx"
    output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output)
    print(f"Saved {output}")


if __name__ == "__main__":
    main()
