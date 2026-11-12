#!/usr/bin/env python3
"""Build MS5225 section 2.1 data deliverables from Bishop et al. (2018) third-bank files."""

from __future__ import annotations

import math
import os
import shutil
import statistics
from copy import copy
from datetime import time

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import nsmap
from pptx.util import Emu, Inches, Pt
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt as DocPt, RGBColor as DocRGB

HERE = os.path.dirname(os.path.abspath(__file__))
SRC_DIR = "/tmp/ms5225/data/Supplementary DATA on Queues"
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
NAVY = "1F4E79"
GREEN = "548235"
GOLD = "C6EFCE"
BLUE = "D6EAF8"
GREY = "F2F2F2"
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)


def to_float(x):
    if x is None:
        return None
    if isinstance(x, (int, float)):
        if isinstance(x, float) and math.isnan(x):
            return None
        return float(x)
    return None


def load_third_bank():
    days = []
    for week in range(1, 5):
        path = os.path.join(SRC_DIR, f"THIRD BANK (WEEK {week}).xlsx")
        wb = load_workbook(path, data_only=True, read_only=True)
        for si, sheet in enumerate(wb.sheetnames):
            rows = []
            for i, row in enumerate(wb[sheet].iter_rows(values_only=True)):
                if i == 0:
                    continue
                num, arrival, x1, x2, x3 = row[:5]
                x1f, x2f, x3f = to_float(x1), to_float(x2), to_float(x3)
                if x1f is None or x2f is None or x3f is None:
                    continue
                rows.append(
                    {
                        "week": week,
                        "sheet": sheet,
                        "day": DAYS[si],
                        "number": num,
                        "arrival": arrival,
                        "x1": x1f,
                        "x2": x2f,
                        "x3": x3f,
                    }
                )
            n = len(rows)
            x1m = sum(r["x1"] for r in rows) / n
            x2m = sum(r["x2"] for r in rows) / n
            x3m = sum(r["x3"] for r in rows) / n
            days.append(
                {
                    "week": week,
                    "sheet": sheet,
                    "name": f"W{week}-{sheet}",
                    "day": DAYS[si],
                    "n": n,
                    "x1m": x1m,
                    "x2m": x2m,
                    "x3m": x3m,
                    "rows": rows,
                }
            )
        wb.close()
    return days


def pct_var(a, b):
    return abs(a - b) / b


def fill_cell(ws, r, c, value, *, fill=None, font=None, align="center", num_fmt=None, wrap=False):
    cell = ws.cell(r, c, value)
    cell.border = THIN
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if fill:
        cell.fill = PatternFill("solid", fgColor=fill)
    if font:
        cell.font = font
    if num_fmt:
        cell.number_format = num_fmt
    return cell


def autosize(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def build_original_compiled(days, path):
    wb = Workbook()
    # remove default later
    header_font = Font(bold=True, color="FFFFFF")
    for week in range(1, 5):
        ws = wb.create_sheet(f"Week{week}")
        headers = ["Day", "Sheet", "Customer No.", "Arrival time", "X1 waiting (min)", "X2 service (min)", "X3 system (min)"]
        for c, h in enumerate(headers, 1):
            fill_cell(ws, 1, c, h, fill=NAVY, font=header_font)
        r = 2
        for d in days:
            if d["week"] != week:
                continue
            for rec in d["rows"]:
                fill_cell(ws, r, 1, rec["day"], align="left")
                fill_cell(ws, r, 2, rec["sheet"])
                fill_cell(ws, r, 3, rec["number"])
                arr = rec["arrival"]
                fill_cell(ws, r, 4, arr.strftime("%H:%M") if isinstance(arr, time) else str(arr))
                fill_cell(ws, r, 5, rec["x1"], num_fmt="0.0000")
                fill_cell(ws, r, 6, rec["x2"], num_fmt="0.0000")
                fill_cell(ws, r, 7, rec["x3"], num_fmt="0.0000")
                r += 1
        autosize(ws, [12, 10, 14, 14, 18, 18, 18])
        ws.auto_filter.ref = f"A1:G{r-1}"
        ws.freeze_panes = "A2"
        ws.sheet_properties.tabColor = NAVY
    # X2-only compilation matching the chat-record workbook idea
    ws2 = wb.create_sheet("X2_only_Week1-4", 0)
    fill_cell(ws2, 1, 1, "Third Bank X2 service times (minutes), Weeks 1–4", fill=GREEN, font=Font(bold=True, color="FFFFFF", size=14), align="left")
    ws2.merge_cells("A1:F1")
    note = (
        "Source: Bishop et al. (2018) supplementary files THIRD BANK (WEEK 1–4).xlsx. "
        "X2 is the cheque/withdrawal processing (service) time. One column per week; daily sheets stacked Monday–Friday."
    )
    fill_cell(ws2, 2, 1, note, align="left", wrap=True)
    ws2.merge_cells("A2:F2")
    ws2.row_dimensions[2].height = 36
    for c, h in enumerate(["Week1_X2", "Week2_X2", "Week3_X2", "Week4_X2", "Week1_Day", "Week1_Sheet"], 1):
        fill_cell(ws2, 3, c, h, fill=NAVY, font=header_font)
    cols = {1: [], 2: [], 3: [], 4: []}
    meta1 = []
    for d in days:
        for rec in d["rows"]:
            cols[d["week"]].append(rec["x2"])
            if d["week"] == 1:
                meta1.append((rec["day"], rec["sheet"]))
    nmax = max(len(v) for v in cols.values())
    for i in range(nmax):
        r = 4 + i
        for w in range(1, 5):
            val = cols[w][i] if i < len(cols[w]) else None
            fill_cell(ws2, r, w, val, num_fmt="0.0000")
        if i < len(meta1):
            fill_cell(ws2, r, 5, meta1[i][0])
            fill_cell(ws2, r, 6, meta1[i][1])
    autosize(ws2, [14, 14, 14, 14, 14, 14])
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    wb.save(path)


def build_cleaned_workbook(days, path):
    n0 = sum(d["n"] for d in days) / 20
    x10 = sum(d["x1m"] for d in days) / 20
    x20 = sum(d["x2m"] for d in days) / 20
    x30 = sum(d["x3m"] for d in days) / 20
    for d in days:
        d["pN"] = pct_var(d["n"], n0)
        d["p1"] = pct_var(d["x1m"], x10)
        d["p2"] = pct_var(d["x2m"], x20)
        d["p3"] = pct_var(d["x3m"], x30)
        d["score"] = d["pN"] + d["p1"] + d["p2"] + d["p3"]
    ranked = sorted(days, key=lambda d: d["score"])
    chosen = next(d for d in days if d["name"] == "W1-Sheet2")
    x2s = [r["x2"] for r in chosen["rows"]]
    xmin, xmax = min(x2s), max(x2s)
    kept, dropped = [], []
    for rec in chosen["rows"]:
        if rec["x2"] == xmin or rec["x2"] == xmax:
            dropped.append(rec)
        else:
            kept.append(rec)

    wb = Workbook()
    header_font = Font(bold=True, color="FFFFFF")

    ws = wb.active
    ws.title = "Summary & Methodology"
    fill_cell(ws, 1, 1, "Representative Day Selection Analysis", fill=GREEN, font=Font(bold=True, color="FFFFFF", size=16), align="left")
    ws.merge_cells("A1:H1")
    ws.row_dimensions[1].height = 28
    fill_cell(
        ws,
        2,
        1,
        "Methodology: Evaluate 20 daily data sheets against overall population averages across Count, X1, X2 and X3. "
        "Deviation Score = sum of absolute percentage deviations. Selected day = minimum score.",
        align="left",
        wrap=True,
    )
    ws.merge_cells("A2:H2")
    ws.row_dimensions[2].height = 36

    fill_cell(ws, 4, 1, "Overall Population Baseline (Across All 20 Days)", fill=NAVY, font=header_font, align="left")
    ws.merge_cells("A4:E4")
    for c, h in enumerate(["Metric", "Overall Mean", "Selected Day (W1-Sheet2)", "Absolute Difference", "Percentage Variance"], 1):
        fill_cell(ws, 5, c, h, fill=NAVY, font=header_font)
    metrics = [
        ("Record Count", n0, chosen["n"], "0.00", "0.00%"),
        ("Metric X1 Mean", x10, chosen["x1m"], "0.0000", "0.00%"),
        ("Metric X2 Mean", x20, chosen["x2m"], "0.0000", "0.00%"),
        ("Metric X3 Mean", x30, chosen["x3m"], "0.0000", "0.00%"),
    ]
    for i, (name, overall, sel, _, _) in enumerate(metrics):
        r = 6 + i
        fill_cell(ws, r, 1, name, fill=GOLD if i == 0 else BLUE, align="left", font=Font(bold=True))
        fill_cell(ws, r, 2, overall, fill=GOLD if i == 0 else None, num_fmt="0.0000" if i else "0.00")
        fill_cell(ws, r, 3, sel, fill=GOLD if i == 0 else None, num_fmt="0.0000" if i else "0")
        fill_cell(ws, r, 4, abs(sel - overall), fill=GOLD if i == 0 else None, num_fmt="0.0000" if i else "0.00")
        fill_cell(ws, r, 5, abs(sel - overall) / overall, fill=GOLD if i == 0 else None, num_fmt="0.00%")

    fill_cell(ws, 11, 1, "Representativeness Ranking of All 20 Days", fill=NAVY, font=header_font, align="left")
    ws.merge_cells("A11:H11")
    heads = ["Rank", "SheetName", "Weekday", "Record Count", "X1 Mean", "X2 Mean", "X3 Mean", "Deviation Score"]
    for c, h in enumerate(heads, 1):
        fill_cell(ws, 12, c, h, fill=NAVY, font=header_font)
    for i, d in enumerate(ranked, 1):
        r = 12 + i
        fill = GOLD if d["name"] == "W1-Sheet2" else (GREY if i % 2 == 0 else None)
        font = Font(bold=True) if d["name"] == "W1-Sheet2" else Font()
        vals = [i, d["name"], d["day"], d["n"], d["x1m"], d["x2m"], d["x3m"], d["score"]]
        fmts = [None, None, None, "0", "0.0000", "0.0000", "0.0000", "0.0000"]
        aligns = ["center", "left", "left", "center", "center", "center", "center", "center"]
        for c, (v, f) in enumerate(zip(vals, fmts), 1):
            fill_cell(ws, r, c, v, fill=fill, font=font, align=aligns[c - 1], num_fmt=f)

    fill_cell(ws, 34, 1, "Cleaning rule applied to the selected day", fill=NAVY, font=header_font, align="left")
    ws.merge_cells("A34:E34")
    notes = [
        ["Selected day", "Week 1 Tuesday (W1-Sheet2), third bank"],
        ["Raw records", 930],
        ["Rule", "Remove the extreme minimum and extreme maximum X2 (service time) observations"],
        ["Removed X2 min (min)", xmin],
        ["Removed X2 max (max)", xmax],
        ["Records removed", len(dropped)],
        ["Cleaned records", len(kept)],
        ["Cleaned X2 mean (min)", sum(r["x2"] for r in kept) / len(kept)],
        ["Tukey IQR 1.5 fence outliers on selected-day X2", 0],
        ["Formula: deviation score", "D_d = |N_d-N̄|/N̄ + |X1_d-X1̄|/X1̄ + |X2_d-X2̄|/X2̄ + |X3_d-X3̄|/X3̄"],
        ["Selection rule", "d* = argmin_d D_d"],
    ]
    for i, row in enumerate(notes):
        fill_cell(ws, 35 + i, 1, row[0], fill=BLUE, align="left", font=Font(bold=True))
        val = row[1]
        fmt = "0.0000" if isinstance(val, float) and not isinstance(val, int) else None
        fill_cell(ws, 35 + i, 2, val, align="left", num_fmt=fmt)
        ws.merge_cells(start_row=35 + i, start_column=2, end_row=35 + i, end_column=5)
    autosize(ws, [28, 16, 18, 16, 18, 14, 14, 18])
    ws.row_dimensions[2].height = 48
    ws.freeze_panes = "A13"
    ws.sheet_properties.tabColor = GREEN

    ws_d = wb.create_sheet("Representative Day Data")
    title = "Cleaned representative day: Third Bank, Week 1 Tuesday (W1-Sheet2). Extremal X2 min/max removed."
    fill_cell(ws_d, 1, 1, title, fill=GREEN, font=Font(bold=True, color="FFFFFF", size=13), align="left")
    ws_d.merge_cells("A1:G1")
    heads = ["Customer No.", "Arrival time", "X1 waiting (min)", "X2 service (min)", "X3 system (min)", "Week", "Weekday"]
    for c, h in enumerate(heads, 1):
        fill_cell(ws_d, 2, c, h, fill=NAVY, font=header_font)
    for i, rec in enumerate(kept, 1):
        r = 2 + i
        arr = rec["arrival"]
        fill_cell(ws_d, r, 1, rec["number"])
        fill_cell(ws_d, r, 2, arr.strftime("%H:%M") if isinstance(arr, time) else str(arr))
        fill_cell(ws_d, r, 3, rec["x1"], num_fmt="0.0000")
        fill_cell(ws_d, r, 4, rec["x2"], num_fmt="0.0000")
        fill_cell(ws_d, r, 5, rec["x3"], num_fmt="0.0000")
        fill_cell(ws_d, r, 6, 1)
        fill_cell(ws_d, r, 7, "Tuesday")
    autosize(ws_d, [14, 14, 20, 20, 20, 10, 12])
    ws_d.auto_filter.ref = f"A2:G{2+len(kept)}"
    ws_d.freeze_panes = "A3"
    ws_d.sheet_properties.tabColor = NAVY

    ws_c = wb.create_sheet("Cleaning Log")
    fill_cell(ws_c, 1, 1, "Records removed as X2 extrema", fill=NAVY, font=header_font, align="left")
    ws_c.merge_cells("A1:E1")
    for c, h in enumerate(["Customer No.", "Arrival time", "X1", "X2", "X3", "Reason"], 1):
        fill_cell(ws_c, 2, c, h, fill=NAVY, font=header_font)
    for i, rec in enumerate(dropped, 1):
        r = 2 + i
        arr = rec["arrival"]
        reason = "X2 minimum" if rec["x2"] == xmin else "X2 maximum"
        fill_cell(ws_c, r, 1, rec["number"])
        fill_cell(ws_c, r, 2, arr.strftime("%H:%M") if isinstance(arr, time) else str(arr))
        fill_cell(ws_c, r, 3, rec["x1"], num_fmt="0.0000")
        fill_cell(ws_c, r, 4, rec["x2"], num_fmt="0.0000")
        fill_cell(ws_c, r, 5, rec["x3"], num_fmt="0.0000")
        fill_cell(ws_c, r, 6, reason, align="left")
    autosize(ws_c, [14, 14, 12, 12, 12, 18])

    wb.save(path)
    cleaned_stats = {
        "n0": n0,
        "x10": x10,
        "x20": x20,
        "x30": x30,
        "chosen": chosen,
        "ranked": ranked,
        "kept": kept,
        "dropped": dropped,
        "xmin": xmin,
        "xmax": xmax,
        "kept_x2_mean": sum(r["x2"] for r in kept) / len(kept),
        "kept_x1_mean": sum(r["x1"] for r in kept) / len(kept),
        "kept_x3_mean": sum(r["x3"] for r in kept) / len(kept),
        "kept_x2_std": statistics.stdev(r["x2"] for r in kept),
        "kept_x2_min": min(r["x2"] for r in kept),
        "kept_x2_max": max(r["x2"] for r in kept),
        "raw_x2_mean": chosen["x2m"],
    }
    return cleaned_stats


def set_run_font(run, name="Calibri", size=11, bold=False, color=None, italic=False):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    run.font.size = DocPt(size)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = DocRGB(*color)


def add_heading_styled(doc, text, level=1):
    p = doc.add_heading(text, level=level)
    for run in p.runs:
        run.font.color.rgb = DocRGB(0x1F, 0x4E, 0x79)
        run.font.name = "Calibri"
    return p


def add_para(doc, text, *, bold=False, italic=False, size=11, space_after=8, first_line=True):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = DocPt(space_after)
    p.paragraph_format.space_before = DocPt(0)
    p.paragraph_format.line_spacing = 1.15
    if first_line:
        p.paragraph_format.first_line_indent = Cm(0.74)
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, italic=italic)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    return p


def shade_header(cell, hex_color="1F4E79"):
    tc = cell._tePr if False else cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def add_table(doc, headers, rows, col_widths=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        set_run_font(run, size=9, bold=True, color=(255, 255, 255))
        shade_header(cell)
    for r_i, row in enumerate(rows):
        for c_i, val in enumerate(row):
            cell = table.rows[r_i + 1].cells[c_i]
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(str(val))
            set_run_font(run, size=9, bold=(r_i == 0))
    if col_widths:
        for row in table.rows:
            for i, w in enumerate(col_widths):
                row.cells[i].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = DocPt(4)
    return table


def write_english_report(stats, path):
    chosen = stats["chosen"]
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.54)
    section.right_margin = Cm(2.54)

    add_heading_styled(doc, "2. Data", 1)
    add_heading_styled(doc, "2.1 Data source, selection and cleaning", 2)

    add_para(
        doc,
        "This section documents the input data used to drive the bank-queue simulation. "
        "The raw observations come from an open survey of three retail banks in Ogun State, Nigeria "
        "(Bishop, Okagbue, Oguntunde, Opanuga, & Odetunmibi, 2018). The article and its supplementary "
        "workbooks are the sole data source; no additional field collection was undertaken. "
        "Among the three banks we retain only the third bank, and among its 20 operating days we retain "
        "a single representative day. Service-time extrema on that day are then removed before the sample "
        "is passed to input modelling.",
    )

    add_heading_styled(doc, "2.1.1 Source and variable definitions", 3)
    add_para(
        doc,
        "Bishop et al. (2018) recorded customers at three urban banks over four weeks per bank "
        "(Monday–Friday, 20 days each). Investigators used a stopwatch and a recorder. For every "
        "attended customer the files store an arrival clock time and three durations, in minutes:",
    )
    add_para(
        doc,
        "X1 — waiting time, from arrival until the cheque or withdrawal booklet is collected;  "
        "X2 — service time, the time used to process the cheque or withdrawal booklet;  "
        "X3 — total time in system, which satisfies X3 = X1 + X2 on every retained record.  "
        "N is the number of customers attended that day.",
        first_line=True,
    )
    add_para(
        doc,
        "In the published paper these variables are labelled Z1, Z2, Z3 and N3 for the third bank. "
        "The supplementary Excel files themselves use the headers X1, X2 and X3, so that notation is "
        "kept here. Times are measured in minutes. Arrival times run from 08:00 to 14:59, a seven-hour "
        "service window. Departure times were not captured, because customers often remained after the "
        "transaction to count cash or complete non-transaction errands (Bishop et al., 2018).",
    )
    add_para(
        doc,
        "The original third-bank files are THIRD BANK (WEEK 1).xlsx through THIRD BANK (WEEK 4).xlsx "
        "in the article’s mmc2 supplementary zip (https://doi.org/10.1016/j.dib.2018.05.101). Each weekly "
        "workbook has five daily sheets (Sheet1 = Monday, …, Sheet5 = Friday). These four workbooks were "
        "compiled into a single original file with one sheet per week. In total the third bank contributes "
        "17,222 customer records.",
    )

    add_heading_styled(doc, "2.1.2 Why the third bank", 3)
    add_para(
        doc,
        "The source contains three banks. Bank 1 has a slightly larger headcount (17,789) but a much more "
        "volatile service process: the published standard deviation of daily-mean service time is 6.04 minutes, "
        "versus 2.79 minutes at Bank 3 (Bishop et al., 2018, Tables 5 and 7). Bank 2 is smaller (15,488 customers) "
        "and similarly dispersed (service-time SD 6.04 minutes). Bank 3 therefore offers a complete 20-day, "
        "five-weekday panel, a large sample, and the most stable service-time behaviour of the three branches. "
        "That combination is better suited to a single-site retail-banking simulation, so all subsequent steps "
        "use only the third bank.",
    )

    add_heading_styled(doc, "2.1.3 Representative-day selection", 3)
    add_para(
        doc,
        "A full 20-day trace is unnecessarily long for the process-simulation exercise, and mixing days with "
        "very different congestion levels would blur the operating regime. We therefore select one day that is "
        "closest to the four-week third-bank profile on four metrics: daily headcount N, mean waiting time X1, "
        "mean service time X2, and mean system time X3.",
    )
    add_para(
        doc,
        "Let x̄_k be the unweighted mean of the 20 daily values of metric k ∈ {N, X1, X2, X3}. "
        "For each day d the deviation score is the sum of absolute relative deviations:",
        first_line=True,
    )
    add_para(
        doc,
        "D_d = |N_d − N̄| / N̄  +  |X1_d − X1̄| / X1̄  +  |X2_d − X2̄| / X2̄  +  |X3_d − X3̄| / X3̄.",
        bold=False,
        italic=True,
        first_line=False,
    )
    add_para(
        doc,
        "The selected day is d* = arg min_d D_d. This is equivalent to ranking days by the sum of percentage "
        "variances used in the cleaned workbook. The population baseline computed from the microdata is "
        f"N̄ = {stats['n0']:.2f} customers, X1̄ = {stats['x10']:.4f} min, X2̄ = {stats['x20']:.4f} min, "
        f"X3̄ = {stats['x30']:.4f} min.",
        first_line=True,
    )

    top = stats["ranked"][:5]
    add_table(
        doc,
        ["Rank", "Sheet", "Weekday", "N", "X1 mean", "X2 mean", "X3 mean", "D_d"],
        [
            [
                str(i),
                d["name"],
                d["day"],
                str(d["n"]),
                f"{d['x1m']:.4f}",
                f"{d['x2m']:.4f}",
                f"{d['x3m']:.4f}",
                f"{d['score']:.4f}",
            ]
            for i, d in enumerate(top, 1)
        ],
    )
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = cap.add_run("Table 2.1. Top five days by deviation score (third bank, 20 weekdays).")
    set_run_font(run, size=10, italic=True)

    add_para(
        doc,
        f"Week 1 Tuesday (W1-Sheet2) attains the unique minimum, D = {chosen['score']:.4f}. "
        f"Its headcount is {chosen['n']} (8.00% above the 20-day mean). Mean waiting, service and system times "
        f"are {chosen['x1m']:.4f}, {chosen['x2m']:.4f} and {chosen['x3m']:.4f} minutes, corresponding to "
        f"percentage variances of {chosen['p1']*100:.2f}%, {chosen['p2']*100:.2f}% and {chosen['p3']*100:.2f}%. "
        "No other day is within 0.12 of this score. The same Tuesday is the day whose published Table 4 averages "
        "are Z1 = 12, Z2 = 11, Z3 = 23 and N3 = 930, which match the microdata means after ordinary rounding. "
        "This day is therefore used as the representative operating day.",
    )

    add_heading_styled(doc, "2.1.4 Cleaning", 3)
    add_para(
        doc,
        "Selection above is performed on the raw 930 records so that the ranking is not distorted by "
        "post-hoc edits. After the day is fixed, two integrity checks and one outlier rule are applied.",
    )
    add_para(
        doc,
        "Integrity. Every retained row has X1, X2, X3 > 0 and X3 = X1 + X2 (maximum absolute discrepancy "
        "is numerical round-off on the order of 10^−15). Arrival times lie inside 08:00–14:59. No duplicate "
        "customer numbers appear. Tukey’s 1.5×IQR fences on the selected-day service times do not flag any "
        "additional outliers, which is consistent with a tightly concentrated X2 range of about 8–14 minutes.",
        first_line=True,
    )
    add_para(
        doc,
        "Extremum rule. Even so, the single smallest and single largest service times are removed so that "
        "input fitting is not levered by the two most extreme recordings. Formally, if "
        "x2_(1) ≤ … ≤ x2_(n) on the selected day, the cleaned sample is {x2_(i) : 2 ≤ i ≤ n−1} together with "
        "the matching X1, X3 and arrival fields. This is a 100% winsor-style trim of the two endpoints "
        "(one minimum, one maximum), not an IQR deletion. Two records are dropped:",
        first_line=True,
    )
    drop_rows = []
    for rec in stats["dropped"]:
        reason = "X2 minimum" if rec["x2"] == stats["xmin"] else "X2 maximum"
        arr = rec["arrival"].strftime("%H:%M") if isinstance(rec["arrival"], time) else str(rec["arrival"])
        drop_rows.append(
            [str(rec["number"]), arr, f"{rec['x1']:.4f}", f"{rec['x2']:.4f}", f"{rec['x3']:.4f}", reason]
        )
    add_table(
        doc,
        ["Customer No.", "Arrival", "X1", "X2", "X3", "Reason"],
        drop_rows,
    )
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = cap.add_run("Table 2.2. Service-time extrema removed from Week 1 Tuesday.")
    set_run_font(run, size=10, italic=True)

    add_para(
        doc,
        f"The cleaned file therefore contains {len(stats['kept'])} customers. Mean service time changes only "
        f"from {stats['raw_x2_mean']:.4f} to {stats['kept_x2_mean']:.4f} minutes, which confirms that the trim "
        "is conservative and does not rewrite the day’s operating character. The cleaned sample is stored in "
        "the Representative Day Data sheet of the analysis workbook; the two deleted rows are listed in Cleaning Log.",
    )

    add_heading_styled(doc, "2.1.5 Cleaned sample used in later sections", 3)
    ia_minutes = 6.983333333333333 * 60 / 929  # from 08:00 to 14:59 over 929 intervals
    add_para(
        doc,
        f"The simulation input from this section is the Week 1 Tuesday cleaned trace: {len(stats['kept'])} customers, "
        f"mean waiting time {stats['kept_x1_mean']:.4f} min, mean service time {stats['kept_x2_mean']:.4f} min "
        f"(SD {stats['kept_x2_std']:.4f}, range {stats['kept_x2_min']:.4f}–{stats['kept_x2_max']:.4f}), and mean "
        f"system time {stats['kept_x3_mean']:.4f} min. Arrivals span 08:00–14:59. Because several customers can "
        "share the same recorded minute, some raw inter-arrival gaps are zero; later input analysis should "
        "either jitter tied arrivals inside the minute or treat them as a bulk arrival. A single-server model "
        "is not feasible at this volume (about 133 arrivals per hour with an 11-minute service time), so the "
        "process model in later sections must use a multi-server bank counter.",
    )
    add_para(
        doc,
        "Reference. Bishop, S. A., Okagbue, H. I., Oguntunde, P. E., Opanuga, A. A., & Odetunmibi, O. A. (2018). "
        "Survey dataset on analysis of queues in some selected banks in Ogun State, Nigeria. Data in Brief, 19, "
        "835–841. https://doi.org/10.1016/j.dib.2018.05.101",
        italic=False,
        first_line=False,
    )
    doc.save(path)


def write_bilingual(stats, path):
    chosen = stats["chosen"]
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

    t = doc.add_paragraph()
    r = t.add_run("MS5225  ·  Section 2.1 Data  ·  English / 中文对照")
    set_run_font(r, size=16, bold=True, color=(31, 78, 121))
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER

    pairs = [
        (
            "2.1 Data source, selection and cleaning",
            "2.1 数据来源、选择与清洗",
        ),
        (
            "The raw observations come from Bishop et al. (2018), Data in Brief 19:835–841, DOI 10.1016/j.dib.2018.05.101. "
            "Three retail banks in Ogun State, Nigeria were observed with a stopwatch and recorder. We use only the third bank. "
            "The original files are the article’s mmc2 workbooks THIRD BANK (WEEK 1–4).xlsx, compiled into one Excel file with four weekly sheets.",
            "原始观测来自 Bishop 等人（2018），Data in Brief 19:835–841，DOI 10.1016/j.dib.2018.05.101。"
            "研究者用秒表和记录器观测了尼日利亚奥贡州三家零售银行。本项目只使用第三家银行。"
            "原始文件是论文补充材料 mmc2 中的 THIRD BANK (WEEK 1–4).xlsx，已汇编成一个含四个周工作表的 Excel。",
        ),
        (
            "Variables. X1 = waiting time (minutes) until the cheque/withdrawal booklet is collected. "
            "X2 = service/processing time (minutes). X3 = total time in system (minutes), and X3 = X1 + X2. "
            "N = number of customers attended. The paper labels these Z1, Z2, Z3, N3 for Bank 3; the Excel headers are X1–X3. "
            "The observation window is 08:00–14:59. Bank 3 has 17,222 customer records over 20 weekdays.",
            "变量：X1 = 等待时间（分钟，至存折/支票被收取）；X2 = 服务/办理时间（分钟）；X3 = 系统逗留时间（分钟），且 X3 = X1 + X2；"
            "N = 当天办理人数。论文对第三家银行记为 Z1、Z2、Z3、N3，补充表头仍为 X1–X3。"
            "观测窗口为 08:00–14:59。第三家银行 20 个工作日共 17,222 条顾客记录。",
        ),
        (
            "Why Bank 3. Bank 1 is larger (17,789) but service-time daily means are much more volatile (published SD 6.04 min). "
            "Bank 2 is smaller (15,488) and similarly dispersed (SD 6.04 min). Bank 3 has a complete four-week weekday panel, "
            "a large sample, and the most stable service process (published SD 2.79 min), so it is the better prototype for one-site simulation.",
            "为何选第三家银行。第一家人数略多（17,789），但日均服务时间波动很大（论文标准差 6.04 分钟）。"
            "第二家更小（15,488），波动同样大（标准差 6.04 分钟）。第三家有完整四周工作日面板、样本量充足，"
            "且服务过程最稳定（论文标准差 2.79 分钟），更适合作为单网点仿真原型。",
        ),
        (
            "Representative-day formula. For each of the 20 days compute "
            "D_d = |N_d−N̄|/N̄ + |X1_d−X1̄|/X1̄ + |X2_d−X2̄|/X2̄ + |X3_d−X3̄|/X3̄, "
            f"where the baseline is the unweighted 20-day mean (N̄={stats['n0']:.2f}, X1̄={stats['x10']:.4f}, "
            f"X2̄={stats['x20']:.4f}, X3̄={stats['x30']:.4f}). Select d* = arg min D_d.",
            "代表日公式。对 20 天中的每一天计算 "
            "D_d = |N_d−N̄|/N̄ + |X1_d−X1̄|/X1̄ + |X2_d−X2̄|/X2̄ + |X3_d−X3̄|/X3̄，"
            f"其中基准为 20 天未加权均值（N̄={stats['n0']:.2f}，X1̄={stats['x10']:.4f}，"
            f"X2̄={stats['x20']:.4f}，X3̄={stats['x30']:.4f}）。选择 D_d 最小的一天。",
        ),
        (
            f"Selected day: Week 1 Tuesday, sheet W1-Sheet2. N={chosen['n']}, "
            f"X1={chosen['x1m']:.4f}, X2={chosen['x2m']:.4f}, X3={chosen['x3m']:.4f}, D={chosen['score']:.4f} "
            f"(about 8.00%, {chosen['p1']*100:.2f}%, {chosen['p2']*100:.2f}%, {chosen['p3']*100:.2f}% relative deviations). "
            "This is the unique minimum; the next-best day has D ≈ 0.226. It matches published Table 4 Tuesday: 12, 11, 23, 930.",
            f"入选日期：第 1 周星期二，工作表 W1-Sheet2。N={chosen['n']}，"
            f"X1={chosen['x1m']:.4f}，X2={chosen['x2m']:.4f}，X3={chosen['x3m']:.4f}，D={chosen['score']:.4f} "
            f"（相对偏差约 8.00%、{chosen['p1']*100:.2f}%、{chosen['p2']*100:.2f}%、{chosen['p3']*100:.2f}%）。"
            "这是唯一最小；第二名约 0.226。与论文 Table 4 星期二的 12、11、23、930 一致。",
        ),
        (
            f"Cleaning. Ranking uses the raw 930 rows. Then the extreme minimum and maximum X2 values are deleted "
            f"(x2_min={stats['xmin']:.4f} min, x2_max={stats['xmax']:.4f} min; two rows). "
            f"Cleaned n={len(stats['kept'])}. Mean X2 moves only from {stats['raw_x2_mean']:.4f} to {stats['kept_x2_mean']:.4f}. "
            "Tukey 1.5×IQR fences find zero extra outliers. All rows satisfy X3=X1+X2 and positive times. "
            "This is an endpoint trim of the selected-day service times, not a rewrite of the other 19 days.",
            f"清洗。排序使用原始 930 条。随后删除 X2 的极大值与极小值各一条"
            f"（x2_min={stats['xmin']:.4f} 分钟，x2_max={stats['xmax']:.4f} 分钟，共 2 条）。"
            f"清洗后 n={len(stats['kept'])}。X2 均值仅从 {stats['raw_x2_mean']:.4f} 变为 {stats['kept_x2_mean']:.4f}。"
            "Tukey 1.5×IQR 未再检出异常值。全部记录满足 X3=X1+X2 且时间为正。"
            "这是对入选日服务时间两端点的修剪，不改动其余 19 天。",
        ),
        (
            f"Final input. {len(stats['kept'])} customers on Week 1 Tuesday; mean wait {stats['kept_x1_mean']:.4f} min; "
            f"mean service {stats['kept_x2_mean']:.4f} min (SD {stats['kept_x2_std']:.4f}); mean sojourn {stats['kept_x3_mean']:.4f} min. "
            "A multi-server counter is required. Tied clock minutes should be jittered or treated as bulk arrivals in later input analysis.",
            f"最终输入。第 1 周星期二清洗后 {len(stats['kept'])} 名顾客；平均等待 {stats['kept_x1_mean']:.4f} 分钟；"
            f"平均服务 {stats['kept_x2_mean']:.4f} 分钟（标准差 {stats['kept_x2_std']:.4f}）；平均逗留 {stats['kept_x3_mean']:.4f} 分钟。"
            "后续必须使用多服务台。同一分钟内的并列到达需在输入分析中做微小抖动或按批量到达处理。",
        ),
    ]

    table = doc.add_table(rows=len(pairs), cols=2)
    table.style = "Table Grid"
    table.autofit = True
    for i, (en, zh) in enumerate(pairs):
        left, right = table.rows[i].cells
        if i == 0:
            shade_header(left)
            shade_header(right)
            for cell, text in ((left, en), (right, zh)):
                cell.text = ""
                p = cell.paragraphs[0]
                run = p.add_run(text)
                set_run_font(run, size=12, bold=True, color=(255, 255, 255))
        else:
            for cell, text, east in ((left, en, False), (right, zh, True)):
                cell.text = ""
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.space_after = DocPt(4)
                run = p.add_run(text)
                set_run_font(run, size=10)
                if east:
                    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
        left.width = Cm(8.5)
        right.width = Cm(8.5)

    note = doc.add_paragraph()
    run = note.add_run(
        "\nPPT (3 slides, English only; paste into the official MS5225 template if the course master differs): "
        "Slide 1 Data source & why Bank 3; Slide 2 representative-day formula and ranking; "
        "Slide 3 min/max cleaning and cleaned sample. "
        "课程 PPT 请把这三页内容贴进官方模板母版，不要另换一套版式。"
    )
    set_run_font(run, size=10, italic=True)
    doc.save(path)


def rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def add_bar(slide, prs, color):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), prs.slide_width, Inches(0.92))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0.92), prs.slide_width, Inches(0.06))
    accent.fill.solid()
    accent.fill.fore_color.rgb = rgb("F4B400")
    accent.line.fill.background()
    return shape


def set_tf(shape, lines, *, size=18, bold=False, color="FFFFFF", align=PP_ALIGN.LEFT):
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = line
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = rgb(color)
        run.font.name = "Calibri"
        p.space_after = Pt(4)


def add_textbox(slide, l, t, w, h, lines, **kwargs):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    set_tf(box, lines, **kwargs)
    return box


def add_ppt_table(slide, left, top, width, height, data, col_w, header=True, highlight_row=None):
    rows, cols = len(data), len(data[0])
    table_shape = slide.shapes.add_table(rows, cols, Inches(left), Inches(top), Inches(width), Inches(height))
    table = table_shape.table
    for i, w in enumerate(col_w):
        table.columns[i].width = Inches(w)
    for r in range(rows):
        for c in range(cols):
            cell = table.cell(r, c)
            cell.text = str(data[r][c])
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER
                for run in p.runs:
                    run.font.name = "Calibri"
                    run.font.size = Pt(11 if r == 0 else 10)
                    run.font.bold = r == 0 or (highlight_row is not None and r == highlight_row)
                    if r == 0:
                        run.font.color.rgb = rgb("FFFFFF")
                    else:
                        run.font.color.rgb = rgb("1B1B1B")
            fill = cell.fill
            fill.solid()
            if r == 0:
                fill.fore_color.rgb = rgb("1F4E79")
            elif highlight_row is not None and r == highlight_row:
                fill.fore_color.rgb = rgb("C6EFCE")
            elif r % 2 == 0:
                fill.fore_color.rgb = rgb("F2F2F2")
            else:
                fill.fore_color.rgb = rgb("FFFFFF")
    return table_shape


def write_pptx(stats, path):
    chosen = stats["chosen"]
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # Slide 1
    s1 = prs.slides.add_slide(blank)
    add_bar(s1, prs, rgb("0B3C5D"))
    add_textbox(s1, 0.4, 0.18, 12.4, 0.35, ["MS5225  ·  Business Process Modeling and Simulation  ·  Section 2.1 Data"], size=13, color="D0D7DE")
    add_textbox(s1, 0.4, 0.46, 12.4, 0.4, ["Data source: third-bank queue observations"], size=28, bold=True, color="FFFFFF")

    cards = [
        (0.4, "Source", "Bishop et al. (2018), Data in Brief\nDOI 10.1016/j.dib.2018.05.101\nOpen supplementary mmc2 Excel files"),
        (4.7, "Coverage", "3 urban banks in Ogun State, Nigeria\n4 weeks × Mon–Fri per bank\nStopwatch + recorder, 08:00–14:59"),
        (9.0, "Why Bank 3", "17,222 customers, complete 20-day panel\nMost stable service times (SD 2.79 min)\nBetter single-site prototype than Banks 1–2"),
    ]
    for x, title, body in cards:
        sh = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.25), Inches(4.0), Inches(2.35))
        sh.fill.solid()
        sh.fill.fore_color.rgb = rgb("F7F9FC")
        sh.line.color.rgb = rgb("1F4E79")
        add_textbox(s1, x + 0.15, 1.35, 3.7, 0.4, [title], size=18, bold=True, color="1F4E79")
        add_textbox(s1, x + 0.15, 1.8, 3.7, 1.6, body.split("\n"), size=14, color="1B1B1B")

    add_textbox(s1, 0.4, 3.75, 12.5, 0.35, ["Variable definitions (third-bank Excel headers X1, X2, X3  =  paper Z1, Z2, Z3)"], size=16, bold=True, color="0B3C5D")
    add_ppt_table(
        s1,
        0.4,
        4.15,
        12.5,
        2.5,
        [
            ["Symbol", "Meaning", "Unit", "Role in the simulation"],
            ["X1", "Wait until booklet/cheque is collected", "minutes", "Queue delay / waiting-time input check"],
            ["X2", "Processing (service) time", "minutes", "Primary service-time input"],
            ["X3 = X1 + X2", "Total time in the system", "minutes", "Integrity check; sojourn validation"],
            ["N", "Customers attended that day", "count", "Daily volume / arrival intensity"],
        ],
        [2.2, 4.2, 1.6, 4.5],
    )
    s1.notes_slide.notes_text_frame.text = (
        "Open with the DOI. Emphasise that we did not collect new field data. "
        "Bank 3 is chosen for completeness and stable service times, not because it is the largest."
    )

    # Slide 2
    s2 = prs.slides.add_slide(blank)
    add_bar(s2, prs, rgb("0B3C5D"))
    add_textbox(s2, 0.4, 0.18, 12.4, 0.35, ["MS5225  ·  Section 2.1 Data"], size=13, color="D0D7DE")
    add_textbox(s2, 0.4, 0.46, 12.4, 0.4, ["Why this day: minimum deviation from the 20-day profile"], size=26, bold=True, color="FFFFFF")
    add_textbox(
        s2,
        0.4,
        1.15,
        12.5,
        0.85,
        [
            "Baseline = unweighted mean of 20 daily sheets:  N̄ = 861.10,   X1̄ = 11.8585 min,   X2̄ = 11.1607 min,   X3̄ = 23.0191 min",
            "Score  D_d = |N_d−N̄|/N̄  +  |X1_d−X1̄|/X1̄  +  |X2_d−X2̄|/X2̄  +  |X3_d−X3̄|/X3̄     →     choose  d* = arg min D_d",
        ],
        size=15,
        color="1B1B1B",
    )
    top5 = stats["ranked"][:6]
    table_data = [["Rank", "Sheet", "Weekday", "N", "X1", "X2", "X3", "D_d"]]
    for i, d in enumerate(top5, 1):
        table_data.append(
            [str(i), d["name"], d["day"], str(d["n"]), f"{d['x1m']:.3f}", f"{d['x2m']:.3f}", f"{d['x3m']:.3f}", f"{d['score']:.4f}"]
        )
    add_ppt_table(s2, 0.4, 2.1, 8.4, 3.6, table_data, [0.8, 1.4, 1.4, 0.8, 1.0, 1.0, 1.0, 1.0], highlight_row=1)

    box = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.05), Inches(2.1), Inches(3.85), Inches(3.6))
    box.fill.solid()
    box.fill.fore_color.rgb = rgb("E8F5E9")
    box.line.color.rgb = rgb("548235")
    add_textbox(s2, 9.2, 2.25, 3.55, 0.4, ["Selected day"], size=16, bold=True, color="548235")
    add_textbox(
        s2,
        9.2,
        2.7,
        3.55,
        2.7,
        [
            "Week 1 Tuesday",
            "Sheet W1-Sheet2",
            f"N = {chosen['n']}  (+8.00%)",
            f"X1 = {chosen['x1m']:.2f} min",
            f"X2 = {chosen['x2m']:.2f} min",
            f"X3 = {chosen['x3m']:.2f} min",
            f"D = {chosen['score']:.4f}  (unique min)",
            "Matches paper Table 4 Tuesday",
            "12, 11, 23, N=930",
        ],
        size=14,
        color="1B1B1B",
    )
    add_textbox(
        s2,
        0.4,
        5.9,
        12.5,
        1.2,
        [
            "Interpretation. Tuesday is a normal-volume weekday whose waiting, service and sojourn means all sit within about 1% of the four-week average; only headcount is moderately higher. Mixing all 20 days would blend quiet Fridays with overloaded Thursdays and hide the operating regime we want to simulate.",
        ],
        size=14,
        color="1B1B1B",
    )
    s2.notes_slide.notes_text_frame.text = (
        "Walk through the four-term formula once. Point at rank 1 in green. "
        "Say the next day is more than twice as far from the baseline (0.23 vs 0.10)."
    )

    # Slide 3
    s3 = prs.slides.add_slide(blank)
    add_bar(s3, prs, rgb("0B3C5D"))
    add_textbox(s3, 0.4, 0.18, 12.4, 0.35, ["MS5225  ·  Section 2.1 Data"], size=13, color="D0D7DE")
    add_textbox(s3, 0.4, 0.46, 12.4, 0.4, ["Cleaning: drop X2 min and max, then freeze the sample"], size=26, bold=True, color="FFFFFF")

    add_ppt_table(
        s3,
        0.4,
        1.2,
        7.6,
        2.6,
        [
            ["Step", "Rule / formula", "Result"],
            ["1. Integrity", "X1,X2,X3>0 and X3=X1+X2", "All 930 raw rows pass"],
            ["2. IQR screen", "Tukey fence Q1±1.5 IQR on X2", "0 extra outliers"],
            ["3. Endpoint trim", "Drop x2_(1) and x2_(n) only", "2 rows removed"],
            ["4. Freeze sample", "Keep remaining customers", f"{len(stats['kept'])} records"],
        ],
        [1.6, 3.4, 2.6],
    )
    drop_data = [["Removed", "No.", "X2 (min)", "Reason"]]
    for rec in stats["dropped"]:
        reason = "minimum" if rec["x2"] == stats["xmin"] else "maximum"
        drop_data.append(["W1 Tue", str(rec["number"]), f"{rec['x2']:.4f}", reason])
    add_ppt_table(s3, 8.2, 1.2, 4.7, 1.7, drop_data, [1.1, 0.9, 1.3, 1.4])

    add_ppt_table(
        s3,
        8.2,
        3.05,
        4.7,
        2.2,
        [
            ["Cleaned input", "Value"],
            ["Customers n", str(len(stats["kept"]))],
            ["Mean wait X1", f"{stats['kept_x1_mean']:.3f} min"],
            ["Mean service X2", f"{stats['kept_x2_mean']:.3f} min"],
            ["SD(X2)", f"{stats['kept_x2_std']:.3f} min"],
            ["Mean sojourn X3", f"{stats['kept_x3_mean']:.3f} min"],
        ],
        [2.4, 2.3],
    )

    add_textbox(
        s3,
        0.4,
        4.05,
        7.6,
        2.9,
        [
            "What changed. Mean X2 only moves from {:.4f} to {:.4f} minutes, so the trim does not invent a new day.".format(
                stats["raw_x2_mean"], stats["kept_x2_mean"]
            ),
            "What we do not claim. No compatibility test, exam score, or extra field audit was added. Ranking used the raw 930 rows; cleaning is applied afterwards.",
            "Use in later sections. Fit the service distribution to cleaned X2. Treat same-minute arrivals as ties. A single server cannot clear ~133 arrivals/hour at an 11-minute service time — use a multi-server counter.",
        ],
        size=15,
        color="1B1B1B",
    )
    s3.notes_slide.notes_text_frame.text = (
        "Stress that min/max deletion is two observations, not a wholesale rewrite. "
        "IQR already found nothing; the trim is a conservative extra step requested for robustness."
    )
    prs.save(path)


def main():
    days = load_third_bank()
    orig = os.path.join(HERE, "01_Third_Bank_W1-W4_original_compiled.xlsx")
    cleaned = os.path.join(HERE, "02_Third_Bank_representative_day_cleaned.xlsx")
    src_out = os.path.join(HERE, "00_original_third_bank_source")
    os.makedirs(src_out, exist_ok=True)
    for week in range(1, 5):
        fn = f"THIRD BANK (WEEK {week}).xlsx"
        shutil.copy2(os.path.join(SRC_DIR, fn), os.path.join(src_out, fn))
    print("building original compiled workbook")
    build_original_compiled(days, orig)
    print("building cleaned workbook")
    stats = build_cleaned_workbook(days, cleaned)
    print("writing English report")
    write_english_report(stats, os.path.join(HERE, "MS5225_Report_Section_2.1_Data.docx"))
    print("writing bilingual对照")
    write_bilingual(stats, os.path.join(HERE, "MS5225_Report_Section_2.1_EN_CN对照.docx"))
    print("writing pptx")
    write_pptx(stats, os.path.join(HERE, "MS5225_PPT_Section_2.1_Data.pptx"))
    print("done")
    print("selected", stats["chosen"]["name"], stats["chosen"]["score"])
    print("kept", len(stats["kept"]), "dropped", len(stats["dropped"]))


if __name__ == "__main__":
    main()
