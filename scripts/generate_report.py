# generate_report.py
# Generates a Word (.docx) LO-to-PO Alignment Report from alignment.json
# and program_outcomes.json produced by the pipeline.
#
# Usage: python scripts/generate_report.py --program ITSM

import argparse
import io
import json
import os
from datetime import date

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── Argument parsing ──────────────────────────────────────────────────────────
parser = argparse.ArgumentParser(
    description="Generate a Word report from alignment.json and program_outcomes.json."
)
parser.add_argument("--program", required=True,
                    help="Program folder name under data/processed/ (e.g. ITSM)")
args = parser.parse_args()

_BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_PROC_DIR = os.path.join(_BASE_DIR, "data", "processed", args.program)
OUTPUT    = os.path.join(_PROC_DIR, "LO_PO_Alignment_Report.docx")

# ── Load data ─────────────────────────────────────────────────────────────────
with open(os.path.join(_PROC_DIR, "program_outcomes.json"), encoding="utf-8") as f:
    po_data = json.load(f)   # [{"id": "PO1", "text": "..."}, ...]

with open(os.path.join(_PROC_DIR, "alignment.json"), encoding="utf-8") as f:
    alignment_data = json.load(f)  # [{course_code, course_name, outcomes: [{outcome_text, primary, supporting}]}]

# ── Derive display structures ─────────────────────────────────────────────────
N_PO     = len(po_data)
po_ids   = [po["id"]   for po in po_data]   # ["PO1", "PO2", ...]
po_texts = [po["text"] for po in po_data]   # full text strings

def po_num(po_id: str) -> int:
    """'PO3' -> 3"""
    return int(po_id.replace("PO", ""))

# COURSES ordered list
COURSES = [c["course_code"] for c in alignment_data]

# RAW heat map data: course_code -> {po_number (int): "P" | "S"}
# Aggregate across all LOs — primary wins over supporting
RAW = {}
for course in alignment_data:
    code = course["course_code"]
    RAW[code] = {}
    for outcome in course.get("outcomes", []):
        for pid in outcome.get("primary", []):
            RAW[code][po_num(pid)] = "P"
        for pid in outcome.get("supporting", []):
            if po_num(pid) not in RAW[code]:   # don't downgrade P to S
                RAW[code][po_num(pid)] = "S"

# LO breakdown table rows: (code, name, lo_text, "PO1, PO2, ...")
LO_DATA = []
for course in alignment_data:
    code = course["course_code"]
    name = course["course_name"]
    for outcome in course.get("outcomes", []):
        all_pos    = outcome.get("primary", []) + outcome.get("supporting", [])
        all_pos_s  = sorted(set(all_pos), key=po_num)
        mapped_str = ", ".join(all_pos_s) if all_pos_s else "—"
        LO_DATA.append((code, name, outcome["outcome_text"], mapped_str))

# PO coverage counts (courses that touch each PO at all)
po_labels = po_ids
po_counts = []
for pid in po_ids:
    count = sum(
        1 for c in alignment_data
        if any(pid in o.get("primary", []) + o.get("supporting", [])
               for o in c.get("outcomes", []))
    )
    po_counts.append(count)

# Auto-generate gap analysis
GAPS = []
total_primaries = sum(
    len(o.get("primary", []))
    for c in alignment_data
    for o in c.get("outcomes", [])
)

for i, pid in enumerate(po_ids):
    total = po_counts[i]
    short = po_texts[i][:55] + ("..." if len(po_texts[i]) > 55 else "")
    label = f"{pid} — {short}"

    touching_courses = [
        c["course_code"] for c in alignment_data
        if any(pid in o.get("primary", []) + o.get("supporting", [])
               for o in c.get("outcomes", []))
    ]

    if total == 0:
        GAPS.append((
            label,
            "No course learning outcomes align to this Program Outcome.",
            "Review whether this PO requires dedicated course-level coverage, "
            "or document it as a program-level milestone met outside the course curriculum."
        ))
    elif total <= 2:
        GAPS.append((
            label,
            f"Only {total} course(s) support this outcome ({', '.join(touching_courses)}). "
            "Coverage is thin across the program.",
            "Consider adding an explicit reference to this outcome in 2–3 additional "
            "course learning objectives."
        ))

# Flag dominant PO (>35% of all primary mappings)
if total_primaries > 0:
    for i, pid in enumerate(po_ids):
        primary_count = sum(
            1 for c in alignment_data
            for o in c.get("outcomes", [])
            if pid in o.get("primary", [])
        )
        if primary_count / total_primaries > 0.35:
            GAPS.append((
                f"{pid} — Dominant Coverage",
                f"{primary_count} course LOs map primarily to {pid} — "
                "the highest of any single Program Outcome.",
                f"Consider whether {pid} should be split into more specific sub-outcomes "
                "to better distinguish areas of curriculum focus."
            ))

PROGRAM_NAME = args.program
REPORT_DATE  = date.today().strftime("%B %Y")

# ── Colour palette ────────────────────────────────────────────────────────────
NAVY_HEX       = "1F497D"
MED_BLUE_HEX   = "276FB4"
LIGHT_BLUE_HEX = "BDD7EE"
PALE_GREY_HEX  = "F2F2F2"
WHITE_HEX      = "FFFFFF"
DARK_GREY_HEX  = "404040"

NAVY       = RGBColor(0x1F, 0x49, 0x7D)
MED_BLUE   = RGBColor(0x27, 0x6F, 0xB4)
LIGHT_BLUE = RGBColor(0xBD, 0xD7, 0xEE)
PALE_GREY  = RGBColor(0xF2, 0xF2, 0xF2)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
DARK_GREY  = RGBColor(0x40, 0x40, 0x40)

# ── Helpers ───────────────────────────────────────────────────────────────────
def set_cell_bg(cell, hex_color: str):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    existing = tcPr.find(qn("w:shd"))
    if existing is not None:
        tcPr.remove(existing)
    tcPr.append(shd)

def cell_para(cell, text, bold=False, size=9, color=None, align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ""
    p    = cell.paragraphs[0]
    p.alignment = align
    run  = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = color
    return p

def add_heading(doc, text, level=1):
    p = doc.add_heading(text, level=level)
    for run in p.runs:
        run.font.color.rgb = NAVY
    return p

def no_space_para(doc, text="", size=11, bold=False, color=None, align=WD_ALIGN_PARAGRAPH.LEFT):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(4)
    p.alignment = align
    if text:
        run = p.add_run(text)
        run.font.size = Pt(size)
        run.bold = bold
        if color:
            run.font.color.rgb = color
    return p

def add_page_break(doc):
    doc.add_page_break()

def set_repeat_header(row):
    tr   = row._tr
    trPr = tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:tblHeader"))

# ── Document setup ────────────────────────────────────────────────────────────
doc = Document()

style = doc.styles["Normal"]
style.font.name = "Calibri"
style.font.size = Pt(11)

for lvl, sz, sp_b, sp_a in [(1, 16, 12, 6), (2, 13, 10, 4), (3, 11, 8, 2)]:
    h = doc.styles[f"Heading {lvl}"]
    h.font.name  = "Calibri"
    h.font.size  = Pt(sz)
    h.font.color.rgb = NAVY
    h.font.bold  = True
    h.paragraph_format.space_before = Pt(sp_b)
    h.paragraph_format.space_after  = Pt(sp_a)

section = doc.sections[0]
section.page_width  = Inches(8.5)
section.page_height = Inches(11)
for attr in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
    setattr(section, attr, Inches(1))

def add_header_footer(sec, show_header=True):
    if show_header:
        hdr = sec.header
        hdr.is_linked_to_previous = False
        hp  = hdr.paragraphs[0]
        hp.clear()
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = hp.add_run(f"NSCC — {PROGRAM_NAME} Program  |  LO–PO Alignment Report  |  {REPORT_DATE}")
        run.font.size = Pt(9)
        run.font.color.rgb = DARK_GREY
        pPr  = hp._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bot  = OxmlElement("w:bottom")
        bot.set(qn("w:val"),   "single")
        bot.set(qn("w:sz"),    "4")
        bot.set(qn("w:space"), "1")
        bot.set(qn("w:color"), NAVY_HEX)
        pBdr.append(bot)
        pPr.append(pBdr)

    ftr = sec.footer
    ftr.is_linked_to_previous = False
    fp  = ftr.paragraphs[0]
    fp.clear()
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = fp.add_run("Page ")
    run.font.size = Pt(9)
    run.font.color.rgb = DARK_GREY
    fldChar1     = OxmlElement("w:fldChar");  fldChar1.set(qn("w:fldCharType"), "begin")
    instrText    = OxmlElement("w:instrText"); instrText.text = "PAGE"
    fldChar2     = OxmlElement("w:fldChar");  fldChar2.set(qn("w:fldCharType"), "end")
    run2 = fp.add_run()
    run2.font.size = Pt(9)
    run2.font.color.rgb = DARK_GREY
    run2._r.append(fldChar1)
    run2._r.append(instrText)
    run2._r.append(fldChar2)

add_header_footer(section, show_header=False)   # title page: footer only

# ════════════════════════════════════════════════════════════════════════════
# SECTION 1 – TITLE PAGE
# ════════════════════════════════════════════════════════════════════════════
for _ in range(6):
    doc.add_paragraph()

title_p = doc.add_paragraph()
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
tr = title_p.add_run("Course Learning Outcomes to\nProgram Outcomes Alignment Report")
tr.font.size = Pt(26); tr.font.bold = True; tr.font.color.rgb = NAVY

doc.add_paragraph()

sub_p = doc.add_paragraph()
sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
sr = sub_p.add_run(f"{PROGRAM_NAME} — NSCC")
sr.font.size = Pt(14); sr.font.color.rgb = DARK_GREY

doc.add_paragraph()

ay_p = doc.add_paragraph()
ay_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
ay_p.add_run(f"Generated: {REPORT_DATE}").font.size = Pt(12)

div = doc.add_paragraph()
div.alignment = WD_ALIGN_PARAGRAPH.CENTER
pPr  = div._p.get_or_add_pPr()
pBdr = OxmlElement("w:pBdr")
top  = OxmlElement("w:top")
top.set(qn("w:val"), "single"); top.set(qn("w:sz"), "12")
top.set(qn("w:space"), "1");    top.set(qn("w:color"), NAVY_HEX)
pBdr.append(top); pPr.append(pBdr)

add_page_break(doc)

from docx.enum.section import WD_SECTION
new_sec = doc.add_section(WD_SECTION.NEW_PAGE)
new_sec.page_width  = Inches(8.5)
new_sec.page_height = Inches(11)
for attr in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
    setattr(new_sec, attr, Inches(1))
add_header_footer(new_sec, show_header=True)

# ════════════════════════════════════════════════════════════════════════════
# SECTION 2 – EXECUTIVE SUMMARY
# ════════════════════════════════════════════════════════════════════════════
add_heading(doc, "Executive Summary", 1)
no_space_para(doc, (
    f"This report maps all course-level learning outcomes (LOs) from required courses in the "
    f"{PROGRAM_NAME} diploma program to the {N_PO} Program Outcomes (POs). "
    "The mapping was conducted to assess the breadth and depth of coverage each Program Outcome "
    "receives across the curriculum, identify outcomes with insufficient course-level support, "
    "and provide evidence-based recommendations to inform potential revisions to the Program "
    "Outcome list. The report includes a coverage heat map, a bar chart summarising coverage "
    "counts, a full per-course breakdown table, and an auto-generated gap analysis."
), size=11)

# ════════════════════════════════════════════════════════════════════════════
# SECTION 3 – PROGRAM OUTCOMES REFERENCE
# ════════════════════════════════════════════════════════════════════════════
add_heading(doc, "Program Outcomes Reference", 1)
for po in po_data:
    p = doc.add_paragraph(style="List Number")
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(2)
    run = p.add_run(f"{po['id']}: {po['text']}")
    run.font.size = Pt(10)

# ════════════════════════════════════════════════════════════════════════════
# SECTION 4 – HEAT MAP (landscape section)
# ════════════════════════════════════════════════════════════════════════════
add_page_break(doc)
land_sec = doc.add_section(WD_SECTION.NEW_PAGE)
land_sec.orientation   = 1
land_sec.page_width    = Inches(11)
land_sec.page_height   = Inches(8.5)
land_sec.left_margin   = Inches(0.5)
land_sec.right_margin  = Inches(0.5)
land_sec.top_margin    = Inches(0.75)
land_sec.bottom_margin = Inches(0.75)
add_header_footer(land_sec, show_header=True)

add_heading(doc, "Coverage Heat Map", 1)
no_space_para(doc,
    "Each cell indicates whether a course provides primary (●) or supporting (○) alignment to a Program Outcome.",
    size=10)

COURSE_COL = 1.0
PO_COL     = (10.0 - COURSE_COL) / N_PO
col_widths  = [COURSE_COL] + [PO_COL] * N_PO

tbl = doc.add_table(rows=1, cols=N_PO + 1)
tbl.style = "Table Grid"

hdr_row = tbl.rows[0]
set_repeat_header(hdr_row)
set_cell_bg(hdr_row.cells[0], NAVY_HEX)
cell_para(hdr_row.cells[0], "Course", bold=True, size=8, color=WHITE)
for i, pid in enumerate(po_ids):
    set_cell_bg(hdr_row.cells[i + 1], NAVY_HEX)
    cell_para(hdr_row.cells[i + 1], pid, bold=True, size=8,
              color=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)

for r_idx, course in enumerate(COURSES):
    row    = tbl.add_row()
    bg_hex = PALE_GREY_HEX if r_idx % 2 == 0 else WHITE_HEX
    set_cell_bg(row.cells[0], PALE_GREY_HEX)
    cell_para(row.cells[0], course, bold=True, size=8)
    po_map = RAW.get(course, {})
    for i, pid in enumerate(po_ids):
        val = po_map.get(po_num(pid), "")
        ci  = i + 1
        if val == "P":
            set_cell_bg(row.cells[ci], MED_BLUE_HEX)
            cell_para(row.cells[ci], "●", bold=True, size=9,
                      color=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)
        elif val == "S":
            set_cell_bg(row.cells[ci], LIGHT_BLUE_HEX)
            cell_para(row.cells[ci], "○", size=9,
                      color=NAVY, align=WD_ALIGN_PARAGRAPH.CENTER)
        else:
            set_cell_bg(row.cells[ci], bg_hex)
            cell_para(row.cells[ci], "", size=8)

for row in tbl.rows:
    for c_idx, w in enumerate(col_widths):
        row.cells[c_idx].width = Inches(w)

leg = doc.add_paragraph()
leg.paragraph_format.space_before = Pt(6)
leg.paragraph_format.space_after  = Pt(0)
r1 = leg.add_run("  ●  Primary alignment    ")
r1.font.size = Pt(9); r1.font.bold = True; r1.font.color.rgb = MED_BLUE
r2 = leg.add_run("  ○  Supporting alignment")
r2.font.size = Pt(9); r2.font.color.rgb = NAVY

# ════════════════════════════════════════════════════════════════════════════
# Back to portrait
# ════════════════════════════════════════════════════════════════════════════
port_sec = doc.add_section(WD_SECTION.NEW_PAGE)
port_sec.orientation   = 0
port_sec.page_width    = Inches(8.5)
port_sec.page_height   = Inches(11)
for attr in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
    setattr(port_sec, attr, Inches(1))
add_header_footer(port_sec, show_header=True)

# ════════════════════════════════════════════════════════════════════════════
# SECTION 5 – BAR CHART
# ════════════════════════════════════════════════════════════════════════════
add_heading(doc, "Program Outcome Coverage — Bar Chart", 1)
no_space_para(doc,
    "Number of required courses providing primary or supporting alignment to each Program Outcome.",
    size=10)

def bar_color(n):
    if n >= 10: return "#27ae60"
    if n >= 4:  return "#f39c12"
    return "#e74c3c"

colors = [bar_color(n) for n in po_counts]

fig, ax = plt.subplots(figsize=(7.5, max(3.5, N_PO * 0.35)))
bars = ax.barh(po_labels[::-1], po_counts[::-1], color=colors[::-1],
               edgecolor="white", height=0.65)
threshold = max(5, len(COURSES) // 4)
ax.axvline(x=threshold, color="#1F497D", linestyle="--", linewidth=1.2)
for bar, count in zip(bars, po_counts[::-1]):
    ax.text(bar.get_width() + 0.15, bar.get_y() + bar.get_height() / 2,
            str(count), va="center", ha="left", fontsize=9, color="#333333")
ax.set_xlabel("Number of Courses", fontsize=10)
ax.set_title("Number of Courses Supporting Each Program Outcome",
             fontsize=11, fontweight="bold", color="#1F497D")
ax.set_xlim(0, max(po_counts) + 3)
ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)
green_p  = mpatches.Patch(color="#27ae60", label="10+ courses (strong)")
yellow_p = mpatches.Patch(color="#f39c12", label="4–9 courses (moderate)")
red_p    = mpatches.Patch(color="#e74c3c", label="0–3 courses (weak / gap)")
ax.legend(handles=[green_p, yellow_p, red_p,
                   plt.Line2D([0], [0], color="#1F497D", linestyle="--",
                              label=f"Min. threshold ({threshold})")],
          loc="lower right", fontsize=8)
plt.tight_layout()

img_buf = io.BytesIO()
plt.savefig(img_buf, format="png", dpi=150, bbox_inches="tight")
img_buf.seek(0)
plt.close()

doc.add_picture(img_buf, width=Inches(6.5))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

# ════════════════════════════════════════════════════════════════════════════
# SECTION 6 – PER-COURSE BREAKDOWN TABLE
# ════════════════════════════════════════════════════════════════════════════
add_page_break(doc)
add_heading(doc, "Per-Course Learning Outcome Breakdown", 1)
no_space_para(doc,
    "The table below lists every required-course learning outcome alongside its mapped Program Outcomes.",
    size=10)

HDR_WIDTHS = [0.85, 1.55, 4.5, 1.1]
lo_tbl = doc.add_table(rows=1, cols=4)
lo_tbl.style = "Table Grid"

hdr = lo_tbl.rows[0]
set_repeat_header(hdr)
for ci, (txt, w) in enumerate(zip(["Course Code", "Course Name", "Learning Outcome", "Mapped POs"], HDR_WIDTHS)):
    set_cell_bg(hdr.cells[ci], NAVY_HEX)
    cell_para(hdr.cells[ci], txt, bold=True, size=8, color=WHITE)
    hdr.cells[ci].width = Inches(w)

course_groups = []
current = None
for row in LO_DATA:
    if row[0] != current:
        course_groups.append([])
        current = row[0]
    course_groups[-1].append(row)

for g_idx, group in enumerate(course_groups):
    bg_hex = PALE_GREY_HEX if g_idx % 2 == 0 else WHITE_HEX
    for r_idx, (code, name, lo, mapped) in enumerate(group):
        row = lo_tbl.add_row()
        for ci, w in enumerate(HDR_WIDTHS):
            row.cells[ci].width = Inches(w)
            set_cell_bg(row.cells[ci], bg_hex)
        cell_para(row.cells[0], code if r_idx == 0 else "", bold=True, size=8)
        cell_para(row.cells[1], name if r_idx == 0 else "", size=8)
        cell_para(row.cells[2], lo,     size=8)
        cell_para(row.cells[3], mapped, size=8, align=WD_ALIGN_PARAGRAPH.CENTER)

# ════════════════════════════════════════════════════════════════════════════
# SECTION 7 – GAP ANALYSIS
# ════════════════════════════════════════════════════════════════════════════
add_page_break(doc)
add_heading(doc, "Gap Analysis and Recommendations", 1)
no_space_para(doc,
    "The following Program Outcomes are either absent from, or disproportionately represented in, "
    "the required curriculum. Each finding is paired with a practical recommendation.",
    size=10)
doc.add_paragraph()

if not GAPS:
    no_space_para(doc, "No significant gaps detected. All Program Outcomes have moderate or strong coverage.", size=10)
else:
    GAP_WIDTHS = [1.6, 3.3, 3.1]
    gap_tbl = doc.add_table(rows=1, cols=3)
    gap_tbl.style = "Table Grid"

    ghdr = gap_tbl.rows[0]
    set_repeat_header(ghdr)
    for ci, (txt, w) in enumerate(zip(["Program Outcome", "Gap Finding", "Recommendation"], GAP_WIDTHS)):
        set_cell_bg(ghdr.cells[ci], NAVY_HEX)
        cell_para(ghdr.cells[ci], txt, bold=True, size=9, color=WHITE)
        ghdr.cells[ci].width = Inches(w)

    ROW_BG = ["FFF0F0", "FFF5E6", "FFFFE6", "F0F5FF", "F0FFF0"]
    for i, (po, finding, rec) in enumerate(GAPS):
        grow = gap_tbl.add_row()
        for ci, w in enumerate(GAP_WIDTHS):
            grow.cells[ci].width = Inches(w)
            set_cell_bg(grow.cells[ci], ROW_BG[i % len(ROW_BG)])
        cell_para(grow.cells[0], po,      bold=True, size=9)
        cell_para(grow.cells[1], finding, size=9)
        cell_para(grow.cells[2], rec,     size=9)

# ════════════════════════════════════════════════════════════════════════════
# SECTION 8 – APPENDIX
# ════════════════════════════════════════════════════════════════════════════
doc.add_paragraph()
add_heading(doc, "Appendix: Scope and Limitations", 1)
no_space_para(doc, (
    f"This report covers required courses for the {PROGRAM_NAME} program only. "
    "Elective courses were not mapped, as student selections vary and electives do not "
    "constitute a shared program experience. Work-integrated learning and non-credit "
    "safety courses (e.g. WHMIS, OH&S) are excluded from the course LO mapping; their "
    "contribution to relevant Program Outcomes should be documented separately in the "
    "program's compliance record."
), size=10)

# ════════════════════════════════════════════════════════════════════════════
# SAVE
# ════════════════════════════════════════════════════════════════════════════
os.makedirs(_PROC_DIR, exist_ok=True)
doc.save(OUTPUT)
print(f"Saved: {OUTPUT}")
