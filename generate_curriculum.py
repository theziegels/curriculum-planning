#!/usr/bin/env python3
"""Generate Homeschool Curriculum Planning workbook — with semesters & electives."""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── palette ───────────────────────────────────────────────────────────────────
NAVY         = "1F3864"
MED_BLUE     = "2E74B5"
LIGHT_BLUE   = "DEEAF1"
SEM1_COLOR   = "1A5E3A"   # dark green  – Semester 1
SEM1_LIGHT   = "D4EFDF"
SEM2_COLOR   = "7B3F00"   # dark amber  – Semester 2
SEM2_LIGHT   = "FDEBD0"
WHITE        = "FFFFFF"
LIGHT_GRAY   = "F2F2F2"
MID_GRAY     = "D9D9D9"
DARK_GRAY    = "595959"
INPUT_YELLOW = "FFFACD"
FORMULA_BG   = "EBF3FB"
DIVIDER_COL  = "BDC3C7"

# ── subjects ──────────────────────────────────────────────────────────────────
# (display name, header fill hex, row fill hex, editable_name)
CORE_SUBJECTS = [
    ("Bible",                    "5B2C6F", "F5EEF8", False),
    ("Math",                     "1A5276", "D6EAF8", False),
    ("English / Language Arts",  "922B21", "FDEDEC", False),
    ("Reading",                  "784212", "FDEBD0", False),
    ("Writing / Composition",    "7D6608", "FEFDE2", False),
    ("Science",                  "0E6655", "D5F5E3", False),
    ("Social Studies / History", "6E2F0E", "FAE5D3", False),
    ("Health",                   "1E8449", "D4EFDF", False),
    ("Physical Education",       "117A65", "D1F2EB", False),
    ("Art / Music",              "76448A", "F4ECF7", False),
]

ELECTIVE_SUBJECTS = [
    ("Elective / Other 1",       "C7540A", "FEF0E7", True),
    ("Elective / Other 2",       "B7770D", "FEF9E7", True),
    ("Elective / Other 3",       "2E4057", "EAF0FB", True),
    ("Elective / Other 4",       "4A235A", "F5EEF8", True),
]

ALL_SUBJECTS = CORE_SUBJECTS + ELECTIVE_SUBJECTS

NUM_STUDENTS = 10

# Column widths for each sheet (10 cols A–J)
COL_WIDTHS_GS = [4, 28, 18, 16, 14, 14, 14, 20, 14, 14]
COL_WIDTHS_ST = [22, 14, 14, 14, 18, 4, 14, 14, 14, 18]


# ── style helpers ─────────────────────────────────────────────────────────────
def fill(hex_color):
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

def fnt(bold=False, size=11, color="000000", italic=False):
    return Font(name="Calibri", size=size, bold=bold, italic=italic, color=color)

def aln(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def bdr(style="thin"):
    s = Side(style=style)
    return Border(left=s, right=s, top=s, bottom=s)

def outer_bdr():
    s = Side(style="medium")
    return Border(left=s, right=s, top=s, bottom=s)

def set_col_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

def sc(cell, value=None, bold=False, size=11, fcolor="000000",
       bg=None, italic=False, h="left", v="center", wrap=False, b=None):
    if value is not None:
        cell.value = value
    cell.font  = fnt(bold=bold, size=size, color=fcolor, italic=italic)
    cell.alignment = aln(h=h, v=v, wrap=wrap)
    if bg:
        cell.fill = fill(bg)
    if b:
        cell.border = b

def lbl(ws, row, col, text, bg=LIGHT_GRAY, bold=True, size=10, color=DARK_GRAY):
    c = ws.cell(row=row, column=col)
    sc(c, value=text, bold=bold, size=size, fcolor=color, bg=bg, b=bdr())

def inp(ws, row, col, value=None, formula=None, bg=INPUT_YELLOW,
        fmt=None, h="left", bold=False, size=10):
    c = ws.cell(row=row, column=col)
    if formula:
        c.value = formula
        c.fill  = fill(FORMULA_BG)
        c.font  = fnt(size=size, bold=True)
    else:
        if value is not None:
            c.value = value
        c.fill = fill(bg)
        c.font = fnt(size=size, bold=bold)
    c.border    = bdr()
    c.alignment = aln(h=h, v="center")
    if fmt:
        c.number_format = fmt

def merge_inp(ws, row, c1, c2, formula=None, default=None,
              bg=INPUT_YELLOW, fmt=None, h="left", bold=False):
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    c = ws.cell(row=row, column=c1)
    if formula:
        c.value = formula
        c.fill  = fill(FORMULA_BG)
        c.font  = fnt(size=10, bold=True)
    else:
        if default is not None:
            c.value = default
        c.fill = fill(bg)
        c.font = fnt(size=10, bold=bold)
    c.border    = bdr()
    c.alignment = aln(h=h, v="center")
    if fmt:
        c.number_format = fmt

def sec_hdr(ws, row, c1, c2, text, bg, fc="FFFFFF", size=12):
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    c = ws.cell(row=row, column=c1)
    sc(c, value=text, bold=True, size=size, fcolor=fc, bg=bg,
       h="center", v="center", b=bdr())
    ws.row_dimensions[row].height = 22


# ── GETTING STARTED SHEET ─────────────────────────────────────────────────────
def build_getting_started(wb):
    ws = wb.create_sheet("Getting Started", 0)
    set_col_widths(ws, COL_WIDTHS_GS)

    # Title
    ws.merge_cells("A1:J1")
    sc(ws["A1"], value="Homeschool Curriculum Planner",
       bold=True, size=22, fcolor=WHITE, bg=NAVY, h="center", v="center")
    ws.row_dimensions[1].height = 48

    ws.merge_cells("A2:J2")
    sc(ws["A2"], value="Getting Started  —  School & Student Information",
       bold=True, size=13, fcolor=WHITE, bg=MED_BLUE, h="center", v="center")
    ws.row_dimensions[2].height = 26

    ws.row_dimensions[3].height = 6

    # ── Settings block ────────────────────────────────────────────────────────
    sec_hdr(ws, 4, 1, 10, "SCHOOL YEAR SETTINGS", NAVY)

    # Basic settings rows 5-7
    basic = [
        (5,  "School Name:",        None,           "Westfield Academy"),
        (6,  "School Year:",        None,           "2025-2026"),
        (7,  "School Days / Week:", None,           5),
    ]
    for row, label, formula, default in basic:
        ws.row_dimensions[row].height = 20
        lbl(ws, row, 2, label)
        merge_inp(ws, row, 3, 6, formula=formula, default=default)

    ws.row_dimensions[8].height = 8

    # Semester 1 sub-header
    sec_hdr(ws, 9, 2, 10, "SEMESTER 1  (September – December)", SEM1_COLOR, size=11)

    sem1_rows = [
        (10, "Sem 1 Start Date:",  None,                              None, "MM/DD/YYYY"),
        (11, "Sem 1 End Date:",    None,                              None, "MM/DD/YYYY"),
        (12, "Sem 1 Weeks:",       '=IFERROR(ROUNDDOWN((C11-C10)/7,0),"")', None, "0"),
    ]
    for row, label, formula, default, fmt in sem1_rows:
        ws.row_dimensions[row].height = 20
        lbl(ws, row, 2, label)
        merge_inp(ws, row, 3, 6, formula=formula, default=default, fmt=fmt)
    ws["C12"].fill = fill(FORMULA_BG)
    ws.merge_cells("G10:J10"); ws["G10"].value = "← Enter as MM/DD/YYYY"
    ws["G10"].font = fnt(size=9, italic=True, color="808080")
    ws.merge_cells("G11:J11"); ws["G11"].value = "← Enter as MM/DD/YYYY"
    ws["G11"].font = fnt(size=9, italic=True, color="808080")
    ws.merge_cells("G12:J12"); ws["G12"].value = "← Auto-calculated"
    ws["G12"].font = fnt(size=9, italic=True, color="808080")

    ws.row_dimensions[13].height = 8

    # Semester 2 sub-header
    sec_hdr(ws, 14, 2, 10, "SEMESTER 2  (January – May)", SEM2_COLOR, size=11)

    sem2_rows = [
        (15, "Sem 2 Start Date:",  None,                              None, "MM/DD/YYYY"),
        (16, "Sem 2 End Date:",    None,                              None, "MM/DD/YYYY"),
        (17, "Sem 2 Weeks:",       '=IFERROR(ROUNDDOWN((C16-C15)/7,0),"")', None, "0"),
    ]
    for row, label, formula, default, fmt in sem2_rows:
        ws.row_dimensions[row].height = 20
        lbl(ws, row, 2, label)
        merge_inp(ws, row, 3, 6, formula=formula, default=default, fmt=fmt)
    ws["C17"].fill = fill(FORMULA_BG)
    ws.merge_cells("G15:J15"); ws["G15"].value = "← Enter as MM/DD/YYYY"
    ws["G15"].font = fnt(size=9, italic=True, color="808080")
    ws.merge_cells("G16:J16"); ws["G16"].value = "← Enter as MM/DD/YYYY"
    ws["G16"].font = fnt(size=9, italic=True, color="808080")
    ws.merge_cells("G17:J17"); ws["G17"].value = "← Auto-calculated"
    ws["G17"].font = fnt(size=9, italic=True, color="808080")

    ws.row_dimensions[18].height = 8

    # Total weeks row
    ws.row_dimensions[19].height = 20
    lbl(ws, 19, 2, "Total School Weeks:", bg=LIGHT_BLUE, bold=True, color=NAVY)
    merge_inp(ws, 19, 3, 6,
              formula='=IFERROR(C12+C17,"")', fmt="0")
    ws.cell(row=19, column=3).fill = fill(FORMULA_BG)
    ws.cell(row=19, column=3).font = fnt(size=11, bold=True, color=NAVY)
    ws.merge_cells("G19:J19"); ws["G19"].value = "← Sem 1 + Sem 2 combined"
    ws["G19"].font = fnt(size=9, italic=True, color="808080")

    ws.row_dimensions[20].height = 8

    # ── Student roster ────────────────────────────────────────────────────────
    sec_hdr(ws, 21, 1, 10, "STUDENT ROSTER", NAVY)

    # Column headers row 22
    ws.row_dimensions[22].height = 32
    col_spans = [(1,1),(2,3),(4,4),(5,5),(6,6),(7,8),(9,10)]
    col_labels = ["#", "Student Name", "Date of Birth",
                  "Class of\n(Grad Year)", "Grade\n(Auto)", "Curriculum Notes", ""]
    for (c1, c2), lbl_text in zip(col_spans, col_labels):
        if c1 != c2:
            ws.merge_cells(start_row=22, start_column=c1, end_row=22, end_column=c2)
        c = ws.cell(row=22, column=c1)
        sc(c, value=lbl_text, bold=True, size=10, fcolor=WHITE,
           bg=MED_BLUE, h="center", v="center", wrap=True, b=bdr())

    for i in range(NUM_STUDENTS):
        row = 23 + i
        ws.row_dimensions[row].height = 20

        # # badge
        c = ws.cell(row=row, column=1)
        sc(c, value=i+1, bold=True, size=10, fcolor=WHITE,
           bg=NAVY if i % 2 == 0 else MED_BLUE, h="center", b=bdr())

        # Name (cols 2-3)
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
        inp(ws, row, 2)

        # DOB col 4
        inp(ws, row, 4, fmt="MM/DD/YYYY")

        # Class of col 5
        inp(ws, row, 5, fmt="0")

        # Grade formula col 6
        c = ws.cell(row=row, column=6)
        c.value = (
            f'=IF(OR(B{row}="",E{row}=""),"",LET(g,12-(E{row}-YEAR($C$16)),'
            f'IF(g=11,"11th",IF(g=12,"12th",IF(g=1,"1st",'
            f'IF(g=2,"2nd",IF(g=3,"3rd",g&"th")))))))'
        )
        c.fill = fill(FORMULA_BG); c.border = bdr()
        c.font = fnt(size=10, bold=True); c.alignment = aln(h="center")

        # Notes (cols 7-10)
        ws.merge_cells(start_row=row, start_column=7, end_row=row, end_column=10)
        inp(ws, row, 7)

    # ── Legend ────────────────────────────────────────────────────────────────
    r = 23 + NUM_STUDENTS + 1
    ws.row_dimensions[r].height = 6
    r += 1

    sec_hdr(ws, r, 1, 10, "LEGEND & INSTRUCTIONS", DARK_GRAY)
    r += 1

    for color, desc in [(INPUT_YELLOW, "User Input Cell — type your data here"),
                        (FORMULA_BG,   "Formula Cell — auto-calculated, do not edit")]:
        ws.row_dimensions[r].height = 18
        ws.cell(row=r, column=2).fill = fill(color)
        ws.cell(row=r, column=2).border = bdr()
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=10)
        sc(ws.cell(row=r, column=3), value=desc, size=10, fcolor=DARK_GRAY)
        r += 1

    instructions = [
        "1.  Fill in School Year Settings: school name, year, days/week, and both semester dates.",
        "2.  Semester 1 and Semester 2 weeks auto-calculate from the dates you enter.",
        "3.  Enter each student's name, DOB, and graduation year — grade level auto-fills.",
        "4.  Open each Student tab and fill in each subject's curriculum title and unit counts.",
        "5.  Enter units separately for Semester 1 and Semester 2, or just fill in one semester.",
        "6.  Pace (units/week, units/day, est. end date) auto-calculates for each semester.",
        "7.  Use the Component Breakdown table for multi-volume curricula (e.g., 10 math books).",
        "8.  Elective rows have an editable name — just type over 'Elective / Other 1' etc.",
    ]
    r += 1
    sec_hdr(ws, r, 1, 10, "HOW TO USE", "375623")
    r += 1
    for i, text in enumerate(instructions):
        ws.row_dimensions[r].height = 18
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=10)
        sc(ws.cell(row=r, column=2), value=text, size=10, fcolor=DARK_GRAY,
           bg="F0FFF4" if i % 2 == 0 else WHITE)
        r += 1

    return ws


# ── STUDENT SHEET ─────────────────────────────────────────────────────────────
# Getting Started cell references (fixed)
GS = "Getting Started"
GS_S1_START  = f"'{GS}'!C10"
GS_S1_END    = f"'{GS}'!C11"
GS_S1_WEEKS  = f"'{GS}'!C12"
GS_S2_START  = f"'{GS}'!C15"
GS_S2_END    = f"'{GS}'!C16"
GS_S2_WEEKS  = f"'{GS}'!C17"
GS_DAYS_WEEK = f"'{GS}'!C7"


def build_student_sheet(wb, idx):
    n       = idx + 1
    gs_row  = 23 + idx  # student data row in Getting Started

    ws = wb.create_sheet(f"Student {n}")

    # 10 columns; col A is wide (subject labels span it alone)
    # Layout: A=label(22), B=S1units(14), C=S1/wk(14), D=S1/day(14), E=S1end(18),
    #         F=divider(4), G=S2units(14), H=S2/wk(14), I=S2/day(14), J=S2end(18)
    for i, w in enumerate(COL_WIDTHS_ST, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ── Banner ────────────────────────────────────────────────────────────────
    ws.merge_cells("A1:J1")
    ws["A1"].value = f"=IF('{GS}'!B{gs_row}=\"\",\"Student {n}\",'{GS}'!B{gs_row})"
    sc(ws["A1"], bold=True, size=20, fcolor=WHITE, bg=NAVY, h="center", v="center")
    ws.row_dimensions[1].height = 44

    ws.merge_cells("A2:J2")
    ws["A2"].value = (
        f'=IFERROR("Grade: "&\'{GS}\'!F{gs_row}'
        f'&"   ·   School Year: "&\'{GS}\'!C6",'
        f'"")'
    )
    # cleaner:
    ws["A2"].value = (
        f"=IFERROR(\"Grade: \"&'{GS}'!F{gs_row}"
        f"&\"   ·   School Year: \"&'{GS}'!C6,\"\")"
    )
    sc(ws["A2"], size=11, fcolor=WHITE, bg=MED_BLUE, h="center", v="center")
    ws.row_dimensions[2].height = 22

    ws.row_dimensions[3].height = 6

    # ── Semester overview strip ───────────────────────────────────────────────
    sec_hdr(ws, 4, 1, 10, "SEMESTER OVERVIEW  (auto-populated from Getting Started)", NAVY)

    ws.row_dimensions[5].height = 20

    # S1 block (cols 1-5)
    ws.merge_cells("A5:E5")
    ws["A5"].value = (
        f'=IFERROR("SEMESTER 1:  "'
        f'&TEXT({GS_S1_START},"MMM D")&"  –  "'
        f'&TEXT({GS_S1_END},"MMM D, YYYY")'
        f'&"   ("&{GS_S1_WEEKS}&" weeks)","SEMESTER 1")'
    )
    sc(ws["A5"], bold=True, size=11, fcolor=WHITE, bg=SEM1_COLOR, h="center", b=bdr())

    # divider
    ws.cell(row=5, column=6).fill = fill(DIVIDER_COL)

    # S2 block (cols 7-10)
    ws.merge_cells("G5:J5")
    ws["G5"].value = (
        f'=IFERROR("SEMESTER 2:  "'
        f'&TEXT({GS_S2_START},"MMM D")&"  –  "'
        f'&TEXT({GS_S2_END},"MMM D, YYYY")'
        f'&"   ("&{GS_S2_WEEKS}&" weeks)","SEMESTER 2")'
    )
    sc(ws["G5"], bold=True, size=11, fcolor=WHITE, bg=SEM2_COLOR, h="center", b=bdr())

    # ── Column labels (row 6) ─────────────────────────────────────────────────
    ws.row_dimensions[6].height = 28

    col_hdrs = [
        (1, 1, "Subject / Curriculum",      NAVY,      WHITE),
        (2, 2, "Sem 1\nUnits",              SEM1_COLOR, WHITE),
        (3, 3, "Sem 1\n/ Week",             SEM1_COLOR, WHITE),
        (4, 4, "Sem 1\n/ Day",              SEM1_COLOR, WHITE),
        (5, 5, "Sem 1\nEst. End",           SEM1_COLOR, WHITE),
        (6, 6, "",                           DIVIDER_COL, DIVIDER_COL),
        (7, 7, "Sem 2\nUnits",              SEM2_COLOR, WHITE),
        (8, 8, "Sem 2\n/ Week",             SEM2_COLOR, WHITE),
        (9, 9, "Sem 2\n/ Day",              SEM2_COLOR, WHITE),
        (10,10,"Sem 2\nEst. End",           SEM2_COLOR, WHITE),
    ]
    for c1, c2, text, bg, fc in col_hdrs:
        if c1 != c2:
            ws.merge_cells(start_row=6, start_column=c1, end_row=6, end_column=c2)
        c = ws.cell(row=6, column=c1)
        sc(c, value=text, bold=True, size=9, fcolor=fc, bg=bg,
           h="center", v="center", wrap=True, b=bdr())

    ws.row_dimensions[7].height = 6   # spacer before first subject

    # ── Subject sections ──────────────────────────────────────────────────────
    current_row = 8
    for subj_name, hdr_fill, row_fill, editable in ALL_SUBJECTS:
        current_row = add_subject_section(
            ws, current_row, subj_name, hdr_fill, row_fill, editable
        )
        ws.row_dimensions[current_row].height = 6   # spacer
        current_row += 1

    return ws


def add_subject_section(ws, r0, subj_name, hdr_fill, row_fill, editable_name):
    """
    Rows relative to r0:
      r0+0  Subject header (editable name input if elective, else static label)
      r0+1  Curriculum / book title  +  unit-type label
      r0+2  Pace row: S1 units|/wk|/day|est.end  ||  S2 units|/wk|/day|est.end
      r0+3  Component breakdown header
      r0+4  Component column headers
      r0+5..r0+10  6 component rows
      r0+11 Notes
    Returns first row after section.
    """
    # ── Row 0: subject header ─────────────────────────────────────────────────
    ws.row_dimensions[r0].height = 24

    if editable_name:
        # Left part: static "ELECTIVE" badge (cols 1-2), editable name (cols 3-10)
        ws.merge_cells(start_row=r0, start_column=1, end_row=r0, end_column=2)
        c = ws.cell(row=r0, column=1)
        sc(c, value="ELECTIVE / OTHER", bold=True, size=11, fcolor=WHITE,
           bg=hdr_fill, h="center", v="center", b=bdr())
        ws.merge_cells(start_row=r0, start_column=3, end_row=r0, end_column=10)
        c = ws.cell(row=r0, column=3)
        sc(c, value=subj_name, bold=True, size=12, fcolor=hdr_fill,
           bg=INPUT_YELLOW, h="left", v="center", b=outer_bdr())
        c.font = fnt(bold=True, size=12, color=hdr_fill)
    else:
        ws.merge_cells(start_row=r0, start_column=1, end_row=r0, end_column=10)
        c = ws.cell(row=r0, column=1)
        sc(c, value=f"  {subj_name.upper()}", bold=True, size=13, fcolor=WHITE,
           bg=hdr_fill, h="left", v="center", b=outer_bdr())

    # ── Row 1: curriculum title + unit type ───────────────────────────────────
    r1 = r0 + 1
    ws.row_dimensions[r1].height = 20

    lbl(ws, r1, 1, "Curriculum / Title:", bg=row_fill)
    ws.merge_cells(start_row=r1, start_column=2, end_row=r1, end_column=5)
    inp(ws, r1, 2)   # user types curriculum name

    # divider
    ws.cell(row=r1, column=6).fill = fill(DIVIDER_COL)

    lbl(ws, r1, 7, "Unit type:", bg=row_fill)
    ws.merge_cells(start_row=r1, start_column=8, end_row=r1, end_column=10)
    c = ws.cell(row=r1, column=8)
    c.value = "pages"; c.fill = fill(INPUT_YELLOW)
    c.border = bdr(); c.font = fnt(size=10)

    # ── Row 2: pace strip ─────────────────────────────────────────────────────
    r2 = r0 + 2
    ws.row_dimensions[r2].height = 22

    # S1 cells: col B=units input, C=/wk, D=/day, E=est end
    # S2 cells: col G=units input, H=/wk, I=/day, J=est end

    s1u = f"B{r2}"; s1w = f"C{r2}"; s1d = f"D{r2}"
    s2u = f"G{r2}"; s2w = f"H{r2}"; s2d = f"I{r2}"

    # S1 units – input
    inp(ws, r2, 2, h="center")

    # S1 /week
    c = ws.cell(row=r2, column=3)
    c.value = f'=IFERROR(IF({s1u}="","",ROUND({s1u}/{GS_S1_WEEKS},1)),"")'
    c.fill = fill(SEM1_LIGHT); c.border = bdr()
    c.font = fnt(size=10, bold=True); c.alignment = aln(h="center")

    # S1 /day
    c = ws.cell(row=r2, column=4)
    c.value = f'=IFERROR(IF({s1w}="","",ROUND({s1w}/{GS_DAYS_WEEK},1)),"")'
    c.fill = fill(SEM1_LIGHT); c.border = bdr()
    c.font = fnt(size=10, bold=True); c.alignment = aln(h="center")

    # S1 est end
    c = ws.cell(row=r2, column=5)
    c.value = (
        f'=IFERROR(IF({s1u}="","",TEXT('
        f'{GS_S1_START}+({s1u}/{s1d}),"MMM D, YYYY")),"")'
    )
    c.fill = fill(SEM1_LIGHT); c.border = bdr()
    c.font = fnt(size=10, bold=True); c.alignment = aln(h="center")

    # divider col F
    ws.cell(row=r2, column=6).fill = fill(DIVIDER_COL)

    # S2 units – input
    inp(ws, r2, 7, h="center")

    # S2 /week
    c = ws.cell(row=r2, column=8)
    c.value = f'=IFERROR(IF({s2u}="","",ROUND({s2u}/{GS_S2_WEEKS},1)),"")'
    c.fill = fill(SEM2_LIGHT); c.border = bdr()
    c.font = fnt(size=10, bold=True); c.alignment = aln(h="center")

    # S2 /day
    c = ws.cell(row=r2, column=9)
    c.value = f'=IFERROR(IF({s2w}="","",ROUND({s2w}/{GS_DAYS_WEEK},1)),"")'
    c.fill = fill(SEM2_LIGHT); c.border = bdr()
    c.font = fnt(size=10, bold=True); c.alignment = aln(h="center")

    # S2 est end
    c = ws.cell(row=r2, column=10)
    c.value = (
        f'=IFERROR(IF({s2u}="","",TEXT('
        f'{GS_S2_START}+({s2u}/{s2d}),"MMM D, YYYY")),"")'
    )
    c.fill = fill(SEM2_LIGHT); c.border = bdr()
    c.font = fnt(size=10, bold=True); c.alignment = aln(h="center")

    # ── Row 3: component breakdown sub-header ─────────────────────────────────
    r3 = r0 + 3
    ws.row_dimensions[r3].height = 16
    ws.merge_cells(start_row=r3, start_column=1, end_row=r3, end_column=10)
    sc(ws.cell(row=r3, column=1),
       value="  Optional: Component / Volume Breakdown  "
             "(fill in if the curriculum has multiple books or parts)",
       size=9, italic=True, fcolor=WHITE, bg=DARK_GRAY, h="left", b=bdr())

    # ── Row 4: component column headers ──────────────────────────────────────
    r4 = r0 + 4
    ws.row_dimensions[r4].height = 28

    comp_cols = [
        (1, 1, "#"),
        (2, 4, "Component / Book Title"),
        (5, 5, "Units in\nComponent"),
        (6, 6, "Cumul.\nUnits"),
        (7, 7, "Start\nWeek"),
        (8, 8, "End\nWeek"),
        (9, 9, "Start\nDate"),
        (10,10,"End\nDate"),
    ]
    for c1, c2, hdr_text in comp_cols:
        if c1 != c2:
            ws.merge_cells(start_row=r4, start_column=c1, end_row=r4, end_column=c2)
        c = ws.cell(row=r4, column=c1)
        sc(c, value=hdr_text, bold=True, size=9, fcolor=WHITE,
           bg=DARK_GRAY, h="center", v="center", wrap=True, b=bdr())

    # ── Rows 5-10: 6 component rows ───────────────────────────────────────────
    comp_start = r0 + 5
    # units/week for this subject's S1 and S2 (we'll use full-year combined pace for components)
    # For component scheduling we use S1+S2 total pace. Use a combined row.
    # Actually let's key components off total units across both semesters.
    # total_subj_units = B{r2} + G{r2}
    # combined /week = (S1units + S2units) / total_weeks
    total_weeks_ref = f"(IFERROR({GS_S1_WEEKS},0)+IFERROR({GS_S2_WEEKS},0))"

    for ci in range(6):
        r = comp_start + ci
        ws.row_dimensions[r].height = 18
        alt = INPUT_YELLOW if ci % 2 == 0 else "FAFAD2"

        # # col
        c = ws.cell(row=r, column=1)
        sc(c, value=ci+1, bold=True, size=10, fcolor=WHITE,
           bg=NAVY if ci % 2 == 0 else MED_BLUE, h="center", b=bdr())

        # title cols 2-4
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        c = ws.cell(row=r, column=2)
        c.fill = fill(alt); c.border = bdr(); c.font = fnt(size=10)

        # units col 5
        c = ws.cell(row=r, column=5)
        c.fill = fill(alt); c.border = bdr()
        c.font = fnt(size=10); c.alignment = aln(h="center")

        # cumulative col 6
        c = ws.cell(row=r, column=6)
        c.value = f'=IFERROR(SUM(E{comp_start}:E{r}),"")'
        c.fill = fill(FORMULA_BG); c.border = bdr()
        c.font = fnt(size=10); c.alignment = aln(h="center")

        # combined subj /week for component scheduling
        s1u_ref = f"B{r2}"; s2u_ref = f"G{r2}"
        upw_expr = (
            f'(IFERROR({s1u_ref},0)+IFERROR({s2u_ref},0))/{total_weeks_ref}'
        )

        # start week col 7
        c = ws.cell(row=r, column=7)
        if ci == 0:
            c.value = f'=IF(E{r}="","",1)'
        else:
            prev_cum = f"F{r-1}"
            c.value = f'=IFERROR(IF(E{r}="","",FLOOR({prev_cum}/({upw_expr}),1)+1),"")'
        c.fill = fill(FORMULA_BG); c.border = bdr()
        c.font = fnt(size=10); c.alignment = aln(h="center")

        # end week col 8
        cum = f"F{r}"
        c = ws.cell(row=r, column=8)
        c.value = f'=IFERROR(IF(E{r}="","",CEILING({cum}/({upw_expr}),1)),"")'
        c.fill = fill(FORMULA_BG); c.border = bdr()
        c.font = fnt(size=10); c.alignment = aln(h="center")

        # start date col 9
        sw = f"G{r}"
        c = ws.cell(row=r, column=9)
        c.value = (
            f'=IFERROR(IF(E{r}="","",TEXT({GS_S1_START}+({sw}-1)*7,"MMM D")),"")'
        )
        c.fill = fill(FORMULA_BG); c.border = bdr()
        c.font = fnt(size=10); c.alignment = aln(h="center")

        # end date col 10
        ew = f"H{r}"
        c = ws.cell(row=r, column=10)
        c.value = (
            f'=IFERROR(IF(E{r}="","",TEXT({GS_S1_START}+({ew})*7,"MMM D")),"")'
        )
        c.fill = fill(FORMULA_BG); c.border = bdr()
        c.font = fnt(size=10); c.alignment = aln(h="center")

    # ── Notes row ─────────────────────────────────────────────────────────────
    r_notes = comp_start + 6
    ws.row_dimensions[r_notes].height = 20
    lbl(ws, r_notes, 1, "Notes:", bg=row_fill, size=10)
    ws.merge_cells(start_row=r_notes, start_column=2,
                   end_row=r_notes, end_column=10)
    c = ws.cell(row=r_notes, column=2)
    c.fill = fill(INPUT_YELLOW); c.border = bdr()
    c.font = fnt(size=10); c.alignment = aln(h="left", v="center")

    return r_notes + 1   # next available row


# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    build_getting_started(wb)

    tab_colors = [
        "1F3864","2E74B5","5B2C6F","922B21","784212",
        "7D6608","0E6655","6E2F0E","1E8449","76448A",
    ]
    for i in range(NUM_STUDENTS):
        ws = build_student_sheet(wb, i)
        ws.sheet_properties.tabColor = tab_colors[i % len(tab_colors)]

    wb["Getting Started"].sheet_properties.tabColor = "1F3864"

    out = "/home/user/curriculum-planning/Homeschool_Curriculum_Planner.xlsx"
    wb.save(out)
    print(f"Saved: {out}")


if __name__ == "__main__":
    main()
