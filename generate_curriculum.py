#!/usr/bin/env python3
"""Homeschool Curriculum Planner — warm neutrals, semester split, Course Content hub."""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

# ── Palette ───────────────────────────────────────────────────────────────────
ACCENT       = "CE8282"   # brand rose  — used sparingly
ACCENT_PALE  = "FAF0F0"   # very light rose tint (highlight)
WARM_900     = "3D3530"   # near-black warm brown
WARM_700     = "6B5C55"   # medium warm taupe  (labels)
WARM_500     = "A08878"   # lighter taupe      (hints)
WARM_300     = "D4C4BB"   # border / divider
WARM_200     = "EDE3DC"   # section header bg
WARM_100     = "F7F2EE"   # formula cell bg
LINEN        = "F5EDE3"   # subject-header bg
WHITE        = "FFFFFF"   # input cells
OFF_WHITE    = "FAF8F6"   # alt rows

# Semester tints (light, non-distracting)
S1_HDR   = "5C7A52"   # sage green  for Sem1 col header
S1_LIGHT = "EEF4EB"   # very light sage
S2_HDR   = "8A6A30"   # warm amber  for Sem2 col header
S2_LIGHT = "F8F2E6"   # very light amber

# ── Per-student colors (header bg / tab color) ────────────────────────────────
# Light-to-medium, print-friendly, clearly distinct from each other
STUDENT_COLORS = [
    ("CE8282", "FAEEEE"),   # 1  dusty rose      (brand accent)
    ("7A9E7E", "EDF4EE"),   # 2  sage green
    ("7AA0C4", "EDF3FA"),   # 3  sky blue
    ("C4A462", "FAF3E3"),   # 4  warm amber
    ("9B8EC4", "F1EEF9"),   # 5  soft lavender
    ("6AADA6", "E8F5F4"),   # 6  muted teal
    ("C47A6A", "FAF0EE"),   # 7  warm terracotta
    ("7A8EC4", "EEF1FA"),   # 8  periwinkle
    ("8FA87A", "F1F5EE"),   # 9  olive green
    ("A87A9E", "F5EEF4"),   # 10 soft plum
]   # (header_fill, row_light_tint)

# ── Subjects ──────────────────────────────────────────────────────────────────
# (name, editable)  — colours come from palette now, not per-subject
CORE = [
    ("Bible",                   False),
    ("Math",                    False),
    ("English / Language Arts", False),
    ("Reading",                 False),
    ("Writing / Composition",   False),
    ("Science",                 False),
    ("Social Studies / History",False),
    ("Health",                  False),
    ("Physical Education",      False),
    ("Art / Music",             False),
]
ELECTIVES = [
    ("Elective / Other 1", True),
    ("Elective / Other 2", True),
    ("Elective / Other 3", True),
    ("Elective / Other 4", True),
]
ALL_SUBJECTS = CORE + ELECTIVES   # 14 total

NUM_STUDENTS  = 10
NUM_SUBJECTS  = len(ALL_SUBJECTS)   # 14
CC_HDR_ROW    = 7    # Course Content column-header row
CC_DATA_START = 8    # first data row in Course Content


# ── Style helpers ─────────────────────────────────────────────────────────────
def fill(h):
    return PatternFill(start_color=h, end_color=h, fill_type="solid")

def fnt(size=10, bold=False, color=WARM_900, italic=False):
    return Font(name="Calibri", size=size, bold=bold, color=color, italic=italic)

def aln(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def side(style="thin", color=WARM_300):
    return Side(style=style, color=color)

def box(style="thin", color=WARM_300):
    s = side(style, color)
    return Border(left=s, right=s, top=s, bottom=s)

def left_accent():
    """Thick left border in brand accent, thin others in warm gray."""
    return Border(
        left=Side(style="medium", color=ACCENT),
        right=side(),
        top=side(),
        bottom=side(),
    )

def bottom_only(color=WARM_300):
    return Border(bottom=Side(style="thin", color=color))

def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def cell(ws, row, col):
    return ws.cell(row=row, column=col)

def sc(c, value=None, size=10, bold=False, color=WARM_900, italic=False,
       bg=None, h="left", v="center", wrap=False, b=None):
    if value is not None:
        c.value = value
    c.font      = fnt(size=size, bold=bold, color=color, italic=italic)
    c.alignment = aln(h=h, v=v, wrap=wrap)
    if bg is not None:
        c.fill = fill(bg)
    if b is not None:
        c.border = b
    return c

def lbl(ws, row, col, text, bg=WHITE, bold=True, color=WARM_700, size=10):
    c = ws.cell(row=row, column=col)
    sc(c, value=text, size=size, bold=bold, color=color, bg=bg, b=box())

def inp(ws, row, col, value=None, h="left", fmt=None, bg=WHITE):
    c = ws.cell(row=row, column=col)
    c.fill      = fill(bg)
    c.border    = box()
    c.font      = fnt(size=10)
    c.alignment = aln(h=h, v="center")
    if value is not None:
        c.value = value
    if fmt:
        c.number_format = fmt
    return c

def formula_cell(ws, row, col, formula, h="center", bold=True, size=10):
    c = ws.cell(row=row, column=col)
    c.value     = formula
    c.fill      = fill(WARM_100)
    c.border    = box()
    c.font      = fnt(size=size, bold=bold, color=WARM_700)
    c.alignment = aln(h=h, v="center")
    return c

def merge_sc(ws, r, c1, c2, **kwargs):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    return sc(ws.cell(r, c1), **kwargs)

def merge_inp(ws, r, c1, c2, value=None, formula=None, fmt=None, h="left", default=None):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    c = ws.cell(r, c1)
    if formula:
        c.value     = formula
        c.fill      = fill(WARM_100)
        c.font      = fnt(size=10, bold=True, color=WARM_700)
    else:
        v = value if value is not None else default
        if v is not None:
            c.value = v
        c.fill = fill(WHITE)
        c.font = fnt(size=10)
    c.border    = box()
    c.alignment = aln(h=h, v="center")
    if fmt:
        c.number_format = fmt
    return c

def title_bar(ws, row, c1, c2, text, bg=ACCENT, fc=WHITE, size=16, height=44):
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    c = ws.cell(row, c1)
    sc(c, value=text, size=size, bold=True, color=fc, bg=bg, h="center", v="center")
    ws.row_dimensions[row].height = height
    return c

def sub_banner(ws, row, c1, c2, text, bg=WARM_200, fc=WARM_700, size=11, height=22):
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    c = ws.cell(row, c1)
    sc(c, value=text, size=size, bold=True, color=fc, bg=bg, h="center", v="center")
    ws.row_dimensions[row].height = height
    return c

def section_label(ws, row, c1, c2, text, height=20):
    """Linen bg, accent left border, warm bold text."""
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    c = ws.cell(row, c1)
    c.value     = "  " + text
    c.font      = fnt(size=11, bold=True, color=WARM_700)
    c.fill      = fill(LINEN)
    c.border    = left_accent()
    c.alignment = aln(h="left", v="center")
    ws.row_dimensions[row].height = height
    return c

def spacer(ws, row, height=8):
    ws.row_dimensions[row].height = height


# ── GETTING STARTED ───────────────────────────────────────────────────────────
# GS cell refs used elsewhere
GS_S1_START  = "'Getting Started'!C10"
GS_S1_END    = "'Getting Started'!C11"
# C12 = S1 break start, C13 = S1 break end, C14 = S1 break weeks
GS_S1_WEEKS  = "'Getting Started'!C15"   # school weeks after break
GS_S2_START  = "'Getting Started'!C18"
GS_S2_END    = "'Getting Started'!C19"
# C20 = S2 break start, C21 = S2 break end, C22 = S2 break weeks
GS_S2_WEEKS  = "'Getting Started'!C23"   # school weeks after break
GS_DAYS      = "'Getting Started'!C7"
GS_STUDENT_ROW = 29   # first student data row in Getting Started (1-based)

def gs_student_name(i):   # i = 1-based
    return f"'Getting Started'!B{GS_STUDENT_ROW - 1 + i}"

def build_getting_started(wb):
    ws = wb.create_sheet("Getting Started", 0)
    # cols: A(4) B(28) C(20) D(14) E(14) F(14) G(14) H(20) I(14) J(14)
    set_widths(ws, [4, 28, 20, 14, 14, 14, 14, 20, 14, 14])

    title_bar(ws, 1, 1, 10, "Homeschool Curriculum Planner")

    sub_banner(ws, 2, 1, 10, "Getting Started  ·  School & Student Information")
    spacer(ws, 3)

    # Settings block header
    sub_banner(ws, 4, 2, 10, "SCHOOL YEAR SETTINGS", bg=WARM_200, size=10, height=20)

    # Basic settings rows 5–7
    for row, label, default, fmt in [
        (5, "School Name",        "Westfield Academy", None),
        (6, "School Year",        "2025-2026",         None),
        (7, "School Days / Week", 5,                   None),
    ]:
        ws.row_dimensions[row].height = 20
        lbl(ws, row, 2, label + ":", bg=WHITE, color=WARM_700)
        merge_inp(ws, row, 3, 6, value=default, fmt=fmt)

    spacer(ws, 8)

    # ── Semester 1 ────────────────────────────────────────────────────────────
    ws.merge_cells("B9:J9")
    sc(ws.cell(9, 2), value="Semester 1  ·  September – December",
       size=10, bold=True, color=S1_HDR, bg=S1_LIGHT, h="center")
    ws.row_dimensions[9].height = 20

    # Sem 1 date inputs (rows 10–11) and break inputs (rows 12–13)
    s1_rows = [
        (10, "Sem 1 Start Date",           None,  "MM/DD/YYYY", "← MM/DD/YYYY"),
        (11, "Sem 1 End Date",             None,  "MM/DD/YYYY", "← MM/DD/YYYY"),
        (12, "Sem 1 Break Start",          None,  "MM/DD/YYYY", "← optional  (e.g. Thanksgiving)"),
        (13, "Sem 1 Break End",            None,  "MM/DD/YYYY", "← optional"),
        (14, "Sem 1 Break Weeks",
             '=IFERROR(IF(OR(C12="",C13=""),0,ROUNDDOWN((C13-C12)/7,0)),0)', "0", "← auto"),
        (15, "Sem 1 School Weeks",
             '=IFERROR(ROUNDDOWN((C11-C10)/7,0)-C14,"")', "0", "← auto  (total minus break)"),
    ]
    for row, label, formula, fmt, hint in s1_rows:
        ws.row_dimensions[row].height = 20
        is_break = "Break" in label and "Weeks" not in label
        lbl(ws, row, 2, label + ":",
            bg=WHITE, color=WARM_500 if is_break else WARM_700)
        merge_inp(ws, row, 3, 6, formula=formula, fmt=fmt)
        ws.merge_cells(f"G{row}:J{row}")
        sc(ws.cell(row, 7), value=hint, size=9, italic=True, color=WARM_500)

    # Date picker validation — Sem 1 semester dates + break dates
    dv_s1 = DataValidation(type="date", operator="between",
                           formula1="DATE(2000,1,1)", formula2="DATE(2099,12,31)",
                           showErrorMessage=False, showInputMessage=False)
    ws.add_data_validation(dv_s1)
    for r in (10, 11, 12, 13):
        dv_s1.add(f"C{r}:F{r}")

    spacer(ws, 16)

    # ── Semester 2 ────────────────────────────────────────────────────────────
    ws.merge_cells("B17:J17")
    sc(ws.cell(17, 2), value="Semester 2  ·  January – May",
       size=10, bold=True, color=S2_HDR, bg=S2_LIGHT, h="center")
    ws.row_dimensions[17].height = 20

    s2_rows = [
        (18, "Sem 2 Start Date",           None,  "MM/DD/YYYY", "← MM/DD/YYYY"),
        (19, "Sem 2 End Date",             None,  "MM/DD/YYYY", "← MM/DD/YYYY"),
        (20, "Sem 2 Break Start",          None,  "MM/DD/YYYY", "← optional  (e.g. Spring / Easter)"),
        (21, "Sem 2 Break End",            None,  "MM/DD/YYYY", "← optional"),
        (22, "Sem 2 Break Weeks",
             '=IFERROR(IF(OR(C20="",C21=""),0,ROUNDDOWN((C21-C20)/7,0)),0)', "0", "← auto"),
        (23, "Sem 2 School Weeks",
             '=IFERROR(ROUNDDOWN((C19-C18)/7,0)-C22,"")', "0", "← auto  (total minus break)"),
    ]
    for row, label, formula, fmt, hint in s2_rows:
        ws.row_dimensions[row].height = 20
        is_break = "Break" in label and "Weeks" not in label
        lbl(ws, row, 2, label + ":",
            bg=WHITE, color=WARM_500 if is_break else WARM_700)
        merge_inp(ws, row, 3, 6, formula=formula, fmt=fmt)
        ws.merge_cells(f"G{row}:J{row}")
        sc(ws.cell(row, 7), value=hint, size=9, italic=True, color=WARM_500)

    dv_s2 = DataValidation(type="date", operator="between",
                           formula1="DATE(2000,1,1)", formula2="DATE(2099,12,31)",
                           showErrorMessage=False, showInputMessage=False)
    ws.add_data_validation(dv_s2)
    for r in (18, 19, 20, 21):
        dv_s2.add(f"C{r}:F{r}")

    spacer(ws, 24)

    # ── Total weeks ───────────────────────────────────────────────────────────
    ws.row_dimensions[25].height = 20
    lbl(ws, 25, 2, "Total School Weeks:", bg=ACCENT_PALE, color=ACCENT, bold=True)
    merge_inp(ws, 25, 3, 6, formula='=IFERROR(C15+C23,"")', fmt="0")
    ws.cell(25, 3).fill = fill(WARM_100)
    ws.cell(25, 3).font = fnt(size=11, bold=True, color=WARM_700)
    ws.merge_cells("G25:J25")
    sc(ws.cell(25, 7), value="← Sem 1 + Sem 2 (breaks excluded)",
       size=9, italic=True, color=WARM_500)

    spacer(ws, 26)

    # ── Student roster ────────────────────────────────────────────────────────
    sub_banner(ws, 27, 1, 10, "STUDENT ROSTER", bg=WARM_200, size=10, height=20)

    # Col headers row 28
    ws.row_dimensions[28].height = 30
    hdr_def = [
        (1, 1, "#"),
        (2, 3, "Student Name"),
        (4, 4, "Date of Birth"),
        (5, 5, "Class of\n(Grad Year)"),
        (6, 6, "Grade\n(Auto)"),
        (7, 10, "Curriculum Notes"),
    ]
    for c1, c2, txt in hdr_def:
        if c1 != c2:
            ws.merge_cells(start_row=28, start_column=c1, end_row=28, end_column=c2)
        c = ws.cell(28, c1)
        sc(c, value=txt, size=10, bold=True, color=WARM_700, bg=WARM_200,
           h="center", v="center", wrap=True, b=box(color=WARM_300))

    for i in range(NUM_STUDENTS):
        row = GS_STUDENT_ROW + i
        ws.row_dimensions[row].height = 22
        bg = WHITE if i % 2 == 0 else OFF_WHITE

        # Badge
        c = ws.cell(row, 1)
        sc(c, value=i+1, size=10, bold=True, color=WHITE, bg=ACCENT,
           h="center", b=box(color=WARM_300))

        # Name
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
        inp(ws, row, 2, bg=bg)

        # DOB
        inp(ws, row, 4, fmt="MM/DD/YYYY", bg=bg)

        # Class of
        inp(ws, row, 5, fmt="0", bg=bg)

        # Grade formula
        c = ws.cell(row, 6)
        c.value = (
            f'=IF(OR(B{row}="",E{row}=""),"",LET(yr,IF($C$19<>"",YEAR($C$19),IF($C$6<>"",VALUE(RIGHT($C$6,4)),YEAR(TODAY()))),g,12-(E{row}-yr),'
            f'IF(g=11,"11th",IF(g=12,"12th",IF(g=1,"1st",'
            f'IF(g=2,"2nd",IF(g=3,"3rd",g&"th")))))))'
        )
        c.fill = fill(WARM_100); c.border = box(color=WARM_300)
        c.font = fnt(size=10, bold=True, color=WARM_700)
        c.alignment = aln(h="center")

        # Notes
        ws.merge_cells(start_row=row, start_column=7, end_row=row, end_column=10)
        inp(ws, row, 7, bg=bg)

    # Instructions
    spacer(ws, GS_STUDENT_ROW + NUM_STUDENTS)   # row after last student
    r = GS_STUDENT_ROW + NUM_STUDENTS + 1
    sub_banner(ws, r, 1, 10, "QUICK START GUIDE", bg=WARM_200, size=10, height=20)
    steps = [
        "1.  Set semester dates above (Sem 1 and Sem 2) — week counts auto-calculate.",
        "2.  Enter each student's name, DOB, and graduation year — grade fills automatically.",
        "3.  Go to the  Course Content  tab to assign curricula to each student.",
        "4.  Each student's individual schedule tab will auto-populate from Course Content.",
    ]
    for i, step in enumerate(steps):
        rr = r + 1 + i
        ws.row_dimensions[rr].height = 18
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=10)
        sc(ws.cell(rr, 1), value=step, size=10, color=WARM_700,
           bg=OFF_WHITE if i % 2 else WHITE)


# ── COURSE CONTENT ────────────────────────────────────────────────────────────
# Col layout (10 cols):
# A(26): Subject name (or elective custom-name input)
# B(30): Curriculum 1 title   C(11): Unit type   D(9): Sem 1   E(9): Sem 2
# F(3):  divider
# G(30): Curriculum 2 title   H(11): Unit type   I(9): Sem 1   J(9): Sem 2
CC_WIDTHS = [26, 30, 11, 9, 9, 3, 30, 11, 9, 9]

# Rows per student block: 1 header + 14 subject rows + 1 spacer = 16
CC_ROWS_PER_STUDENT = 16
CC_GROUP_HDR_ROW    = 7
CC_GROUP_DATA_START = 8    # rows 8–21 = 14 group-studies subject rows
CC_STUDENT_FIRST    = 23   # Student 1 header row (data rows 24–37, spacer 38)

def cc_student_hdr_row(si):   # 0-based student index
    return CC_STUDENT_FIRST + si * CC_ROWS_PER_STUDENT

def cc_student_data_row(si, sj):   # subject j for student i
    return cc_student_hdr_row(si) + 1 + sj


def _cc_col_headers(ws, row):
    """Draw the two-level column header at the given row (uses row and row+1)."""
    ws.row_dimensions[row].height = 18
    ws.row_dimensions[row + 1].height = 24

    # Row 1: slot labels spanning their columns
    ws.merge_cells(start_row=row, start_column=1, end_row=row+1, end_column=1)
    sc(ws.cell(row, 1), value="Subject", size=9, bold=True, color=WARM_700,
       bg=WARM_200, h="center", v="center", b=box(color=WARM_300))

    for c1, c2, label, bg_col in [
        (2, 5, "── Curriculum Slot 1 ──", S1_LIGHT),
        (7,10, "── Curriculum Slot 2 ──", S2_LIGHT),
    ]:
        ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
        sc(ws.cell(row, c1), value=label, size=9, bold=True, color=WARM_700,
           bg=bg_col, h="center", b=box(color=WARM_300))

    ws.cell(row, 6).fill = fill(OFF_WHITE)   # divider top half

    # Row 2: sub-column labels
    sub = [
        (2, "Title / Book"),
        (3, "Unit\nType"),
        (4, "Sem 1\nUnits"),
        (5, "Sem 2\nUnits"),
        (7, "Title / Book"),
        (8, "Unit\nType"),
        (9, "Sem 1\nUnits"),
        (10,"Sem 2\nUnits"),
    ]
    for col, txt in sub:
        bg = S1_LIGHT if col <= 5 else S2_LIGHT
        sc(ws.cell(row+1, col), value=txt, size=9, bold=True, color=WARM_700,
           bg=bg, h="center", v="center", wrap=True, b=box(color=WARM_300))

    ws.cell(row+1, 6).fill = fill(OFF_WHITE)   # divider bottom half


def _cc_subject_row(ws, row, subj_name, editable=False, outline_level=0,
                    hidden=False, dv_unit_type=None):
    """Draw one subject data row."""
    ws.row_dimensions[row].height   = 20
    ws.row_dimensions[row].outlineLevel = outline_level
    if hidden:
        ws.row_dimensions[row].hidden = True

    alt = OFF_WHITE if (row % 2 == 0) else WHITE

    # Col A: subject label (or editable input for electives)
    c = ws.cell(row, 1)
    if editable:
        c.fill      = fill(WHITE)
        c.font      = fnt(size=10, italic=True, color=WARM_500)
        if not c.value:
            c.value = subj_name   # default placeholder
    else:
        c.value     = subj_name
        c.fill      = fill(LINEN)
        c.font      = fnt(size=10, bold=True, color=WARM_700)
    c.border    = left_accent()
    c.alignment = aln(h="left", v="center")

    # Slot 1: cols B–E
    inp(ws, row, 2, bg=WHITE)                           # curriculum title
    c3 = inp(ws, row, 3, value="pages", bg=alt); c3.alignment = aln(h="center")
    c4 = inp(ws, row, 4, bg=alt);  c4.alignment = aln(h="center")
    c5 = inp(ws, row, 5, bg=alt);  c5.alignment = aln(h="center")

    # Divider col F
    ws.cell(row, 6).fill = fill(OFF_WHITE)

    # Slot 2: cols G–J
    inp(ws, row, 7, bg=WHITE)                           # curriculum title
    c8  = inp(ws, row, 8, value="pages", bg=alt); c8.alignment  = aln(h="center")
    c9  = inp(ws, row, 9, bg=alt);  c9.alignment  = aln(h="center")
    c10 = inp(ws, row,10, bg=alt);  c10.alignment = aln(h="center")

    if dv_unit_type:
        dv_unit_type.add(ws.cell(row, 3))
        dv_unit_type.add(ws.cell(row, 8))


def build_course_content(wb):
    ws = wb.create_sheet("Course Content", 1)
    set_widths(ws, CC_WIDTHS)

    # ── Header ────────────────────────────────────────────────────────────────
    title_bar(ws, 1, 1, 10, "Course Content  ·  Curriculum Planning Hub", height=40)
    sub_banner(ws, 2, 1, 10,
               "Each subject has two curriculum slots.  "
               "Group Studies apply to all students.  "
               "Use the  ＋ / −  buttons on the left to expand or collapse each student.",
               bg=WARM_200, size=10, height=20)
    spacer(ws, 3)

    # Tip row
    ws.row_dimensions[4].height = 16
    ws.merge_cells("A4:J4")
    sc(ws.cell(4, 1),
       value="  Tip: click the  −  button beside a student's name to collapse their rows.  "
             "Expand only the student you're currently planning for.",
       size=9, italic=True, color=WARM_500, bg=OFF_WHITE)

    spacer(ws, 5)

    # Column headers (rows 5–6, but we'll use 5=slot labels, 6=sub-cols)
    _cc_col_headers(ws, 5)

    spacer(ws, 7, height=6)  # tiny spacer before Group section, reused as row 7

    # Unit type dropdown — shared across all subject rows
    dv_unit = DataValidation(
        type="list",
        formula1='"pages,lessons,chapters,units,assignments,projects"',
        showDropDown=False,
        showErrorMessage=False,
    )
    ws.add_data_validation(dv_unit)

    # ── Group Studies section ─────────────────────────────────────────────────
    ws.row_dimensions[CC_GROUP_HDR_ROW].height = 26
    ws.merge_cells(f"A{CC_GROUP_HDR_ROW}:J{CC_GROUP_HDR_ROW}")
    c = ws.cell(CC_GROUP_HDR_ROW, 1)
    sc(c, value="  GROUP STUDIES  ·  Applies to All Students",
       size=12, bold=True, color=WHITE, bg=WARM_700,
       h="left", v="center", b=box(color=WARM_300))

    for sj, (subj_name, editable) in enumerate(ALL_SUBJECTS):
        _cc_subject_row(ws, CC_GROUP_DATA_START + sj, subj_name, editable,
                        outline_level=0, dv_unit_type=dv_unit)

    spacer(ws, CC_GROUP_DATA_START + NUM_SUBJECTS)   # row 22

    # ── Per-student sections ──────────────────────────────────────────────────
    for si in range(NUM_STUDENTS):
        hdr_row  = cc_student_hdr_row(si)
        gs_row   = GS_STUDENT_ROW + si
        n        = si + 1
        spc_row  = hdr_row + NUM_SUBJECTS + 1

        # Student header row (outline level 0 — always visible)
        ws.row_dimensions[hdr_row].height = 26
        ws.merge_cells(f"A{hdr_row}:J{hdr_row}")
        c = ws.cell(hdr_row, 1)
        c.value = (
            f"=IF('Getting Started'!B{gs_row}<>\"\","
            f"\"  \"&UPPER('Getting Started'!B{gs_row})"
            f"&\"   ·   \"&'Getting Started'!F{gs_row},"
            f"\"  STUDENT {n}\")"
        )
        hdr_fill, _ = STUDENT_COLORS[si]
        c.font      = fnt(size=11, bold=True, color=WHITE)
        c.fill      = fill(hdr_fill)
        c.border    = box(color=WARM_300)
        c.alignment = aln(h="left", v="center")

        # Subject rows — collapse all but student 1 by default
        collapsed = (si > 0)
        for sj, (subj_name, editable) in enumerate(ALL_SUBJECTS):
            _cc_subject_row(ws, cc_student_data_row(si, sj), subj_name,
                            editable, outline_level=1, hidden=collapsed,
                            dv_unit_type=dv_unit)

        # Spacer inside group (also collapsed with student)
        ws.row_dimensions[spc_row].height = 6
        ws.row_dimensions[spc_row].outlineLevel = 1
        if collapsed:
            ws.row_dimensions[spc_row].hidden = True

    # Freeze subject column and top header rows
    ws.freeze_panes = f"B{CC_GROUP_HDR_ROW}"

    # Sheet outline settings: show outline symbols
    ws.sheet_properties.outlinePr.summaryBelow = False

    return ws


# ── STUDENT SHEET ─────────────────────────────────────────────────────────────
# A(24): Subject  B(28): Curriculum  C(10): S1 Total  D(10): S1/Wk  E(10): S1/Day
# F(4):  divider
# G(10): S2 Total  H(10): S2/Wk  I(10): S2/Day  J(12): Unit Type
ST_WIDTHS = [20, 22, 4, 4, 4, 4, 4, 4, 6, 2, 9, 9, 9, 2, 9, 9, 9, 10]
# 18 cols: A=Subject B=Curriculum C-H=Mon-Sat I=Days/Wk J=÷
#          K=S1Total L=S1/Wk M=S1/Day N=÷ O=S2Total P=S2/Wk Q=S2/Day R=UnitType
ST_NUM_COLS = 18

# CC_COLS — Slot 1 columns in Course Content
CC_COLS = {
    "custom_name": 1,   # col A — subject label / elective name
    "curriculum":  2,   # col B — Slot 1 curriculum title
    "unit_type":   3,   # col C — Slot 1 unit type
    "s1_units":    4,   # col D — Slot 1 Sem 1 units
    "s2_units":    5,   # col E — Slot 1 Sem 2 units
}

def cc_ref(student_idx, subject_idx, col_name):
    r   = cc_student_data_row(student_idx, subject_idx)
    col = CC_COLS[col_name]
    return f"'Course Content'!{get_column_letter(col)}{r}"

def build_student_sheet(wb, idx):
    n            = idx + 1
    gs_name_row  = GS_STUDENT_ROW + idx
    stu_color, _ = STUDENT_COLORS[idx]

    ws = wb.create_sheet(f"Student {n}")
    set_widths(ws, ST_WIDTHS)

    # ── Banner ────────────────────────────────────────────────────────────────
    ws.merge_cells("A1:R1")
    c = ws["A1"]
    c.value = (
        f"=IF('Getting Started'!B{gs_name_row}=\"\","
        f"\"Student {n}\",'Getting Started'!B{gs_name_row})"
    )
    sc(c, bold=True, size=22, color=WHITE, bg=stu_color, h="center", v="center")
    ws.row_dimensions[1].height = 44

    ws.merge_cells("A2:R2")
    c = ws["A2"]
    c.value = (
        f"=IFERROR(\"Grade: \"&'Getting Started'!F{gs_name_row}"
        f"&\"   ·   School Year: \"&'Getting Started'!C6,\"\")"
    )
    sc(c, size=11, color=WHITE, bg=WARM_700, h="center", v="center")
    ws.row_dimensions[2].height = 22
    spacer(ws, 3)

    # ── Semester overview strip (row 4) ───────────────────────────────────────
    ws.row_dimensions[4].height = 20
    ws.merge_cells("A4:I4")
    ws["A4"].value = (
        f'=IFERROR("SEM 1   "'
        f'&TEXT({GS_S1_START},"MMM D")&" – "'
        f'&TEXT({GS_S1_END},"MMM D, YYYY")'
        f'&"   ("&{GS_S1_WEEKS}&" school wks)","SEMESTER 1")'
    )
    sc(ws["A4"], size=10, bold=True, color=WHITE, bg=S1_HDR, h="center", b=box())

    ws.cell(4, 10).fill = fill(OFF_WHITE)   # col J divider

    ws.merge_cells("K4:R4")
    ws["K4"].value = (
        f'=IFERROR("SEM 2   "'
        f'&TEXT({GS_S2_START},"MMM D")&" – "'
        f'&TEXT({GS_S2_END},"MMM D, YYYY")'
        f'&"   ("&{GS_S2_WEEKS}&" school wks)","SEMESTER 2")'
    )
    sc(ws["K4"], size=10, bold=True, color=WHITE, bg=S2_HDR, h="center", b=box())

    spacer(ws, 5)

    # ── School Days selector (rows 6–9) ───────────────────────────────────────
    ws.row_dimensions[6].height = 20
    ws.merge_cells("A6:R6")
    sc(ws["A6"], value="  SCHOOL DAYS  ·  Mark each day this student attends",
       size=10, bold=True, color=WARM_700, bg=LINEN, h="left", b=box())

    # Day labels row 7 — cols C-H match schedule table day-marker columns
    ws.row_dimensions[7].height = 18
    days = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
    day_cols = [3, 4, 5, 6, 7, 8]   # cols C–H
    sc(ws.cell(7, 1), value="", bg=WARM_200, b=box())
    sc(ws.cell(7, 2), value="Typical school week  →", size=9, italic=True,
       color=WARM_700, bg=WARM_200, h="right", b=box())
    for day, col in zip(days, day_cols):
        sc(ws.cell(7, col), value=day, size=9, bold=True,
           color=WARM_700, bg=WARM_200, h="center", b=box())
    sc(ws.cell(7, 9), value="Days /Wk", size=9, bold=True,
       color=WARM_700, bg=WARM_200, h="center", b=box())
    ws.cell(7, 10).fill = fill(OFF_WHITE)
    for col in range(11, ST_NUM_COLS + 1):
        ws.cell(7, col).fill = fill(WARM_200)

    # Input row 8 — users type X in C-H for their school week
    ws.row_dimensions[8].height = 22
    ws.cell(8, 1).fill = fill(OFF_WHITE)
    ws.cell(8, 2).fill = fill(OFF_WHITE)
    for col in day_cols:
        c = ws.cell(8, col)
        c.fill      = fill(WHITE)
        c.border    = box(color=ACCENT)
        c.font      = fnt(size=14, bold=True, color=stu_color)
        c.alignment = aln(h="center", v="center")

    c = ws.cell(8, 9)   # Days/week total for this student's school week
    c.value     = "=MAX(1,COUNTA(C8:H8))"
    c.fill      = fill(WARM_100)
    c.border    = box(color=WARM_300)
    c.font      = fnt(size=14, bold=True, color=WARM_700)
    c.alignment = aln(h="center", v="center")

    ws.cell(8, 10).fill = fill(OFF_WHITE)
    for col in range(11, ST_NUM_COLS + 1):
        ws.cell(8, col).fill = fill(OFF_WHITE)

    spacer(ws, 9)

    # ── Schedule table column headers (row 10) ────────────────────────────────
    ws.row_dimensions[10].height = 28
    hdr_defs = [
        (1,  "Subject",        WARM_200, WARM_700),
        (2,  "Curriculum / Book", WARM_200, WARM_700),
        (3,  "M",              WARM_200, WARM_700),
        (4,  "T",              WARM_200, WARM_700),
        (5,  "W",              WARM_200, WARM_700),
        (6,  "Th",             WARM_200, WARM_700),
        (7,  "F",              WARM_200, WARM_700),
        (8,  "Sa",             WARM_200, WARM_700),
        (9,  "Days\n/Wk",     WARM_200, WARM_700),
        (10, "",               OFF_WHITE, OFF_WHITE),
        (11, "S1\nTotal",     S1_LIGHT,  S1_HDR),
        (12, "S1\n/Wk",      S1_LIGHT,  S1_HDR),
        (13, "S1\n/Day",     S1_LIGHT,  S1_HDR),
        (14, "",               OFF_WHITE, OFF_WHITE),
        (15, "S2\nTotal",    S2_LIGHT,  S2_HDR),
        (16, "S2\n/Wk",     S2_LIGHT,  S2_HDR),
        (17, "S2\n/Day",    S2_LIGHT,  S2_HDR),
        (18, "Unit\nType",   WARM_200,  WARM_700),
    ]
    for col, txt, bg_col, fc in hdr_defs:
        c = ws.cell(10, col)
        if col in (10, 14):
            c.fill = fill(OFF_WHITE)
        else:
            sc(c, value=txt, size=9, bold=True, color=fc, bg=bg_col,
               h="center", v="center", wrap=True, b=box())

    # ── Subject rows ──────────────────────────────────────────────────────────
    for sj, (subj_name, editable) in enumerate(ALL_SUBJECTS):
        row = 11 + sj
        ws.row_dimensions[row].height = 22
        alt = WHITE if sj % 2 == 0 else OFF_WHITE

        s1_ref  = cc_ref(idx, sj, "s1_units")
        s2_ref  = cc_ref(idx, sj, "s2_units")
        cur_ref = cc_ref(idx, sj, "curriculum")
        ut_ref  = cc_ref(idx, sj, "unit_type")
        cst_ref = cc_ref(idx, sj, "custom_name")

        # Col A: subject label
        c = ws.cell(row, 1)
        if editable:
            c.value = f'=IF({cst_ref}="{subj_name}","{subj_name}",{cst_ref})'
        else:
            c.value = subj_name
        c.fill      = fill(LINEN)
        c.border    = left_accent()
        c.font      = fnt(size=10, bold=True, color=WARM_700)
        c.alignment = aln(h="left", v="center")

        # Col B: curriculum title
        c = ws.cell(row, 2)
        c.value     = f'=IFERROR(IF({cur_ref}="","—",{cur_ref}),"")'
        c.fill      = fill(alt)
        c.border    = box()
        c.font      = fnt(size=10, color=WARM_900)
        c.alignment = aln(h="left", v="center")

        # Cols C-H: per-subject day markers (user types X)
        for col in range(3, 9):
            c = ws.cell(row, col)
            c.fill      = fill(WHITE)
            c.border    = box(color=WARM_300)
            c.font      = fnt(size=11, bold=True, color=stu_color)
            c.alignment = aln(h="center", v="center")

        # Col I: per-subject Days/Wk — defaults to student school-week if blank
        days_cell = f"I{row}"
        c = ws.cell(row, 9)
        c.value     = f"=IF(COUNTA(C{row}:H{row})=0,$I$8,COUNTA(C{row}:H{row}))"
        c.fill      = fill(WARM_100)
        c.border    = box()
        c.font      = fnt(size=10, bold=True, color=WARM_700)
        c.alignment = aln(h="center")

        # Col J: divider
        ws.cell(row, 10).fill = fill(OFF_WHITE)

        # Cols K-M: Sem 1
        s1_l = f"L{row}"
        c = ws.cell(row, 11)   # K: S1 Total
        c.value     = f"=IFERROR(IF({s1_ref}=\"\",\"—\",{s1_ref}),\"\")"
        c.fill      = fill(S1_LIGHT); c.border = box()
        c.font      = fnt(size=10, color=S1_HDR, bold=True)
        c.alignment = aln(h="center")

        c = ws.cell(row, 12)   # L: S1/Wk (floor — no decimals)
        c.value     = f'=IFERROR(IF({s1_ref}="","",FLOOR({s1_ref}/{GS_S1_WEEKS},1)),"")'
        c.fill      = fill(S1_LIGHT); c.border = box()
        c.font      = fnt(size=10, color=S1_HDR, bold=True)
        c.alignment = aln(h="center")

        c = ws.cell(row, 13)   # M: S1/Day (floor — no decimals)
        c.value     = f'=IFERROR(IF({s1_l}="","",FLOOR({s1_l}/{days_cell},1)),"")'
        c.fill      = fill(S1_LIGHT); c.border = box()
        c.font      = fnt(size=11, bold=True, color=S1_HDR)
        c.alignment = aln(h="center")

        # Col N: divider
        ws.cell(row, 14).fill = fill(OFF_WHITE)

        # Cols O-Q: Sem 2
        s2_p = f"P{row}"
        c = ws.cell(row, 15)   # O: S2 Total
        c.value     = f"=IFERROR(IF({s2_ref}=\"\",\"—\",{s2_ref}),\"\")"
        c.fill      = fill(S2_LIGHT); c.border = box()
        c.font      = fnt(size=10, color=S2_HDR, bold=True)
        c.alignment = aln(h="center")

        c = ws.cell(row, 16)   # P: S2/Wk (floor — no decimals)
        c.value     = f'=IFERROR(IF({s2_ref}="","",FLOOR({s2_ref}/{GS_S2_WEEKS},1)),"")'
        c.fill      = fill(S2_LIGHT); c.border = box()
        c.font      = fnt(size=10, color=S2_HDR, bold=True)
        c.alignment = aln(h="center")

        c = ws.cell(row, 17)   # Q: S2/Day (floor — no decimals)
        c.value     = f'=IFERROR(IF({s2_p}="","",FLOOR({s2_p}/{days_cell},1)),"")'
        c.fill      = fill(S2_LIGHT); c.border = box()
        c.font      = fnt(size=11, bold=True, color=S2_HDR)
        c.alignment = aln(h="center")

        # Col R: unit type
        c = ws.cell(row, 18)
        c.value     = f'=IFERROR(IF({ut_ref}="","",{ut_ref}),"")'
        c.fill      = fill(alt); c.border = box()
        c.font      = fnt(size=9, italic=True, color=WARM_500)
        c.alignment = aln(h="center")

    ws.freeze_panes = "A11"
    return ws



# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    build_getting_started(wb)
    build_course_content(wb)

    # Add Student 1 tab for now; remaining tabs added once design is finalised
    ws1 = build_student_sheet(wb, 0)
    ws1.sheet_properties.tabColor = STUDENT_COLORS[0][0]

    wb["Getting Started"].sheet_properties.tabColor = ACCENT
    wb["Course Content"].sheet_properties.tabColor  = WARM_700

    out = "/home/user/curriculum-planning/Homeschool_Curriculum_Planner.xlsx"
    wb.save(out)
    print(f"Saved → {out}")
    print(f"Sheets: {wb.sheetnames}")


if __name__ == "__main__":
    main()
