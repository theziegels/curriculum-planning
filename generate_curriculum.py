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
        (7, 9, "Curriculum Notes"),
        (10,10,""),
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
            f'=IF(OR(B{row}="",E{row}=""),"",LET(g,12-(E{row}-YEAR($C$19)),'
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
    spacer(ws, 33 + NUM_STUDENTS - 10)
    r = 40
    sub_banner(ws, r, 1, 10, "QUICK START GUIDE", bg=WARM_200, size=10, height=20)
    steps = [
        "1.  Set semester dates above (Sem 1 and Sem 2) — week counts auto-calculate.",
        "2.  Enter each student's name, DOB, and graduation year — grade fills automatically.",
        "3.  Go to the  Course Content  tab to assign curricula to each student.",
        "4.  Each student's tab auto-populates pace calculations from Course Content.",
        "5.  Yellow cells = type here.   Cream/tinted cells = formulas, leave them alone.",
    ]
    for i, step in enumerate(steps):
        rr = r + 1 + i
        ws.row_dimensions[rr].height = 18
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=10)
        sc(ws.cell(rr, 1), value=step, size=10, color=WARM_700,
           bg=OFF_WHITE if i % 2 else WHITE)


# ── COURSE CONTENT ────────────────────────────────────────────────────────────
# Column layout (10 cols):
# A(5): Student #   B(20): Student Name (formula)   C(28): Subject / Course
# D(16): Custom Name (electives)   E(32): Curriculum / Book Title
# F(12): Unit Type   G(11): Sem 1 Units   H(11): Sem 2 Units   I(30): Notes
CC_WIDTHS = [5, 20, 28, 18, 34, 12, 11, 11, 32, 4]
CC_COLS   = {"student_num":1,"student_name":2,"subject":3,"custom_name":4,
             "curriculum":5,"unit_type":6,"s1_units":7,"s2_units":8,"notes":9}

def build_course_content(wb):
    ws = wb.create_sheet("Course Content", 1)
    set_widths(ws, CC_WIDTHS)

    title_bar(ws, 1, 1, 9, "Course Content  ·  Curriculum Planning Hub", height=40)
    sub_banner(ws, 2, 1, 9,
               "Enter each student's curricula here. "
               "Use the  View Student  filter below to focus on one student at a time.",
               bg=WARM_200, size=10, height=20)
    spacer(ws, 3)

    # View Student control (row 4–5)
    ws.row_dimensions[4].height = 22
    ws.merge_cells("A4:B4")
    sc(ws.cell(4,1), value="View / Filter Student:", size=10, bold=True,
       color=WARM_700, bg=ACCENT_PALE, h="right", b=box(color=ACCENT))

    c_view = ws.cell(4, 3)
    c_view.value = 1
    c_view.fill  = fill(WHITE)
    c_view.border= box(color=ACCENT)
    c_view.font  = fnt(size=12, bold=True, color=ACCENT)
    c_view.alignment = aln(h="center")
    dv_view = DataValidation(type="whole", operator="between",
                             formula1="1", formula2="10",
                             showErrorMessage=True,
                             errorTitle="Invalid", error="Enter 1–10")
    ws.add_data_validation(dv_view)
    dv_view.add(ws["C4"])

    ws.merge_cells("D4:E4")
    ws["D4"].value = (
        f"=IFERROR(\"→  \"&INDEX({{\"Student 1\",\"Student 2\",\"Student 3\","
        f"\"Student 4\",\"Student 5\",\"Student 6\",\"Student 7\","
        f"\"Student 8\",\"Student 9\",\"Student 10\"}},C4),\"\")"
    )
    # simpler: just show name from GS
    ws["D4"].value = (
        '=IFERROR("→  "&CHOOSE(C4,'
        + ",".join(f"'Getting Started'!B{GS_STUDENT_ROW+i}" for i in range(10))
        + '),"→")'
    )
    ws["D4"].fill      = fill(ACCENT_PALE)
    ws["D4"].border    = box(color=ACCENT)
    ws["D4"].font      = fnt(size=11, bold=True, color=WARM_700)
    ws["D4"].alignment = aln(h="left", v="center")

    ws.merge_cells("F4:I4")
    sc(ws.cell(4,6),
       value="Rows highlighted in pink = selected student's courses.",
       size=9, italic=True, color=WARM_500, bg=ACCENT_PALE)

    spacer(ws, 5)

    # Hint row 6
    ws.row_dimensions[6].height = 16
    ws.merge_cells("A6:I6")
    sc(ws.cell(6,1),
       value="  Tip: Use the column filter arrows (▼) to show only one student's rows."
             "  Yellow cells = type here.  All other cells are locked formulas.",
       size=9, italic=True, color=WARM_500, bg=OFF_WHITE)

    # Column headers row 7
    ws.row_dimensions[CC_HDR_ROW].height = 28
    col_hdrs = [
        (1, "Std\n#"),
        (2, "Student Name"),
        (3, "Subject"),
        (4, "Custom Name\n(Electives)"),
        (5, "Curriculum / Book Title"),
        (6, "Unit\nType"),
        (7, "Semester 1\nUnits"),
        (8, "Semester 2\nUnits"),
        (9, "Notes"),
    ]
    for col, txt in col_hdrs:
        c = ws.cell(CC_HDR_ROW, col)
        sc(c, value=txt, size=9, bold=True, color=WARM_700, bg=WARM_200,
           h="center", v="center", wrap=True, b=box(color=WARM_300))

    # Data rows
    subject_names = [s[0] for s in ALL_SUBJECTS]
    dv_std = DataValidation(type="whole", operator="between",
                            formula1="1", formula2="10")
    ws.add_data_validation(dv_std)

    for si in range(NUM_STUDENTS):
        for sj, (subj_name, editable) in enumerate(ALL_SUBJECTS):
            row = CC_DATA_START + si * NUM_SUBJECTS + sj
            ws.row_dimensions[row].height = 20
            bg = WHITE if sj % 2 == 0 else OFF_WHITE

            # Col A: Student number (pre-filled, editable)
            c = ws.cell(row, 1)
            c.value     = si + 1
            c.fill      = fill(bg)
            c.border    = box(color=WARM_300)
            c.font      = fnt(size=10, bold=True, color=WARM_700)
            c.alignment = aln(h="center")
            dv_std.add(c)

            # Col B: Student name (formula)
            c = ws.cell(row, 2)
            c.value     = f"='Getting Started'!B{GS_STUDENT_ROW + si}"
            c.fill      = fill(WARM_100)
            c.border    = box(color=WARM_300)
            c.font      = fnt(size=10, color=WARM_500)
            c.alignment = aln(h="left")

            # Col C: Subject (pre-filled, editable)
            c = ws.cell(row, 3)
            c.value     = subj_name
            c.fill      = fill(bg)
            c.border    = box(color=WARM_300)
            c.font      = fnt(size=10, bold=not editable, color=WARM_700)
            c.alignment = aln(h="left")

            # Col D: Custom name (electives only)
            c = ws.cell(row, 4)
            if editable:
                c.fill   = fill(WHITE)
                c.font   = fnt(size=10, color=WARM_900)
            else:
                c.fill   = fill(WARM_100)
                c.font   = fnt(size=9, italic=True, color=WARM_500)
                c.value  = "—"
            c.border    = box(color=WARM_300)
            c.alignment = aln(h="left")

            # Col E: Curriculum title — user input
            inp(ws, row, 5, bg=WHITE)

            # Col F: Unit type — input
            c = inp(ws, row, 6, value="pages", bg=WHITE)
            c.alignment = aln(h="center")

            # Col G: Sem 1 units — input
            c = inp(ws, row, 7, bg=WHITE)
            c.alignment = aln(h="center")

            # Col H: Sem 2 units — input
            c = inp(ws, row, 8, bg=WHITE)
            c.alignment = aln(h="center")

            # Col I: Notes — input
            inp(ws, row, 9, bg=WHITE)

    # Conditional formatting: highlight selected student's rows
    last_data_row = CC_DATA_START + NUM_STUDENTS * NUM_SUBJECTS - 1
    data_range    = f"A{CC_DATA_START}:I{last_data_row}"
    ws.conditional_formatting.add(
        data_range,
        FormulaRule(
            formula=[f"$A{CC_DATA_START}=$C$4"],
            fill=PatternFill(start_color=ACCENT_PALE, end_color=ACCENT_PALE,
                             fill_type="solid"),
        )
    )

    # AutoFilter on header row
    ws.auto_filter.ref = f"A{CC_HDR_ROW}:I{last_data_row}"

    # Freeze panes: freeze rows 1-7 and col A
    ws.freeze_panes = f"B{CC_DATA_START}"

    return ws


# ── STUDENT SHEET ─────────────────────────────────────────────────────────────
# Col layout: A(28) B(14) C(14) D(14) E(18) F(4) G(14) H(14) I(14) J(18)
ST_WIDTHS = [28, 14, 14, 14, 18, 4, 14, 14, 14, 18]

def cc_row(student_idx, subject_idx):
    """Course Content data row for 0-based student and subject indices."""
    return CC_DATA_START + student_idx * NUM_SUBJECTS + subject_idx

def cc_ref(student_idx, subject_idx, col_name):
    r   = cc_row(student_idx, subject_idx)
    col = CC_COLS[col_name]
    return f"'Course Content'!{get_column_letter(col)}{r}"

def build_student_sheet(wb, idx):
    n  = idx + 1
    gs_name_row = GS_STUDENT_ROW + idx

    ws = wb.create_sheet(f"Student {n}")
    set_widths(ws, ST_WIDTHS)

    # Banner
    title_bar(ws, 1, 1, 10,
              f"='Getting Started'!B{gs_name_row}",
              height=40)
    ws.cell(1, 1).value = (
        f"=IF('Getting Started'!B{gs_name_row}=\"\","
        f"\"Student {n}\","
        f"'Getting Started'!B{gs_name_row})"
    )

    ws.row_dimensions[2].height = 22
    ws.merge_cells("A2:J2")
    ws["A2"].value = (
        f"=IFERROR("
        f"\"Grade: \"&'Getting Started'!F{gs_name_row}"
        f"&\"   ·   School Year: \"&'Getting Started'!C6,\"\")"
    )
    sc(ws["A2"], size=11, bold=False, color=WHITE, bg=WARM_700, h="center", v="center")

    spacer(ws, 3)

    # Semester overview strip
    ws.row_dimensions[4].height = 22
    ws.merge_cells("A4:E4")
    ws["A4"].value = (
        f'=IFERROR("SEMESTER 1   "'
        f'&TEXT({GS_S1_START},"MMM D")&" – "'
        f'&TEXT({GS_S1_END},"MMM D, YYYY")'
        f'&"   ("&{GS_S1_WEEKS}&" wks)","SEMESTER 1")'
    )
    sc(ws["A4"], size=10, bold=True, color=WHITE, bg=S1_HDR, h="center", b=box(color=WARM_300))

    ws.cell(4, 6).fill = fill(OFF_WHITE)   # divider col

    ws.merge_cells("G4:J4")
    ws["G4"].value = (
        f'=IFERROR("SEMESTER 2   "'
        f'&TEXT({GS_S2_START},"MMM D")&" – "'
        f'&TEXT({GS_S2_END},"MMM D, YYYY")'
        f'&"   ("&{GS_S2_WEEKS}&" wks)","SEMESTER 2")'
    )
    sc(ws["G4"], size=10, bold=True, color=WHITE, bg=S2_HDR, h="center", b=box(color=WARM_300))

    # Column sub-headers row 5
    ws.row_dimensions[5].height = 24
    sub_cols = [
        (1, 1, "Subject  /  Curriculum",        WARM_200, WARM_700),
        (2, 2, "Sem 1\nUnits",                  S1_LIGHT, S1_HDR),
        (3, 3, "/ Week",                         S1_LIGHT, S1_HDR),
        (4, 4, "/ Day",                          S1_LIGHT, S1_HDR),
        (5, 5, "Est. End",                        S1_LIGHT, S1_HDR),
        (6, 6, "",                                OFF_WHITE, OFF_WHITE),
        (7, 7, "Sem 2\nUnits",                  S2_LIGHT, S2_HDR),
        (8, 8, "/ Week",                         S2_LIGHT, S2_HDR),
        (9, 9, "/ Day",                          S2_LIGHT, S2_HDR),
        (10,10,"Est. End",                        S2_LIGHT, S2_HDR),
    ]
    for c1, c2, txt, bg, fc in sub_cols:
        if c1 != c2:
            ws.merge_cells(start_row=5, start_column=c1, end_row=5, end_column=c2)
        c = ws.cell(5, c1)
        sc(c, value=txt, size=9, bold=True, color=fc, bg=bg,
           h="center", v="center", wrap=True, b=box(color=WARM_300))

    spacer(ws, 6)

    # Subject sections
    row = 7
    for j, (subj_name, editable) in enumerate(ALL_SUBJECTS):
        row = add_subject_section(ws, row, subj_name, editable, idx, j)
        spacer(ws, row); row += 1

    ws.freeze_panes = "A7"
    return ws


def add_subject_section(ws, r0, subj_name, editable, student_idx, subject_idx):
    """
    r0+0  Subject header  (with curriculum title pulled from CC)
    r0+1  Pace strip: S1 units | /wk | /day | est.end  |  S2 units | /wk | /day | est.end
    r0+2  Component breakdown sub-header
    r0+3  Component column headers
    r0+4..r0+9  6 component rows
    r0+10 Notes
    Returns next row.
    """
    # Short references to Course Content cells for this student+subject
    cur_ref  = cc_ref(student_idx, subject_idx, "curriculum")
    cust_ref = cc_ref(student_idx, subject_idx, "custom_name")
    ut_ref   = cc_ref(student_idx, subject_idx, "unit_type")
    s1_ref   = cc_ref(student_idx, subject_idx, "s1_units")
    s2_ref   = cc_ref(student_idx, subject_idx, "s2_units")
    notes_ref= cc_ref(student_idx, subject_idx, "notes")

    # ── Row 0: subject header ──────────────────────────────────────────────
    ws.row_dimensions[r0].height = 24

    if editable:
        # Elective: show custom name if provided, else show default
        ws.merge_cells(start_row=r0, start_column=1, end_row=r0, end_column=4)
        c = ws.cell(r0, 1)
        c.value = (
            f'=IF({cust_ref}="","  {subj_name.upper()}",'
            f'"  "&UPPER({cust_ref}))'
        )
    else:
        ws.merge_cells(start_row=r0, start_column=1, end_row=r0, end_column=4)
        c = ws.cell(r0, 1)
        c.value = f"  {subj_name.upper()}"

    c.font      = fnt(size=12, bold=True, color=WARM_700)
    c.fill      = fill(LINEN)
    c.border    = left_accent()
    c.alignment = aln(h="left", v="center")

    # Curriculum title (from CC) in cols 5-10
    ws.merge_cells(start_row=r0, start_column=5, end_row=r0, end_column=10)
    c = ws.cell(r0, 5)
    c.value = (
        f'=IFERROR(IF({cur_ref}="","← enter on Course Content tab",{cur_ref}),"")'
    )
    c.fill      = fill(WARM_100)
    c.border    = box(color=WARM_300)
    c.font      = fnt(size=11, bold=True, color=WARM_700)
    c.alignment = aln(h="left", v="center")

    # ── Row 1: pace strip ──────────────────────────────────────────────────
    r1 = r0 + 1
    ws.row_dimensions[r1].height = 22

    s1u = f"B{r1}";  s1w = f"C{r1}"
    s2u = f"G{r1}";  s2w = f"H{r1}"

    # S1 units (formula from CC)
    c = ws.cell(r1, 2)
    c.value = f"={s1_ref}"; c.fill = fill(S1_LIGHT)
    c.border = box(color=WARM_300)
    c.font   = fnt(size=11, bold=True, color=S1_HDR)
    c.alignment = aln(h="center")

    # S1 /week
    formula_cell(ws, r1, 3,
        f'=IFERROR(IF({s1u}="","",ROUND({s1u}/{GS_S1_WEEKS},1)),"")',
        bold=True)
    ws.cell(r1,3).fill = fill(S1_LIGHT)
    ws.cell(r1,3).font = fnt(size=10, bold=True, color=S1_HDR)

    # S1 /day
    formula_cell(ws, r1, 4,
        f'=IFERROR(IF({s1w}="","",ROUND({s1w}/{GS_DAYS},1)),"")')
    ws.cell(r1,4).fill = fill(S1_LIGHT)
    ws.cell(r1,4).font = fnt(size=10, bold=True, color=S1_HDR)

    # S1 est end
    formula_cell(ws, r1, 5,
        f'=IFERROR(IF({s1u}="","",TEXT({GS_S1_START}+({s1u}/D{r1}),"MMM D, YYYY")),"")')
    ws.cell(r1,5).fill = fill(S1_LIGHT)
    ws.cell(r1,5).font = fnt(size=9, bold=False, color=S1_HDR)

    # Divider col F
    ws.cell(r1, 6).fill = fill(OFF_WHITE)

    # S2 units
    c = ws.cell(r1, 7)
    c.value = f"={s2_ref}"; c.fill = fill(S2_LIGHT)
    c.border = box(color=WARM_300)
    c.font   = fnt(size=11, bold=True, color=S2_HDR)
    c.alignment = aln(h="center")

    # S2 /week
    formula_cell(ws, r1, 8,
        f'=IFERROR(IF({s2u}="","",ROUND({s2u}/{GS_S2_WEEKS},1)),"")')
    ws.cell(r1,8).fill = fill(S2_LIGHT)
    ws.cell(r1,8).font = fnt(size=10, bold=True, color=S2_HDR)

    # S2 /day
    formula_cell(ws, r1, 9,
        f'=IFERROR(IF({s2w}="","",ROUND({s2w}/{GS_DAYS},1)),"")')
    ws.cell(r1,9).fill = fill(S2_LIGHT)
    ws.cell(r1,9).font = fnt(size=10, bold=True, color=S2_HDR)

    # S2 est end
    formula_cell(ws, r1, 10,
        f'=IFERROR(IF({s2u}="","",TEXT({GS_S2_START}+({s2u}/I{r1}),"MMM D, YYYY")),"")')
    ws.cell(r1,10).fill = fill(S2_LIGHT)
    ws.cell(r1,10).font = fnt(size=9, bold=False, color=S2_HDR)

    # ── Row 2: component sub-header ────────────────────────────────────────
    r2 = r0 + 2
    ws.row_dimensions[r2].height = 16
    ws.merge_cells(start_row=r2, start_column=1, end_row=r2, end_column=10)
    c = ws.cell(r2, 1)
    sc(c, value="  Component Breakdown  "
                "(optional — use for multi-volume curricula)",
       size=9, italic=True, color=WARM_500, bg=OFF_WHITE,
       b=bottom_only(WARM_300))

    # ── Row 3: component col headers ──────────────────────────────────────
    r3 = r0 + 3
    ws.row_dimensions[r3].height = 22
    comp_hdrs = [
        (1, 1, "#"),
        (2, 5, "Book / Component Title"),
        (6, 6, "Units"),
        (7, 7, "Cumul."),
        (8, 8, "Start\nWeek"),
        (9, 9, "End\nWeek"),
        (10,10,"Est. Dates"),
    ]
    for c1, c2, txt in comp_hdrs:
        if c1 != c2:
            ws.merge_cells(start_row=r3, start_column=c1, end_row=r3, end_column=c2)
        c = ws.cell(r3, c1)
        sc(c, value=txt, size=9, bold=True, color=WARM_700, bg=WARM_200,
           h="center", v="center", wrap=True, b=box(color=WARM_300))

    # ── Rows 4-9: 6 component rows ─────────────────────────────────────────
    comp_start = r0 + 4

    # combined units/week for component scheduling
    total_wks = f"(IFERROR({GS_S1_WEEKS},0)+IFERROR({GS_S2_WEEKS},0))"
    s1u_ref   = f"B{r1}";  s2u_ref = f"G{r1}"
    upw_expr  = f"(IFERROR({s1u_ref},0)+IFERROR({s2u_ref},0))/{total_wks}"

    for ci in range(6):
        r = comp_start + ci
        ws.row_dimensions[r].height = 18
        bg = WHITE if ci % 2 == 0 else OFF_WHITE

        # #
        c = ws.cell(r, 1)
        sc(c, value=ci+1, size=10, bold=True, color=WARM_700,
           bg=WARM_200, h="center", b=box(color=WARM_300))

        # Title cols 2-5
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        inp(ws, r, 2, bg=bg)

        # Units col 6
        inp(ws, r, 6, bg=bg)
        ws.cell(r, 6).alignment = aln(h="center")

        # Cumulative col 7
        formula_cell(ws, r, 7,
            f'=IFERROR(SUM(F{comp_start}:F{r}),"")')

        # Start week col 8
        if ci == 0:
            sw_f = f'=IF(F{r}="","",1)'
        else:
            sw_f = f'=IFERROR(IF(F{r}="","",FLOOR(G{r-1}/({upw_expr}),1)+1),"")'
        formula_cell(ws, r, 8, sw_f)

        # End week col 9
        formula_cell(ws, r, 9,
            f'=IFERROR(IF(F{r}="","",CEILING(G{r}/({upw_expr}),1)),"")')

        # Est dates col 10 (start date – end date)
        formula_cell(ws, r, 10,
            f'=IFERROR(IF(F{r}="","",'
            f'TEXT({GS_S1_START}+(H{r}-1)*7,"MMM D")'
            f'&" – "'
            f'&TEXT({GS_S1_START}+I{r}*7,"MMM D")),"")')
        ws.cell(r,10).font = fnt(size=9, bold=False, color=WARM_700)

    # ── Notes row ──────────────────────────────────────────────────────────
    r_notes = comp_start + 6
    ws.row_dimensions[r_notes].height = 20
    ws.merge_cells(start_row=r_notes, start_column=1, end_row=r_notes, end_column=10)
    c = ws.cell(r_notes, 1)
    # Notes pulled from Course Content
    c.value = (
        f'=IFERROR(IF({notes_ref}="","",{notes_ref}),"")'
    )
    c.fill      = fill(WARM_100)
    c.border    = bottom_only(WARM_300)
    c.font      = fnt(size=9, italic=True, color=WARM_500)
    c.alignment = aln(h="left", v="center")

    return r_notes + 1


# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    build_getting_started(wb)
    build_course_content(wb)

    tab_colors = [
        "C48B8B","A87C6B","7A6E5F","6B8A6B","5C8A8A",
        "8A7A5C","8A5C5C","6B7A8A","8A6B8A","7A8A6B",
    ]
    for i in range(NUM_STUDENTS):
        ws = build_student_sheet(wb, i)
        ws.sheet_properties.tabColor = tab_colors[i]

    wb["Getting Started"].sheet_properties.tabColor  = ACCENT
    wb["Course Content"].sheet_properties.tabColor   = "A08878"

    out = "/home/user/curriculum-planning/Homeschool_Curriculum_Planner.xlsx"
    wb.save(out)
    print(f"Saved → {out}")
    print(f"Sheets: {wb.sheetnames}")


if __name__ == "__main__":
    main()
