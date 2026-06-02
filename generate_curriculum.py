#!/usr/bin/env python3
"""Generate Homeschool Curriculum Planning workbook."""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ── palette ──────────────────────────────────────────────────────────────────
NAVY        = "1F3864"
MED_BLUE    = "2E74B5"
LIGHT_BLUE  = "DEEAF1"
WHITE       = "FFFFFF"
LIGHT_GRAY  = "F2F2F2"
MID_GRAY    = "D9D9D9"
DARK_GRAY   = "595959"
INPUT_YELLOW = "FFFACD"   # light lemon – marks user-input cells
FORMULA_BG  = "EBF3FB"   # light blue – formula cells (read-only feel)

SUBJECTS = [
    # (display name, header fill, row fill)
    ("Bible",                    "5B2C6F", "F5EEF8"),
    ("Math",                     "1A5276", "D6EAF8"),
    ("English / Language Arts",  "922B21", "FDEDEC"),
    ("Reading",                  "784212", "FDEBD0"),
    ("Writing / Composition",    "7D6608", "FEFDE2"),
    ("Science",                  "0E6655", "D5F5E3"),
    ("Social Studies / History", "6E2F0E", "FAE5D3"),
    ("Health",                   "1E8449", "D4EFDF"),
    ("Physical Education",       "117A65", "D1F2EB"),
    ("Art / Music",              "76448A", "F4ECF7"),
]

NUM_STUDENTS = 10
# columns used in every sheet: A..J (10 cols)
COL_WIDTHS_GS = [4, 26, 20, 16, 14, 14, 20, 14, 14, 20]
COL_WIDTHS_ST = [4, 22, 26, 14, 14, 14, 14, 14, 14, 16]


# ── style helpers ─────────────────────────────────────────────────────────────
def fill(hex_color):
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

def font(bold=False, size=11, color="000000", italic=False, name="Calibri"):
    return Font(name=name, size=size, bold=bold, italic=italic, color=color)

def border(style="thin"):
    s = Side(style=style)
    return Border(left=s, right=s, top=s, bottom=s)

def thin_border():
    return border("thin")

def outer_border():
    s = Side(style="medium")
    return Border(left=s, right=s, top=s, bottom=s)

def align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def set_col_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

def style_cell(cell, value=None, bold=False, size=11, fcolor="000000",
               bg=None, italic=False, halign="left", valign="center",
               wrap=False, bdr=None):
    if value is not None:
        cell.value = value
    cell.font = font(bold=bold, size=size, color=fcolor)
    cell.alignment = align(h=halign, v=valign, wrap=wrap)
    if bg:
        cell.fill = fill(bg)
    if bdr:
        cell.border = bdr

def section_header(ws, row, col_start, col_end, text, hdr_fill, txt_color="FFFFFF"):
    ws.merge_cells(
        start_row=row, start_column=col_start,
        end_row=row, end_column=col_end
    )
    c = ws.cell(row=row, column=col_start)
    style_cell(c, value=text, bold=True, size=12, fcolor=txt_color,
               bg=hdr_fill, halign="center", bdr=thin_border())
    ws.row_dimensions[row].height = 22

def label_cell(ws, row, col, text, bg=LIGHT_GRAY):
    c = ws.cell(row=row, column=col)
    style_cell(c, value=text, bold=True, size=10, fcolor=DARK_GRAY,
               bg=bg, bdr=thin_border())

def input_cell(ws, row, col, value=None, formula=None,
               bg=INPUT_YELLOW, fmt=None, note=None):
    c = ws.cell(row=row, column=col)
    if formula:
        c.value = formula
        c.fill = fill(FORMULA_BG)
    else:
        if value is not None:
            c.value = value
        c.fill = fill(bg)
    c.border = thin_border()
    c.alignment = align("left", "center")
    c.font = font(size=10)
    if fmt:
        c.number_format = fmt


# ── GETTING STARTED SHEET ────────────────────────────────────────────────────
def build_getting_started(wb):
    ws = wb.create_sheet("Getting Started", 0)
    set_col_widths(ws, COL_WIDTHS_GS)

    # ── Title bar ──
    ws.merge_cells("A1:J1")
    c = ws["A1"]
    style_cell(c, value="Homeschool Curriculum Planner",
               bold=True, size=22, fcolor=WHITE, bg=NAVY,
               halign="center", valign="center")
    ws.row_dimensions[1].height = 48

    ws.merge_cells("A2:J2")
    c = ws["A2"]
    style_cell(c, value="Getting Started  —  School & Student Information",
               bold=True, size=13, fcolor=WHITE, bg=MED_BLUE,
               halign="center", valign="center")
    ws.row_dimensions[2].height = 26

    # ── School-year settings block (rows 4-10) ──
    ws.merge_cells("A4:J4")
    section_header(ws, 4, 1, 10, "SCHOOL YEAR SETTINGS", NAVY)

    settings = [
        (5,  "School Name",           "B5",  None,           "Westfield Academy"),
        (6,  "School Year",           "B6",  None,           "2025-2026"),
        (7,  "School Start Date",     "B7",  None,           None),
        (8,  "School End Date",       "B8",  None,           None),
        (9,  "Total School Weeks",    "B9",  '=IFERROR(ROUNDDOWN((B8-B7)/7,0),"")', None),
        (10, "School Days per Week",  "B10", None,           5),
    ]

    for row, lbl, cell_ref, formula, default in settings:
        label_cell(ws, row, 2, lbl)
        ws.merge_cells(f"C{row}:E{row}")
        c = ws[cell_ref.replace("B", "C")]  # we'll use column C merged
        # actually just use column B for input:
        pass

    # redo cleanly
    for row in range(5, 11):
        ws.row_dimensions[row].height = 20

    rows_data = [
        (5,  "School Name",          None,           "Westfield Academy"),
        (6,  "School Year",          None,           "2025-2026"),
        (7,  "School Start Date",    None,           None),
        (8,  "School End Date",      None,           None),
        (9,  "Total School Weeks",   '=IFERROR(ROUNDDOWN((B8-B7)/7,0),"")', None),
        (10, "School Days per Week", None,           5),
    ]

    for row, lbl, formula, default in rows_data:
        label_cell(ws, row, 2, lbl + ":", bg=LIGHT_GRAY)
        ws.merge_cells(f"C{row}:F{row}")
        c = ws.cell(row=row, column=3)
        if formula:
            c.value = formula
            c.fill = fill(FORMULA_BG)
            c.font = font(size=10, bold=True)
        else:
            if default is not None:
                c.value = default
            c.fill = fill(INPUT_YELLOW)
            c.font = font(size=10)
        c.border = thin_border()
        c.alignment = align("left", "center")
        if row in (7, 8):
            c.number_format = "MM/DD/YYYY"

    # hint text
    ws.merge_cells("G7:J7")
    ws["G7"].value = "← Enter as MM/DD/YYYY"
    ws["G7"].font = font(size=9, italic=True, color="808080")
    ws.merge_cells("G8:J8")
    ws["G8"].value = "← Enter as MM/DD/YYYY"
    ws["G8"].font = font(size=9, italic=True, color="808080")
    ws.merge_cells("G9:J9")
    ws["G9"].value = "← Auto-calculated from dates above"
    ws["G9"].font = font(size=9, italic=True, color="808080")

    # ── Student table (rows 12-23) ──
    ws.row_dimensions[11].height = 8  # spacer

    ws.merge_cells("A12:J12")
    section_header(ws, 12, 1, 10, "STUDENT ROSTER", NAVY)

    # Column headers (row 13)
    headers = ["#", "Student Name", "Date of Birth", "Class of\n(Grad Year)",
               "Grade Level\n(Auto-Calc)", "Curriculum Notes"]
    header_cols  = [2, 3, 4, 5, 6, 7]
    header_spans = [(2,2),(3,4),(5,5),(6,6),(7,7),(8,10)]

    ws.row_dimensions[13].height = 30
    for (cs, ce), h in zip(header_spans, headers):
        if cs != ce:
            ws.merge_cells(start_row=13, start_column=cs, end_row=13, end_column=ce)
        c = ws.cell(row=13, column=cs)
        style_cell(c, value=h, bold=True, size=10, fcolor=WHITE,
                   bg=MED_BLUE, halign="center", valign="center",
                   wrap=True, bdr=thin_border())

    # Student rows 14-23
    for i in range(NUM_STUDENTS):
        row = 14 + i
        ws.row_dimensions[row].height = 20
        # #
        c = ws.cell(row=row, column=2)
        style_cell(c, value=i+1, bold=True, size=10, fcolor=WHITE,
                   bg=NAVY if i % 2 == 0 else MED_BLUE,
                   halign="center", bdr=thin_border())

        # Name (merged C:D)
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=4)
        c = ws.cell(row=row, column=3)
        style_cell(c, bg=INPUT_YELLOW, bdr=thin_border())
        c.font = font(size=10)

        # DOB
        c = ws.cell(row=row, column=5)
        style_cell(c, bg=INPUT_YELLOW, bdr=thin_border())
        c.font = font(size=10)
        c.number_format = "MM/DD/YYYY"

        # Class of year
        c = ws.cell(row=row, column=6)
        style_cell(c, bg=INPUT_YELLOW, bdr=thin_border())
        c.font = font(size=10)
        c.number_format = "0"

        # Grade level (formula)
        c = ws.cell(row=row, column=7)
        c.value = f'=IFERROR(IF(C{row}="","",IF(F{row}="","",12-(F{row}-YEAR($B$8))))&" th","")'
        # cleaner formula:
        c.value = (
            f'=IF(OR(C{row}="",F{row}=""),"",LET(g,12-(F{row}-YEAR($B$8)),'
            f'IF(g=11,"11th",IF(g=12,"12th",IF(g=1,"1st",IF(g=2,"2nd",IF(g=3,"3rd",g&"th")))))))'
        )
        style_cell(c, bg=FORMULA_BG, bdr=thin_border(), halign="center")
        c.font = font(size=10, bold=True)

        # Notes (merged H:J)
        ws.merge_cells(start_row=row, start_column=8, end_row=row, end_column=10)
        c = ws.cell(row=row, column=8)
        style_cell(c, bg=INPUT_YELLOW, bdr=thin_border())
        c.font = font(size=10)

    # ── Legend (row 26+) ──
    ws.row_dimensions[25].height = 8
    ws.merge_cells("A26:J26")
    section_header(ws, 26, 1, 10, "LEGEND", "595959")

    legend = [
        (INPUT_YELLOW, "User Input Cell — type your data here"),
        (FORMULA_BG,   "Formula Cell — auto-calculated, do not edit"),
    ]
    for i, (color, desc) in enumerate(legend):
        row = 27 + i
        ws.row_dimensions[row].height = 18
        c = ws.cell(row=row, column=2)
        style_cell(c, bg=color, bdr=thin_border())
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=8)
        c = ws.cell(row=row, column=3)
        style_cell(c, value=desc, size=10, fcolor=DARK_GRAY)

    # ── Instructions (row 30+) ──
    ws.row_dimensions[29].height = 8
    ws.merge_cells("A30:J30")
    section_header(ws, 30, 1, 10, "HOW TO USE THIS PLANNER", "375623")

    instructions = [
        "1.  Fill in the School Year Settings above (name, year, start/end dates, days per week).",
        "2.  Enter each student's name, date of birth, and graduation year — grade is auto-calculated.",
        "3.  Navigate to each student's tab (tabs are pre-labeled Student 1 – Student 10).",
        "4.  For each subject, enter the curriculum title and total units (pages, lessons, chapters, etc.).",
        "5.  The planner will calculate how many units per week and per day to stay on track.",
        "6.  For multi-volume curricula (e.g. a math program with 10 books), use the Component Breakdown table.",
        "7.  Yellow cells = you type here.  Blue-tinted cells = formulas, leave them alone.",
    ]
    for i, text in enumerate(instructions):
        row = 31 + i
        ws.row_dimensions[row].height = 18
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=10)
        c = ws.cell(row=row, column=2)
        style_cell(c, value=text, size=10, fcolor=DARK_GRAY,
                   bg="F0FFF4" if i % 2 == 0 else WHITE)

    return ws


# ── STUDENT SHEET ─────────────────────────────────────────────────────────────
def build_student_sheet(wb, student_index):
    """student_index is 0-based."""
    n = student_index + 1
    gs_row = 13 + n   # row in Getting Started for this student

    ws = wb.create_sheet(f"Student {n}")
    set_col_widths(ws, COL_WIDTHS_ST)

    # ── Student header banner ──
    ws.merge_cells("A1:J1")
    c = ws["A1"]
    c.value = f"='Getting Started'!C{gs_row}"
    style_cell(c, bold=True, size=20, fcolor=WHITE, bg=NAVY,
               halign="center", valign="center")
    ws.row_dimensions[1].height = 44

    ws.merge_cells("A2:J2")
    c = ws["A2"]
    c.value = (
        f'=IFERROR("Grade: "&\'Getting Started\'!G{gs_row}'
        f'&"   |   School Year: "&\'Getting Started\'!C6'
        f'&"   |   "  &TEXT(\'Getting Started\'!B7,"MMM D, YYYY")'
        f'&"  –  "&TEXT(\'Getting Started\'!B8,"MMM D, YYYY"),"")'
    )
    style_cell(c, size=11, fcolor=WHITE, bg=MED_BLUE,
               halign="center", valign="center")
    ws.row_dimensions[2].height = 24

    # ── Pace summary row ──
    ws.row_dimensions[3].height = 8

    ws.merge_cells("A4:J4")
    section_header(ws, 4, 1, 10, "SCHOOL YEAR OVERVIEW  (from Getting Started)", MED_BLUE)

    labels = ["Start Date", "End Date", "Total Weeks", "Days/Week"]
    formulas = [
        "=IF('Getting Started'!B7<>\"\",TEXT('Getting Started'!B7,\"MMM D, YYYY\"),\"\")",
        "=IF('Getting Started'!B8<>\"\",TEXT('Getting Started'!B8,\"MMM D, YYYY\"),\"\")",
        "='Getting Started'!B9",
        "='Getting Started'!B10",
    ]
    label_cols  = [2, 4, 6, 8]
    value_cols  = [3, 5, 7, 9]

    ws.row_dimensions[5].height = 20
    for lbl, frm, lc, vc in zip(labels, formulas, label_cols, value_cols):
        label_cell(ws, 5, lc, lbl + ":", bg=LIGHT_BLUE)
        c = ws.cell(row=5, column=vc)
        c.value = frm
        c.fill = fill(FORMULA_BG)
        c.font = font(size=10, bold=True)
        c.border = thin_border()
        c.alignment = align("center", "center")

    # ── Subject sections ──
    current_row = 7

    for subj_name, hdr_fill, row_fill in SUBJECTS:
        current_row = add_subject_section(ws, current_row, subj_name,
                                          hdr_fill, row_fill)
        current_row += 1  # spacer between subjects

    return ws


def add_subject_section(ws, start_row, subj_name, hdr_fill, row_fill):
    """
    Draws one subject section.
    Layout (relative rows from start_row):
      0  Subject header
      1  Curriculum name / Unit type
      2  Total units | Weeks (ref) | /Week (calc) | /Day (calc) | Est end (calc)
      3  [Component breakdown header]
      4  [Component col headers]
      5-10 [6 component rows]
      11 Notes
    Returns the next available row.
    """
    r = start_row

    # Row 0: subject header
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
    c = ws.cell(row=r, column=1)
    style_cell(c, value=f"  {subj_name.upper()}", bold=True, size=13,
               fcolor=WHITE, bg=hdr_fill, halign="left", valign="center",
               bdr=outer_border())
    ws.row_dimensions[r].height = 24

    # Row 1: Curriculum name + unit label
    r += 1
    ws.row_dimensions[r].height = 20
    label_cell(ws, r, 1, "Curriculum / Title:", bg=row_fill)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
    c = ws.cell(row=r, column=2)
    c.fill = fill(INPUT_YELLOW); c.border = thin_border(); c.font = font(size=10)

    label_cell(ws, r, 6, "Unit type:", bg=row_fill)
    ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=8)
    c = ws.cell(row=r, column=7)
    c.value = "pages"   # default; user can change
    c.fill = fill(INPUT_YELLOW); c.border = thin_border(); c.font = font(size=10)

    label_cell(ws, r, 9, "(pages/lessons/etc)", bg=LIGHT_GRAY)
    ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=10)
    c = ws.cell(row=r, column=9)
    style_cell(c, value="← type: pages, lessons, chapters…",
               size=9, fcolor="808080", italic=True, bg=LIGHT_GRAY)

    # Row 2: Pace calculation strip
    r += 1
    ws.row_dimensions[r].height = 22

    pace_labels  = ["Total Units", "School Weeks", "Units / Week", "Units / Day", "Est. End Date"]
    pace_cols    = [1, 3, 5, 7, 9]     # label cols
    value_cols_p = [2, 4, 6, 8, 10]   # value cols

    input_row   = r
    total_units_col = 2

    for lbl, lc in zip(pace_labels, pace_cols):
        label_cell(ws, r, lc, lbl + ":", bg=row_fill)

    # Total units — user input
    c = ws.cell(row=r, column=2)
    c.fill = fill(INPUT_YELLOW); c.border = thin_border(); c.font = font(size=10)

    # School weeks — formula from GS
    c = ws.cell(row=r, column=4)
    c.value = "='Getting Started'!B9"
    c.fill = fill(FORMULA_BG); c.border = thin_border(); c.font = font(size=10)
    c.alignment = align("center", "center")

    # Units/week = total / weeks
    total_ref = f"{get_column_letter(2)}{r}"
    weeks_ref = f"{get_column_letter(4)}{r}"
    days_ref  = "'Getting Started'!B10"

    c = ws.cell(row=r, column=6)
    c.value = f'=IFERROR(IF({total_ref}="","",ROUND({total_ref}/{weeks_ref},1)),"")'
    c.fill = fill(FORMULA_BG); c.border = thin_border(); c.font = font(size=10, bold=True)
    c.alignment = align("center", "center")

    pw_ref = f"{get_column_letter(6)}{r}"

    # Units/day
    c = ws.cell(row=r, column=8)
    c.value = f'=IFERROR(IF({pw_ref}="","",ROUND({pw_ref}/{days_ref},1)),"")'
    c.fill = fill(FORMULA_BG); c.border = thin_border(); c.font = font(size=10, bold=True)
    c.alignment = align("center", "center")

    # Est end date
    c = ws.cell(row=r, column=10)
    c.value = (
        f'=IFERROR(IF({total_ref}="","",TEXT('
        f"'Getting Started'!B7+({total_ref}/{days_ref}),"
        f'"MMM D, YYYY")),"")'
    )
    c.fill = fill(FORMULA_BG); c.border = thin_border(); c.font = font(size=10, bold=True)
    c.alignment = align("center", "center")

    # Row 3: Component breakdown header
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
    c = ws.cell(row=r, column=1)
    style_cell(c, value="  Optional: Component / Volume Breakdown  (fill in if curriculum has multiple books or parts)",
               size=9, italic=True, fcolor=WHITE, bg=DARK_GRAY,
               halign="left", valign="center")
    ws.row_dimensions[r].height = 18

    # Row 4: Component column headers
    r += 1
    ws.row_dimensions[r].height = 18
    comp_hdrs = ["#", "Component / Book Title", "", "Total Units\nin Component",
                 "Cumulative\nUnits", "Start\nWeek", "End\nWeek",
                 "Start\nDate", "End\nDate", "Complete?"]
    comp_col_starts = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]

    for col, hdr in zip(comp_col_starts, comp_hdrs):
        c = ws.cell(row=r, column=col)
        style_cell(c, value=hdr, bold=True, size=9, fcolor=WHITE,
                   bg="595959", halign="center", valign="center",
                   wrap=True, bdr=thin_border())
    # merge title cols 2-3
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)

    comp_header_row = r

    # Rows 5-10: 6 component rows
    COMP_ROWS = 6
    component_start = r + 1
    for ci in range(COMP_ROWS):
        r += 1
        ws.row_dimensions[r].height = 18
        row_bg = INPUT_YELLOW if ci % 2 == 0 else "FAFAD2"

        # # label
        c = ws.cell(row=r, column=1)
        style_cell(c, value=ci + 1, bold=True, size=10, fcolor=DARK_GRAY,
                   bg=LIGHT_GRAY, halign="center", bdr=thin_border())

        # Title (merged 2-3)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c = ws.cell(row=r, column=2)
        c.fill = fill(row_bg); c.border = thin_border(); c.font = font(size=10)

        # Total units in component
        c = ws.cell(row=r, column=4)
        c.fill = fill(row_bg); c.border = thin_border(); c.font = font(size=10)
        c.alignment = align("center", "center")

        # Cumulative units (sum of col4 from row component_start to this row)
        c = ws.cell(row=r, column=5)
        units_range = f"D{component_start}:D{r}"
        c.value = f"=IFERROR(SUM({units_range}),\"\")"
        c.fill = fill(FORMULA_BG); c.border = thin_border(); c.font = font(size=10)
        c.alignment = align("center", "center")

        # Start week: cumulative units of previous row / units per week  + 1
        pace_row = input_row  # row where Total Units and units/week live
        total_units_cell = f"{get_column_letter(total_units_col)}{pace_row}"
        upw_cell = f"{get_column_letter(6)}{pace_row}"

        if ci == 0:
            start_week_formula = '=IF(D{r}="","",1)'.format(r=r)
        else:
            prev_cum = f"E{r-1}"
            start_week_formula = (
                f'=IFERROR(IF(D{r}="","",FLOOR({prev_cum}/{upw_cell},1)+1),"")'
            )

        c = ws.cell(row=r, column=6)
        c.value = start_week_formula
        c.fill = fill(FORMULA_BG); c.border = thin_border(); c.font = font(size=10)
        c.alignment = align("center", "center")

        # End week: cumulative / units per week
        c = ws.cell(row=r, column=7)
        cum_cell = f"E{r}"
        c.value = f'=IFERROR(IF(D{r}="","",CEILING({cum_cell}/{upw_cell},1)),"")'
        c.fill = fill(FORMULA_BG); c.border = thin_border(); c.font = font(size=10)
        c.alignment = align("center", "center")

        # Start date
        c = ws.cell(row=r, column=8)
        start_date_gs = "'Getting Started'!B7"
        start_wk_cell = f"F{r}"
        c.value = f'=IFERROR(IF(D{r}="","",TEXT({start_date_gs}+({start_wk_cell}-1)*7,"MMM D")),"")'
        c.fill = fill(FORMULA_BG); c.border = thin_border(); c.font = font(size=10)
        c.alignment = align("center", "center")

        # End date
        c = ws.cell(row=r, column=9)
        end_wk_cell = f"G{r}"
        c.value = f'=IFERROR(IF(D{r}="","",TEXT({start_date_gs}+({end_wk_cell})*7,"MMM D")),"")'
        c.fill = fill(FORMULA_BG); c.border = thin_border(); c.font = font(size=10)
        c.alignment = align("center", "center")

        # Complete checkbox
        c = ws.cell(row=r, column=10)
        c.fill = fill(INPUT_YELLOW); c.border = thin_border(); c.font = font(size=10)
        c.alignment = align("center", "center")

    # Row: Notes
    r += 1
    ws.row_dimensions[r].height = 20
    label_cell(ws, r, 1, "Notes:", bg=row_fill)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=10)
    c = ws.cell(row=r, column=2)
    c.fill = fill(INPUT_YELLOW); c.border = thin_border(); c.font = font(size=10)
    c.alignment = align("left", "center", wrap=True)

    return r


# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    wb = openpyxl.Workbook()
    # remove default blank sheet
    wb.remove(wb.active)

    build_getting_started(wb)

    for i in range(NUM_STUDENTS):
        build_student_sheet(wb, i)

    # Tab colors: student tabs get color matching their position
    tab_colors = [
        "1F3864","2E74B5","5B2C6F","922B21","784212",
        "7D6608","0E6655","6E2F0E","1E8449","76448A",
    ]
    for i, sheet_name in enumerate([f"Student {n}" for n in range(1, 11)]):
        wb[sheet_name].sheet_properties.tabColor = tab_colors[i % len(tab_colors)]
    wb["Getting Started"].sheet_properties.tabColor = "1F3864"

    out_path = "/home/user/curriculum-planning/Homeschool_Curriculum_Planner.xlsx"
    wb.save(out_path)
    print(f"Saved: {out_path}")


if __name__ == "__main__":
    main()
