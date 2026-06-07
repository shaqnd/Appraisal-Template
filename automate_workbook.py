"""
Full automation overhaul for Analysis_Workbook_OPTIMIZED.xlsx
Outputs Analysis_Workbook_OPTIMIZED.xlsm with:
  1. Property Info expansion (rows 32-60, RealWare fields + dropdowns)
  2. Dynamic H&BU narratives (formula-driven columns D/E/F)
  3. Word Export Prep sheet
  4. Dashboard sheet (inserted as first tab)
  5. .xlsm format (VBA-ready)
"""

import shutil
import openpyxl
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.table import Table, TableStyleInfo

SRC = "/home/user/Appraisal-Template/Analysis_Workbook_OPTIMIZED.xlsx"
DST = "/home/user/Appraisal-Template/Analysis_Workbook_OPTIMIZED.xlsm"

shutil.copy2(SRC, DST)
wb = load_workbook(DST, keep_vba=True)

# ─── helpers ────────────────────────────────────────────────────────────────

def safe_set(ws, cell_ref, value):
    cell = ws[cell_ref]
    if not isinstance(cell, MergedCell):
        cell.value = value
        return True
    return False

def thin_border():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

YELLOW_FILL  = fill("FFFF99")   # required manual input
GREEN_FILL   = fill("C6EFCE")   # auto-calc / formula
HEADER_FILL  = fill("1F4E79")   # dark-blue section header
LABEL_FILL   = fill("D9E1F2")   # light-blue label

HEADER_FONT  = Font(bold=True, color="FFFFFF", size=11)
LABEL_FONT   = Font(bold=True, color="1F4E79", size=10)
VALUE_FONT   = Font(size=10)

def write_label(ws, row, col, text, label_fill=True):
    c = ws.cell(row=row, column=col, value=text)
    c.font = LABEL_FONT
    if label_fill:
        c.fill = LABEL_FILL
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    c.border = thin_border()
    return c

def write_value(ws, row, col, value, req=True, fmt=None):
    c = ws.cell(row=row, column=col, value=value)
    c.fill = YELLOW_FILL if req else GREEN_FILL
    c.font = VALUE_FONT
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    c.border = thin_border()
    if fmt:
        c.number_format = fmt
    return c

def section_header(ws, row, col_start, col_end, text):
    ws.cell(row=row, column=col_start, value=text).font = HEADER_FONT
    ws.cell(row=row, column=col_start).fill = HEADER_FILL
    ws.cell(row=row, column=col_start).alignment = Alignment(
        horizontal="center", vertical="center"
    )
    ws.merge_cells(
        start_row=row, start_column=col_start,
        end_row=row, end_column=col_end
    )
    ws.row_dimensions[row].height = 20

def add_dv(ws, cell_ref, options_str):
    dv = DataValidation(
        type="list",
        formula1=f'"{options_str}"',
        showDropDown=False,
        showErrorMessage=True,
        errorTitle="Invalid value",
        error="Please select from the list."
    )
    dv.sqref = cell_ref
    ws.add_data_validation(dv)

# ─── 1. PROPERTY INFO EXPANSION ─────────────────────────────────────────────

PI_NAME = "Property Info"
if PI_NAME not in wb.sheetnames:
    print(f"WARNING: sheet '{PI_NAME}' not found – skipping PI expansion")
else:
    pi = wb[PI_NAME]

    # Insert blank rows 32-60 to make room (existing rows preserved above 32)
    # openpyxl doesn't have insert_rows for xlsx easily; we'll just write into
    # rows 32+ which should be empty / available
    PI = "'Property Info'!"   # for cross-sheet formulas elsewhere

    # Section: Owner & Legal
    section_header(pi, 32, 1, 4, "OWNER & LEGAL INFORMATION")
    fields_owner = [
        (33, "Owner Name",          None,  True),
        (34, "Owner Address",       None,  True),
        (35, "Owner City/State/Zip",None,  True),
        (36, "Legal Description",   None,  True),
        (37, "Subdivision / Plat",  None,  True),
        (38, "Parcel / Schedule #", None,  True),
    ]
    for r, label, val, req in fields_owner:
        write_label(pi, r, 1, label)
        write_value(pi, r, 2, val, req=req)
        pi.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        pi.row_dimensions[r].height = 18

    # Section: Building Details
    section_header(pi, 39, 1, 4, "BUILDING DETAILS  (from RealWare)")
    bldg_fields = [
        (40, "Quality Class",        True,  "Excellent,Good,Average,Fair,Poor"),
        (41, "Construction Class",   True,  "A,B,C,D,S,W"),
        (42, "Number of Stories",    True,  None),
        (43, "Year Built",           True,  None),
        (44, "Effective Year Built", True,  None),
        (45, "Gross Building Area (SF)", True, None),
        (46, "Net Rentable Area (SF)",   True, None),
        (47, "Clear Height (ft)",    True,  None),
        (48, "Dock Doors (#)",       True,  None),
        (49, "Drive-In Doors (#)",   True,  None),
        (50, "Sprinkler System",     True,  "Yes,No,Partial"),
        (51, "HVAC Type",            True,  "Full HVAC,Partial,None,Evaporative"),
        (52, "Parking Spaces",       True,  None),
        (53, "Condition",            True,  "Excellent,Good,Average,Fair,Poor"),
    ]
    for r, label, req, dv_opts in bldg_fields:
        write_label(pi, r, 1, label)
        write_value(pi, r, 2, None, req=req)
        pi.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        pi.row_dimensions[r].height = 18
        if dv_opts:
            add_dv(pi, f"B{r}", dv_opts)

    # Section: Site Details
    section_header(pi, 54, 1, 4, "SITE DETAILS  (from RealWare)")
    site_fields = [
        (55, "Flood Zone Designation", True, "X,AE,A,AO,VE,D"),
        (56, "FEMA Map Panel #",       True, None),
        (57, "Zoning District",        True, None),
        (58, "Zoning Jurisdiction",    True, None),
    ]
    for r, label, req, dv_opts in site_fields:
        write_label(pi, r, 1, label)
        write_value(pi, r, 2, None, req=req)
        pi.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        pi.row_dimensions[r].height = 18
        if dv_opts:
            add_dv(pi, f"B{r}", dv_opts)

    # Section: Assessed Values
    section_header(pi, 59, 1, 4, "ASSESSED VALUES  (from RealWare)")
    av_fields = [
        (60, "Land Assessed Value",       True,  "$#,##0"),
        (61, "Building Assessed Value",   True,  "$#,##0"),
        (62, "Total Assessed Value",      False, "$#,##0"),   # auto = sum
        (63, "Prior Year Total AV",       True,  "$#,##0"),
        (64, "Tax Rate (mill levy)",      True,  "0.000"),
        (65, "Annual Real Estate Taxes",  False, "$#,##0"),   # auto = AV * rate
    ]
    for r, label, req, fmt in av_fields:
        write_label(pi, r, 1, label)
        write_value(pi, r, 2, None, req=req, fmt=fmt)
        pi.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        pi.row_dimensions[r].height = 18

    # Wire auto-calc formulas
    pi["B62"].value = "=IF(AND(ISNUMBER(B60),ISNUMBER(B61)),B60+B61,\"\")"
    pi["B62"].fill  = GREEN_FILL
    pi["B65"].value = "=IF(AND(ISNUMBER(B62),ISNUMBER(B64)),B62*B64/1000,\"\")"
    pi["B65"].fill  = GREEN_FILL

    # Column widths
    pi.column_dimensions["A"].width = 32
    pi.column_dimensions["B"].width = 26

    print("Property Info expanded.")

# ─── 2. DYNAMIC H&BU NARRATIVES ─────────────────────────────────────────────

HBU_NAME = "H&BU"
if HBU_NAME not in wb.sheetnames:
    # Try alternate names
    for candidate in ["HBU", "H & B U", "Highest & Best Use"]:
        if candidate in wb.sheetnames:
            HBU_NAME = candidate
            break

if HBU_NAME in wb.sheetnames:
    hbu = wb[HBU_NAME]
    PI = "'Property Info'!"

    # Column headers for the narrative system
    hbu.cell(row=1, column=4, value="Generated Draft").font = Font(bold=True)
    hbu.cell(row=1, column=5, value="Appraiser Override").font = Font(bold=True)
    hbu.cell(row=1, column=6, value="FINAL (use this)").font = Font(bold=True)
    for col in [4, 5, 6]:
        hbu.cell(row=1, column=col).fill = fill("1F4E79")
        hbu.cell(row=1, column=col).font = Font(bold=True, color="FFFFFF")
        hbu.cell(row=1, column=col).alignment = Alignment(horizontal="center")
    hbu.column_dimensions["D"].width = 80
    hbu.column_dimensions["E"].width = 80
    hbu.column_dimensions["F"].width = 80

    # Map row -> narrative formula for AS VACANT and AS IMPROVED
    # Rows are approximate; adjust to match actual H&BU layout
    narratives = {
        # AS VACANT
        4: (  # Physically Possible
            "=\"The subject site encompasses \"&TEXT('Property Info'!B11,\"#,##0\")&\" square feet "
            "(\"&TEXT('Property Info'!B12,\"0.00\")&\" acres) located at \"&'Property Info'!B4&\", \"&'Property Info'!B5&\". "
            "The site is \"&IF('Property Info'!B55=\"X\",\"not located in a FEMA designated flood hazard area\","
            "\"located in FEMA Flood Zone \"&'Property Info'!B55)&\", "
            "per FEMA Map Panel \"&'Property Info'!B56&\". "
            "Site dimensions and topography are considered physically adequate to support a variety of improvements.\""
        ),
        5: (  # Legally Permissible
            "=\"The subject site is zoned \"&'Property Info'!B57&\" by the \"&'Property Info'!B58&\". "
            "Permitted uses under this designation include those consistent with the subject neighborhood. "
            "Legally permissible uses are those allowed under the current zoning ordinance.\""
        ),
        6: (  # Financially Feasible
            "=\"Based on current market conditions in the \"&'Property Info'!B5&\" submarket, "
            "financially feasible uses are those generating sufficient revenue to justify construction costs "
            "and provide a market return to the developer.\""
        ),
        7: (  # Maximally Productive
            "=\"The maximally productive use of the subject as vacant is concluded to be development consistent "
            "with the current zoning of \"&'Property Info'!B57&\", which reflects current market demand in the area.\""
        ),
        # AS IMPROVED
        12: (  # Physically Possible
            "=\"The existing improvements consist of a \"&TEXT('Property Info'!B45,\"#,##0\")&\" SF building "
            "constructed in \"&TEXT('Property Info'!B43,\"0\")&\" with an effective year of \"&TEXT('Property Info'!B44,\"0\")&\". "
            "The improvements are in \"&LOWER('Property Info'!B53)&\" condition.\""
        ),
        13: (  # Legally Permissible
            "=\"The existing improvements conform to current zoning requirements under the \"&'Property Info'!B57&\" designation. "
            "The use is legally permissible as a conforming use.\""
        ),
        14: (  # Financially Feasible
            "=\"The subject improvements generate value in excess of land value alone, indicating that continuation "
            "of the existing use is financially feasible under current market conditions.\""
        ),
        15: (  # Maximally Productive
            "=\"The maximally productive use of the subject as improved is the continued use as currently improved, "
            "subject to the income-generating potential of the existing \"&LOWER('Property Info'!B57)&\" improvements.\""
        ),
    }

    for row, formula in narratives.items():
        hbu.cell(row=row, column=4).value = formula
        hbu.cell(row=row, column=4).fill = GREEN_FILL
        hbu.cell(row=row, column=4).alignment = Alignment(wrap_text=True)

        hbu.cell(row=row, column=5).fill = YELLOW_FILL
        hbu.cell(row=row, column=5).alignment = Alignment(wrap_text=True)

        hbu.cell(row=row, column=6).value = f"=IF(E{row}<>\"\",E{row},D{row})"
        hbu.cell(row=row, column=6).fill = fill("E2EFDA")
        hbu.cell(row=row, column=6).alignment = Alignment(wrap_text=True)
        hbu.row_dimensions[row].height = 60

    print("H&BU dynamic narratives added.")
else:
    print(f"WARNING: H&BU sheet not found (tried '{HBU_NAME}') – skipping narratives")

# ─── 3. WORD EXPORT PREP SHEET ──────────────────────────────────────────────

WEP = "Word Export Prep"
if WEP in wb.sheetnames:
    del wb[WEP]  # rebuild fresh if exists

wep = wb.create_sheet(WEP)
PI = "'Property Info'!"

wep.column_dimensions["A"].width = 28
wep.column_dimensions["B"].width = 45
wep.column_dimensions["C"].width = 20
wep.column_dimensions["D"].width = 18

# Master header
wep.merge_cells("A1:D1")
wep["A1"] = "WORD EXPORT PREP  —  All narrative fields in one place"
wep["A1"].font = Font(bold=True, color="FFFFFF", size=13)
wep["A1"].fill = HEADER_FILL
wep["A1"].alignment = Alignment(horizontal="center", vertical="center")
wep.row_dimensions[1].height = 26

wep["A2"] = "Field"
wep["B2"] = "Value (copy to Word)"
wep["C2"] = "Source"
wep["D2"] = "Notes"
for col_letter in ["A","B","C","D"]:
    c = wep[f"{col_letter}2"]
    c.font = Font(bold=True, color="1F4E79")
    c.fill = LABEL_FILL
    c.border = thin_border()

sections = [
    # (section_title, [(field_label, formula_or_value, source_label), ...])
    ("SUBJECT PROPERTY IDENTIFICATION", [
        ("Appraisal Date",          f"={PI}B2",                           "Property Info B2"),
        ("Property Address",        f"={PI}B4",                           "Property Info B4"),
        ("City",                    f"={PI}B5",                           "Property Info B5"),
        ("County",                  f"={PI}B6",                           "Property Info B6"),
        ("State / Zip",             f"={PI}B7",                           "Property Info B7"),
        ("Parcel / Schedule #",     f"={PI}B38",                          "Property Info B38"),
        ("Legal Description",       f"={PI}B36",                          "Property Info B36"),
        ("Owner Name",              f"={PI}B33",                          "Property Info B33"),
    ]),
    ("SITE DATA", [
        ("Land Area (SF)",          f"=TEXT({PI}B11,\"#,##0\")&\" SF\"", "Property Info B11"),
        ("Land Area (Acres)",       f"=TEXT({PI}B12,\"0.00\")&\" AC\"",  "Property Info B12"),
        ("Zoning",                  f"={PI}B57&\" — \"&{PI}B58",         "Property Info B57/B58"),
        ("Flood Zone",              f"=IF({PI}B55=\"X\",\"Zone X (unshaded) — not in FEMA flood hazard area\",\"Zone \"&{PI}B55)", "Property Info B55"),
        ("FEMA Map Panel",          f"={PI}B56",                          "Property Info B56"),
    ]),
    ("IMPROVEMENT DATA", [
        ("Building Size (GBA)",     f"=TEXT({PI}B45,\"#,##0\")&\" SF\"",  "Property Info B45"),
        ("Net Rentable Area",       f"=TEXT({PI}B46,\"#,##0\")&\" SF\"",  "Property Info B46"),
        ("Year Built",              f"=TEXT({PI}B43,\"0\")",              "Property Info B43"),
        ("Effective Year Built",    f"=TEXT({PI}B44,\"0\")",              "Property Info B44"),
        ("Quality Class",           f"={PI}B40",                          "Property Info B40"),
        ("Construction Class",      f"={PI}B41",                          "Property Info B41"),
        ("Stories",                 f"={PI}B42",                          "Property Info B42"),
        ("Clear Height",            f"=TEXT({PI}B47,\"0\")&\" ft\"",     "Property Info B47"),
        ("Dock Doors",              f"=TEXT({PI}B48,\"0\")",              "Property Info B48"),
        ("Drive-In Doors",          f"=TEXT({PI}B49,\"0\")",              "Property Info B49"),
        ("Sprinkler System",        f"={PI}B50",                          "Property Info B50"),
        ("HVAC",                    f"={PI}B51",                          "Property Info B51"),
        ("Parking",                 f"=TEXT({PI}B52,\"0\")&\" spaces\"", "Property Info B52"),
        ("Condition",               f"={PI}B53",                          "Property Info B53"),
    ]),
    ("ASSESSED VALUES", [
        ("Land Assessed Value",     f"=TEXT({PI}B60,\"$#,##0\")",         "Property Info B60"),
        ("Building Assessed Value", f"=TEXT({PI}B61,\"$#,##0\")",         "Property Info B61"),
        ("Total Assessed Value",    f"=TEXT({PI}B62,\"$#,##0\")",         "Property Info B62"),
        ("Prior Year Total AV",     f"=TEXT({PI}B63,\"$#,##0\")",         "Property Info B63"),
        ("Tax Rate (mill levy)",    f"=TEXT({PI}B64,\"0.000\")&\" mills\"","Property Info B64"),
        ("Annual Taxes",            f"=TEXT({PI}B65,\"$#,##0\")",         "Property Info B65"),
    ]),
    ("VALUE CONCLUSIONS", [
        ("Cost Approach Value",     f"='Cost Approach'!D33",              "Cost Approach D33"),
        ("Sales Comp Value",        f"='Basic Sales Grid'!D4",            "Basic Sales Grid D4"),
        ("Income Approach Value",   f"='Income Approach'!D37",            "Income Approach D37"),
        ("Final Value Conclusion",  "",                                    "Enter manually"),
        ("Effective Date",          f"={PI}B2",                           "Property Info B2"),
    ]),
    ("H&BU NARRATIVE (AS VACANT)", [
        ("Physically Possible",     "='H&BU'!F4",                        "H&BU F4"),
        ("Legally Permissible",     "='H&BU'!F5",                        "H&BU F5"),
        ("Financially Feasible",    "='H&BU'!F6",                        "H&BU F6"),
        ("Maximally Productive",    "='H&BU'!F7",                        "H&BU F7"),
    ]),
    ("H&BU NARRATIVE (AS IMPROVED)", [
        ("Physically Possible",     "='H&BU'!F12",                       "H&BU F12"),
        ("Legally Permissible",     "='H&BU'!F13",                       "H&BU F13"),
        ("Financially Feasible",    "='H&BU'!F14",                       "H&BU F14"),
        ("Maximally Productive",    "='H&BU'!F15",                       "H&BU F15"),
    ]),
    ("APPRAISER INFO", [
        ("Appraiser Name",          "",                                    "Enter manually"),
        ("Appraiser Certification", "",                                    "Enter manually"),
        ("Supervisor Name",         "",                                    "Enter manually"),
        ("Report Date",             "",                                    "Enter manually"),
    ]),
]

current_row = 3
for sec_title, fields in sections:
    wep.merge_cells(
        start_row=current_row, start_column=1,
        end_row=current_row, end_column=4
    )
    c = wep.cell(row=current_row, column=1, value=sec_title)
    c.font = HEADER_FONT
    c.fill = HEADER_FILL
    c.alignment = Alignment(horizontal="left", vertical="center")
    wep.row_dimensions[current_row].height = 18
    current_row += 1

    for label, formula, source in fields:
        wep.cell(row=current_row, column=1, value=label).border = thin_border()
        wep.cell(row=current_row, column=1).fill = LABEL_FILL
        wep.cell(row=current_row, column=1).font = LABEL_FONT

        val_cell = wep.cell(row=current_row, column=2, value=formula)
        val_cell.border = thin_border()
        val_cell.fill = GREEN_FILL if formula.startswith("=") else YELLOW_FILL
        val_cell.alignment = Alignment(wrap_text=True, vertical="center")

        wep.cell(row=current_row, column=3, value=source).border = thin_border()
        wep.cell(row=current_row, column=3).font = Font(italic=True, size=9, color="595959")

        wep.row_dimensions[current_row].height = 18
        current_row += 1

    current_row += 1  # blank row between sections

print("Word Export Prep sheet built.")

# ─── 4. DASHBOARD SHEET ─────────────────────────────────────────────────────

DASH = "Dashboard"
if DASH in wb.sheetnames:
    del wb[DASH]

dash = wb.create_sheet(DASH, 0)  # insert as first sheet

dash.column_dimensions["A"].width = 32
dash.column_dimensions["B"].width = 35
dash.column_dimensions["C"].width = 18
dash.column_dimensions["D"].width = 18
dash.column_dimensions["E"].width = 14

# ── Banner ──
dash.merge_cells("A1:E1")
dash["A1"] = "APPRAISAL ANALYSIS DASHBOARD"
dash["A1"].font = Font(bold=True, size=16, color="FFFFFF")
dash["A1"].fill = HEADER_FILL
dash["A1"].alignment = Alignment(horizontal="center", vertical="center")
dash.row_dimensions[1].height = 32

dash.merge_cells("A2:E2")
dash["A2"] = "Adams County Assessor — Industrial Property Analysis"
dash["A2"].font = Font(italic=True, size=11, color="FFFFFF")
dash["A2"].fill = fill("2E74B5")
dash["A2"].alignment = Alignment(horizontal="center", vertical="center")
dash.row_dimensions[2].height = 22

# ── Property Quick-Info ──
dash.merge_cells("A3:E3")
dash["A3"] = "PROPERTY QUICK REFERENCE"
dash["A3"].font = HEADER_FONT
dash["A3"].fill = fill("2E74B5")
dash["A3"].alignment = Alignment(horizontal="center")
dash.row_dimensions[3].height = 18

quick_info = [
    ("Property Address",  "='Property Info'!B4"),
    ("City / County",     "='Property Info'!B5&\"  |  \"&'Property Info'!B6"),
    ("Parcel #",          "='Property Info'!B38"),
    ("Appraisal Date",    "='Property Info'!B2"),
    ("Land Area",         "=TEXT('Property Info'!B11,\"#,##0\")&\" SF  (\"&TEXT('Property Info'!B12,\"0.00\")&\" AC)\""),
    ("Bldg Size (GBA)",   "=TEXT('Property Info'!B45,\"#,##0\")&\" SF\""),
    ("Year Built",        "=TEXT('Property Info'!B43,\"0\")"),
    ("Quality / Class",   "='Property Info'!B40&\" / Class \"&'Property Info'!B41"),
]

for i, (label, formula) in enumerate(quick_info):
    r = 4 + i
    dash.cell(row=r, column=1, value=label).fill = LABEL_FILL
    dash.cell(row=r, column=1).font = LABEL_FONT
    dash.cell(row=r, column=1).border = thin_border()
    dash.cell(row=r, column=2, value=formula).fill = fill("EBF3FB")
    dash.cell(row=r, column=2).border = thin_border()
    dash.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    dash.row_dimensions[r].height = 18

# ── Value Summary ──
r_vs = 4 + len(quick_info) + 1
dash.merge_cells(f"A{r_vs}:E{r_vs}")
dash.cell(row=r_vs, column=1, value="VALUE SUMMARY").font = HEADER_FONT
dash.cell(row=r_vs, column=1).fill = fill("2E74B5")
dash.cell(row=r_vs, column=1).alignment = Alignment(horizontal="center")
dash.row_dimensions[r_vs].height = 18
r_vs += 1

value_rows = [
    ("Cost Approach",          "='Cost Approach'!D33",       "$#,##0"),
    ("Sales Comparison",       "='Basic Sales Grid'!D4",     "$#,##0"),
    ("Income Approach",        "='Income Approach'!D37",     "$#,##0"),
    ("Assessed Value (Total)", "='Property Info'!B62",       "$#,##0"),
    ("Annual Taxes",           "='Property Info'!B65",       "$#,##0"),
]
for label, formula, fmt in value_rows:
    dash.cell(row=r_vs, column=1, value=label).fill = LABEL_FILL
    dash.cell(row=r_vs, column=1).font = LABEL_FONT
    dash.cell(row=r_vs, column=1).border = thin_border()
    vc = dash.cell(row=r_vs, column=2, value=formula)
    vc.fill = fill("EBF3FB")
    vc.border = thin_border()
    vc.number_format = fmt
    dash.merge_cells(start_row=r_vs, start_column=2, end_row=r_vs, end_column=3)
    dash.row_dimensions[r_vs].height = 18
    r_vs += 1

# ── Completion Tracker ──
r_ct = r_vs + 1
dash.merge_cells(f"A{r_ct}:E{r_ct}")
dash.cell(row=r_ct, column=1, value="COMPLETION TRACKER").font = HEADER_FONT
dash.cell(row=r_ct, column=1).fill = HEADER_FILL
dash.cell(row=r_ct, column=1).alignment = Alignment(horizontal="center")
dash.row_dimensions[r_ct].height = 18
r_ct += 1

# Column headers
for col, hdr in enumerate(["Section", "Key Field / Check", "Status", "Tab", "Notes"], 1):
    c = dash.cell(row=r_ct, column=col, value=hdr)
    c.font = Font(bold=True, color="1F4E79")
    c.fill = LABEL_FILL
    c.border = thin_border()
r_ct += 1

tracker_items = [
    ("Property Info",     "Address / Parcel entered",      "='Property Info'!B4",         "Property Info", ""),
    ("Property Info",     "Appraisal Date set",            "='Property Info'!B2",         "Property Info", ""),
    ("Property Info",     "Land area entered",             "='Property Info'!B11",        "Property Info", "SF"),
    ("Property Info",     "Building size entered",         "='Property Info'!B45",        "Property Info", "GBA SF"),
    ("Property Info",     "Year built entered",            "='Property Info'!B43",        "Property Info", ""),
    ("Property Info",     "Owner name entered",            "='Property Info'!B33",        "Property Info", ""),
    ("Property Info",     "Assessed values entered",       "='Property Info'!B62",        "Property Info", "Total AV"),
    ("Land Sales",        "Comps entered",                 "='Land Sales'!C3",            "Land Sales",    ""),
    ("Basic Sales Grid",  "Comps entered",                 "='Basic Sales Grid'!E6",      "Basic Sales Grid", ""),
    ("Income Approach",   "Market rent entered",           "='Income Approach'!D8",       "Income Approach", ""),
    ("Income Approach",   "Cap rate entered",              "='Income Approach'!D27",      "Income Approach", ""),
    ("Cost Approach",     "RCN entered",                   "='Cost Approach'!D10",        "Cost Approach", ""),
    ("H&BU",              "As Vacant narrative reviewed",  "='H&BU'!F7",                  "H&BU", "Check col F"),
    ("H&BU",              "As Improved narrative reviewed","='H&BU'!F15",                 "H&BU", "Check col F"),
    ("Word Export Prep",  "Final value conclusion entered","='Word Export Prep'!B50",     "Word Export Prep", "Manual"),
]

for section, check, formula, tab, notes in tracker_items:
    # Status formula: ✅ if the cell has a non-empty value, else ⬜
    status_formula = f'=IF(ISBLANK({formula.lstrip("=")})|ISERROR({formula.lstrip("=")})|({formula.lstrip("=")}=""),"⬜","✅")'

    dash.cell(row=r_ct, column=1, value=section).border = thin_border()
    dash.cell(row=r_ct, column=1).fill = fill("F2F2F2")
    dash.cell(row=r_ct, column=2, value=check).border = thin_border()
    dash.cell(row=r_ct, column=3, value=status_formula).border = thin_border()
    dash.cell(row=r_ct, column=3).alignment = Alignment(horizontal="center", vertical="center")
    dash.cell(row=r_ct, column=4, value=tab).border = thin_border()
    dash.cell(row=r_ct, column=4).font = Font(italic=True, size=9)
    dash.cell(row=r_ct, column=5, value=notes).border = thin_border()
    dash.cell(row=r_ct, column=5).font = Font(italic=True, size=9)
    dash.row_dimensions[r_ct].height = 18
    r_ct += 1

# Progress bar (text-based)
total_items = len(tracker_items)
completed_formula = (
    f"=COUNTIF(C{r_ct - total_items}:C{r_ct - 1},\"✅\")"
)
pct_formula = f"={completed_formula.lstrip('=')}/{total_items}"

r_ct += 1
dash.merge_cells(f"A{r_ct}:B{r_ct}")
dash.cell(row=r_ct, column=1, value="Completed Items:").font = Font(bold=True)
dash.cell(row=r_ct, column=3, value=completed_formula).font = Font(bold=True, size=12, color="1F4E79")
dash.cell(row=r_ct, column=3).border = thin_border()
dash.cell(row=r_ct, column=3).alignment = Alignment(horizontal="center")

r_ct += 1
dash.merge_cells(f"A{r_ct}:B{r_ct}")
dash.cell(row=r_ct, column=1, value="% Complete:").font = Font(bold=True)
pct_cell = dash.cell(row=r_ct, column=3, value=f"={pct_formula.lstrip('=')}")
pct_cell.font = Font(bold=True, size=12, color="1F4E79")
pct_cell.number_format = "0%"
pct_cell.border = thin_border()
pct_cell.alignment = Alignment(horizontal="center")

dash.sheet_view.showGridLines = False

print("Dashboard sheet built.")

# ─── 5. SAVE AS .xlsm ───────────────────────────────────────────────────────

wb.save(DST)
print(f"\nSaved: {DST}")
print("Sheets:", wb.sheetnames)
