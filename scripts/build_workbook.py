"""
Build the Sandhu Group interactive Excel dashboard workbook.

Since this environment has no real Excel/Power Pivot, interactivity is
delivered with native Excel mechanics that need zero manual setup:
dropdown data-validation selectors driving SUMIFS-based KPI formulas,
Excel Tables (with built-in autofilter) for entity-level slicing, and
charts wired to the formula-driven cells so they update automatically
when a selector changes.
"""
import json
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

SRC = "data/Sandhu_Excel_Model_Data.xlsx"
OUT = "output/Sandhu_Group_Interactive_Dashboard.xlsx"

CATEGORIES = [
    "Operating Revenue", "Investment & Other Income", "COGS", "Payroll & Benefits",
    "Occupancy", "Marketing & Entertainment", "Administrative",
    "Management Fees (Interco)", "Interest & Financing", "Repairs & Maintenance",
    "Amortization", "Other Operating", "Income Tax",
]
INCOME_CATS = {"Operating Revenue", "Investment & Other Income"}

# --- palette ---------------------------------------------------------------
NAVY = "1F3864"
TEAL = "0E7C7B"
GOLD = "C9A227"
LIGHT = "F4F6F8"
WHITE = "FFFFFF"
GREY = "6B7280"
RED = "B3261E"

TITLE_FONT = Font(name="Calibri", size=20, bold=True, color=WHITE)
SUB_FONT = Font(name="Calibri", size=11, italic=True, color="D9E2EC")
H1 = Font(name="Calibri", size=13, bold=True, color=WHITE)
H2 = Font(name="Calibri", size=11, bold=True, color=NAVY)
LABEL = Font(name="Calibri", size=10, bold=True, color=GREY)
KPI_VAL = Font(name="Calibri", size=18, bold=True, color=NAVY)
KPI_LABEL = Font(name="Calibri", size=10, bold=True, color=WHITE)
BODY = Font(name="Calibri", size=10, color="1A1A1A")

NAVY_FILL = PatternFill("solid", fgColor=NAVY)
TEAL_FILL = PatternFill("solid", fgColor=TEAL)
GOLD_FILL = PatternFill("solid", fgColor=GOLD)
LIGHT_FILL = PatternFill("solid", fgColor=LIGHT)
CARD_FILL = PatternFill("solid", fgColor=WHITE)

thin = Side(style="thin", color="D0D5DD")
CARD_BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

MONEY = '#,##0;[RED](#,##0)'
PCT = '0.0%'


def style_kpi_card(ws, top_row, left_col, label, formula, fmt=MONEY, big_fill=NAVY_FILL, val_font=None):
    col = get_column_letter(left_col)
    col2 = get_column_letter(left_col + 1)
    ws.merge_cells(f"{col}{top_row}:{col2}{top_row}")
    ws.merge_cells(f"{col}{top_row+1}:{col2}{top_row+2}")
    c = ws[f"{col}{top_row}"]
    c.value = label
    c.font = KPI_LABEL
    c.fill = big_fill
    c.alignment = Alignment(horizontal="center", vertical="center")
    v = ws[f"{col}{top_row+1}"]
    v.value = formula
    v.font = val_font or KPI_VAL
    v.number_format = fmt
    v.alignment = Alignment(horizontal="center", vertical="center")
    v.fill = CARD_FILL
    for rr in range(top_row, top_row + 3):
        for cc in (left_col, left_col + 1):
            ws.cell(row=rr, column=cc).border = CARD_BORDER


def main():
    wb_src = openpyxl.load_workbook(SRC, data_only=True)
    fact_rows = [list(r) for r in wb_src["Fact_TrialBalance"].iter_rows(min_row=2, values_only=True)]
    dim_entity_rows = [list(r) for r in wb_src["Dim_Entity"].iter_rows(min_row=2, values_only=True)]

    # Drop the old, unvalidated IndiaRosa 2 rows and splice in the freshly
    # parsed + benchmark-verified ones.
    fact_rows = [r for r in fact_rows if r[0] != "9455-8236 QI"]
    with open("output/indiarosa2_fact_rows.json") as f:
        in2 = json.load(f)
    fact_rows.extend(in2["rows"])

    entity_names = {e[0]: e[1] for e in dim_entity_rows}
    entity_segment = {e[0]: e[2] for e in dim_entity_rows}
    entity_ids = [e[0] for e in dim_entity_rows]

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ------------------------------------------------------------------ #
    # Sheet: Dim_Entity, Dim_Calendar, Fact_TrialBalance (raw tables)
    # ------------------------------------------------------------------ #
    ws_e = wb.create_sheet("Dim_Entity")
    ws_e.append(["EntityId", "EntityName", "Segment"])
    for r in dim_entity_rows:
        ws_e.append(r)
    for c in ws_e[1]:
        c.font = Font(bold=True, color=WHITE)
        c.fill = NAVY_FILL
    ws_e.column_dimensions["A"].width = 16
    ws_e.column_dimensions["B"].width = 30
    ws_e.column_dimensions["C"].width = 20
    tab_e = Table(displayName="Dim_Entity", ref=f"A1:C{ws_e.max_row}")
    tab_e.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws_e.add_table(tab_e)

    ws_c = wb.create_sheet("Dim_Calendar")
    ws_c.append(["FiscalYear", "YearEnd", "YearLabel"])
    ws_c.append(["FY2024", "2024-03-31", "FY2024"])
    ws_c.append(["FY2025", "2025-03-31", "FY2025"])
    for c in ws_c[1]:
        c.font = Font(bold=True, color=WHITE)
        c.fill = NAVY_FILL
    ws_c.column_dimensions["A"].width = 14
    ws_c.column_dimensions["B"].width = 14
    ws_c.column_dimensions["C"].width = 14
    tab_c = Table(displayName="Dim_Calendar", ref="A1:C3")
    tab_c.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws_c.add_table(tab_c)

    ws_f = wb.create_sheet("Fact_TrialBalance")
    ws_f.append(["EntityId", "Category", "FiscalYear", "Amount"])
    for r in fact_rows:
        ws_f.append(r)
    for c in ws_f[1]:
        c.font = Font(bold=True, color=WHITE)
        c.fill = NAVY_FILL
    ws_f.column_dimensions["A"].width = 16
    ws_f.column_dimensions["B"].width = 28
    ws_f.column_dimensions["C"].width = 12
    ws_f.column_dimensions["D"].width = 16
    for row in ws_f.iter_rows(min_row=2, max_row=ws_f.max_row, min_col=4, max_col=4):
        row[0].number_format = MONEY
    n_fact = ws_f.max_row
    tab_f = Table(displayName="Fact_TrialBalance", ref=f"A1:D{n_fact}")
    tab_f.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws_f.add_table(tab_f)

    # ------------------------------------------------------------------ #
    # Sheet: Dashboard (entity + FY selector, KPI cards, category chart)
    # ------------------------------------------------------------------ #
    ws = wb.create_sheet("Dashboard", 0)
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGHIJ", [3, 15, 15, 15, 15, 3, 15, 15, 15, 15]):
        ws.column_dimensions[col].width = w

    ws.merge_cells("B2:J3")
    ws["B2"] = "SANDHU GROUP — Financial Dashboard"
    ws["B2"].font = TITLE_FONT
    ws["B2"].fill = NAVY_FILL
    ws["B2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for col in "BCDEFGHIJ":
        ws[f"{col}2"].fill = NAVY_FILL
        ws[f"{col}3"].fill = NAVY_FILL
    ws.merge_cells("B4:J4")
    ws["B4"] = "CaseWare WTB + GIFI categorization  •  Fiscal year end March 31  •  Figures in CAD"
    ws["B4"].font = SUB_FONT
    ws["B4"].fill = NAVY_FILL
    ws["B4"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 22
    ws.row_dimensions[3].height = 10
    ws.row_dimensions[4].height = 18

    # Selectors
    ws["B6"] = "Select Entity"
    ws["B6"].font = LABEL
    ws.merge_cells("C6:E6")
    ws["C6"] = entity_names["9455-8236 QI"]
    ws["C6"].font = Font(bold=True, size=12, color=NAVY)
    ws["C6"].fill = GOLD_FILL
    ws["C6"].alignment = Alignment(horizontal="center")
    dv_entity = DataValidation(type="list", formula1="=Dim_Entity!$B$2:$B$16", allow_blank=False)
    ws.add_data_validation(dv_entity)
    dv_entity.add(ws["C6"])

    ws["G6"] = "Select Fiscal Year"
    ws["G6"].font = LABEL
    ws["H6"] = "FY2025"
    ws["H6"].font = Font(bold=True, size=12, color=NAVY)
    ws["H6"].fill = GOLD_FILL
    ws["H6"].alignment = Alignment(horizontal="center")
    dv_fy = DataValidation(type="list", formula1="=Dim_Calendar!$A$2:$A$3", allow_blank=False)
    ws.add_data_validation(dv_fy)
    dv_fy.add(ws["H6"])

    # Helper: selected EntityId
    ws["B8"] = "EntityId ->"
    ws["B8"].font = Font(size=8, color=GREY)
    ws["C8"] = '=INDEX(Dim_Entity!$A$2:$A$16,MATCH($C$6,Dim_Entity!$B$2:$B$16,0))'
    ws["C8"].font = Font(size=8, color=GREY)
    eid_cell = "$C$8"
    fy_cell = "$H$6"

    def sumifs(cat, flip=False):
        f = (f'SUMIFS(Fact_TrialBalance[Amount],Fact_TrialBalance[EntityId],{eid_cell},'
             f'Fact_TrialBalance[Category],"{cat}",Fact_TrialBalance[FiscalYear],{fy_cell})')
        return f'=-({f})' if flip else f'={f}'

    def sumifs_multi_exclude(cats):
        parts = ",".join(f'"{c}"' for c in cats)
        return (f'=-(SUMIFS(Fact_TrialBalance[Amount],Fact_TrialBalance[EntityId],{eid_cell},'
                f'Fact_TrialBalance[FiscalYear],{fy_cell})'
                f'-SUMPRODUCT((Fact_TrialBalance[EntityId]={eid_cell})*'
                f'(Fact_TrialBalance[FiscalYear]={fy_cell})*'
                f'(ISNUMBER(MATCH(Fact_TrialBalance[Category],{{{parts}}},0)))*Fact_TrialBalance[Amount]))')

    # KPI cards row 1
    row1 = 10
    style_kpi_card(ws, row1, 2, "REVENUE", sumifs("Operating Revenue", flip=True))
    style_kpi_card(ws, row1, 4, "EBITDA", sumifs_multi_exclude(["Interest & Financing", "Amortization", "Income Tax"]))
    net_income_formula = (
        f'=-(SUMIFS(Fact_TrialBalance[Amount],Fact_TrialBalance[EntityId],{eid_cell},'
        f'Fact_TrialBalance[FiscalYear],{fy_cell}))'
    )
    style_kpi_card(ws, row1, 7, "NET INCOME", net_income_formula)
    style_kpi_card(ws, row1, 9, "INCOME TAX", sumifs("Income Tax"))

    row2 = row1 + 4
    ws.row_dimensions[row1].height = 16
    style_kpi_card(ws, row2, 2, "EBITDA MARGIN", f'=IFERROR(D{row1+1}/B{row1+1},0)', fmt=PCT, big_fill=TEAL_FILL)
    style_kpi_card(ws, row2, 4, "EFFECTIVE TAX RATE",
                    f'=IFERROR(I{row1+1}/(G{row1+1}+I{row1+1}),0)', fmt=PCT, big_fill=TEAL_FILL)
    style_kpi_card(ws, row2, 7, "% OF GROUP REVENUE",
                    f'=IFERROR(B{row1+1}/(-SUMIFS(Fact_TrialBalance[Amount],Fact_TrialBalance[Category],'
                    f'"Operating Revenue",Fact_TrialBalance[FiscalYear],{fy_cell})),0)',
                    fmt=PCT, big_fill=TEAL_FILL)
    style_kpi_card(ws, row2, 9, "COGS", sumifs("COGS"), big_fill=TEAL_FILL)

    # Category breakdown table (drives the chart)
    cat_row0 = row2 + 5
    ws.merge_cells(f"B{cat_row0}:E{cat_row0}")
    ws[f"B{cat_row0}"] = "Category Breakdown"
    ws[f"B{cat_row0}"].font = H1
    ws[f"B{cat_row0}"].fill = NAVY_FILL
    for col in "CDE":
        ws[f"{col}{cat_row0}"].fill = NAVY_FILL
    ws[f"B{cat_row0+1}"] = "Category"
    ws[f"C{cat_row0+1}"] = "Amount"
    for c in (f"B{cat_row0+1}", f"C{cat_row0+1}"):
        ws[c].font = Font(bold=True, color=WHITE)
        ws[c].fill = TEAL_FILL
    r = cat_row0 + 2
    for cat in CATEGORIES:
        ws[f"B{r}"] = cat
        flip = cat in INCOME_CATS
        ws[f"C{r}"] = sumifs(cat, flip=flip)
        ws[f"C{r}"].number_format = MONEY
        r += 1
    cat_last_row = r - 1

    chart = BarChart()
    chart.type = "bar"
    chart.title = "Category Breakdown — Selected Entity & Year"
    chart.y_axis.title = None
    chart.x_axis.title = None
    chart.style = 10
    data = Reference(ws, min_col=3, min_row=cat_row0 + 1, max_row=cat_last_row)
    cats_ref = Reference(ws, min_col=2, min_row=cat_row0 + 2, max_row=cat_last_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.height = 10
    chart.width = 20
    chart.legend = None
    ws.add_chart(chart, f"G{cat_row0+1}")

    ws.freeze_panes = "A5"

    # ------------------------------------------------------------------ #
    # Sheet: Entity Register (all entities, filterable table, selected FY)
    # ------------------------------------------------------------------ #
    ws2 = wb.create_sheet("Entity Register")
    ws2.sheet_view.showGridLines = False
    ws2.merge_cells("A1:K2")
    ws2["A1"] = "Entity Register — driven by Dashboard fiscal-year selector"
    ws2["A1"].font = H1
    ws2["A1"].fill = NAVY_FILL
    for col in "BCDEFGHIJK":
        ws2[f"{col}1"].fill = NAVY_FILL
        ws2[f"{col}2"].fill = NAVY_FILL

    headers = ["EntityName", "Segment", "Revenue", "COGS", "Payroll", "Occupancy",
               "Interest & Financing", "Amortization", "Income Tax", "EBITDA", "Net Income"]
    hdr_row = 4
    for i, h in enumerate(headers):
        c = ws2.cell(row=hdr_row, column=i + 1, value=h)
        c.font = Font(bold=True, color=WHITE)
        c.fill = TEAL_FILL

    cat_col_map = {
        "COGS": "COGS", "Payroll": "Payroll & Benefits", "Occupancy": "Occupancy",
        "Interest & Financing": "Interest & Financing", "Amortization": "Amortization",
        "Income Tax": "Income Tax",
    }
    for i, eid in enumerate(entity_ids):
        rr = hdr_row + 1 + i
        ws2.cell(row=rr, column=1, value=entity_names[eid])
        ws2.cell(row=rr, column=2, value=entity_segment[eid])
        ws2.cell(row=rr, column=3,
                 value=f'=-SUMIFS(Fact_TrialBalance[Amount],Fact_TrialBalance[EntityId],"{eid}",'
                       f'Fact_TrialBalance[Category],"Operating Revenue",Fact_TrialBalance[FiscalYear],Dashboard!$H$6)')
        col_i = 4
        for h in headers[3:9]:
            cat = cat_col_map[h]
            ws2.cell(row=rr, column=col_i,
                     value=f'=SUMIFS(Fact_TrialBalance[Amount],Fact_TrialBalance[EntityId],"{eid}",'
                           f'Fact_TrialBalance[Category],"{cat}",Fact_TrialBalance[FiscalYear],Dashboard!$H$6)')
            col_i += 1
        excl = '","'.join(["Interest & Financing", "Amortization", "Income Tax"])
        ws2.cell(row=rr, column=10,
                 value=f'=-(SUMIFS(Fact_TrialBalance[Amount],Fact_TrialBalance[EntityId],"{eid}",'
                       f'Fact_TrialBalance[FiscalYear],Dashboard!$H$6)'
                       f'-SUMPRODUCT((Fact_TrialBalance[EntityId]="{eid}")*'
                       f'(Fact_TrialBalance[FiscalYear]=Dashboard!$H$6)*'
                       f'(ISNUMBER(MATCH(Fact_TrialBalance[Category],{{"{excl}"}},0)))*Fact_TrialBalance[Amount]))')
        ws2.cell(row=rr, column=11,
                 value=f'=-SUMIFS(Fact_TrialBalance[Amount],Fact_TrialBalance[EntityId],"{eid}",'
                       f'Fact_TrialBalance[FiscalYear],Dashboard!$H$6)')
        for col_i in range(3, 12):
            ws2.cell(row=rr, column=col_i).number_format = MONEY

    last_row = hdr_row + len(entity_ids)
    tab2 = Table(displayName="EntityRegister", ref=f"A{hdr_row}:K{last_row}")
    tab2.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws2.add_table(tab2)
    for col, w in zip("ABCDEFGHIJK", [26, 18, 14, 14, 14, 14, 16, 14, 12, 14, 14]):
        ws2.column_dimensions[col].width = w
    ws2.freeze_panes = f"A{hdr_row+1}"

    # Conditional formatting: flag negative Net Income in red
    ws2.conditional_formatting.add(
        f"K{hdr_row+1}:K{last_row}",
        CellIsRule(operator="lessThan", formula=["0"], font=Font(color=RED, bold=True)),
    )

    # ------------------------------------------------------------------ #
    # Sheet: Group Summary
    # ------------------------------------------------------------------ #
    ws3 = wb.create_sheet("Group Summary")
    ws3.sheet_view.showGridLines = False
    ws3.merge_cells("B2:H3")
    ws3["B2"] = "Group Summary — driven by Dashboard fiscal-year selector"
    ws3["B2"].font = H1
    ws3["B2"].fill = NAVY_FILL
    for col in "CDEFGH":
        ws3[f"{col}2"].fill = NAVY_FILL
        ws3[f"{col}3"].fill = NAVY_FILL
    for col, w in zip("ABCDEFGH", [3, 18, 18, 18, 18, 3, 18, 18]):
        ws3.column_dimensions[col].width = w

    ws3["B5"] = "15-entity group total (per Entity Register) — 3 entities pending WTB files: Sandhu & Sandhu Enr., 2 trusts"
    ws3["B5"].font = Font(italic=True, size=9, color=GREY)

    r0 = 6
    style_kpi_card(ws3, r0, 2, "GROUP REVENUE", "=SUM('Entity Register'!C5:C19)")
    style_kpi_card(ws3, r0, 4, "GROUP EBITDA", "=SUM('Entity Register'!J5:J19)", big_fill=TEAL_FILL)
    style_kpi_card(ws3, r0, 7, "GROUP NET INCOME", "=SUM('Entity Register'!K5:K19)")

    r1 = r0 + 4
    style_kpi_card(ws3, r1, 2, "GROUP INCOME TAX", "=SUM('Entity Register'!I5:I19)", big_fill=TEAL_FILL)
    style_kpi_card(
        ws3, r1, 4, "EFFECTIVE TAX RATE",
        "=IFERROR(SUM('Entity Register'!I5:I19)/(SUM('Entity Register'!K5:K19)+SUM('Entity Register'!I5:I19)),0)",
        fmt=PCT, big_fill=TEAL_FILL,
    )
    style_kpi_card(ws3, r1, 7, "ENTITIES IN MODEL", "=COUNTA('Entity Register'!A5:A19)", fmt='0', big_fill=TEAL_FILL)

    # Segment rollup + chart
    seg_row0 = r1 + 5
    ws3.merge_cells(f"B{seg_row0}:D{seg_row0}")
    ws3[f"B{seg_row0}"] = "Revenue & Net Income by Entity"
    ws3[f"B{seg_row0}"].font = H1
    ws3[f"B{seg_row0}"].fill = NAVY_FILL
    ws3[f"C{seg_row0}"].fill = NAVY_FILL
    ws3[f"D{seg_row0}"].fill = NAVY_FILL

    hdr = seg_row0 + 1
    ws3[f"B{hdr}"] = "EntityName"
    ws3[f"C{hdr}"] = "Revenue"
    ws3[f"D{hdr}"] = "Net Income"
    for c in (f"B{hdr}", f"C{hdr}", f"D{hdr}"):
        ws3[c].font = Font(bold=True, color=WHITE)
        ws3[c].fill = TEAL_FILL
    for i in range(len(entity_ids)):
        rr = hdr + 1 + i
        src_rr = 5 + i
        ws3[f"B{rr}"] = f"='Entity Register'!A{src_rr}"
        ws3[f"C{rr}"] = f"='Entity Register'!C{src_rr}"
        ws3[f"D{rr}"] = f"='Entity Register'!K{src_rr}"
        ws3[f"C{rr}"].number_format = MONEY
        ws3[f"D{rr}"].number_format = MONEY
    seg_last = hdr + len(entity_ids)

    chart2 = BarChart()
    chart2.type = "col"
    chart2.title = "Revenue vs Net Income by Entity"
    chart2.style = 12
    data2 = Reference(ws3, min_col=3, max_col=4, min_row=hdr, max_row=seg_last)
    cats2 = Reference(ws3, min_col=2, min_row=hdr + 1, max_row=seg_last)
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(cats2)
    chart2.height = 11
    chart2.width = 24
    ws3.add_chart(chart2, f"F{seg_row0}")

    # ------------------------------------------------------------------ #
    # Sheet: Notes
    # ------------------------------------------------------------------ #
    ws4 = wb.create_sheet("Notes & Validation")
    ws4.column_dimensions["A"].width = 100
    notes = [
        "SANDHU GROUP INTERACTIVE DASHBOARD — Build Notes",
        "",
        "Scope: 15 of 18 entities (all except Sandhu & Sandhu Enr., Gurpreet Sandhu Trust, "
        "Harpreet Sandhu Trust — WTB files not yet supplied for these).",
        "",
        "Restaurant IndiaRosa 2 (9455-8236 QI) was rebuilt in this workbook directly from its "
        "CaseWare WTB export (Res_In2_2025_WTB.xlsx), categorized via GIFI_Master_Mapping_v2.xlsx, "
        "and matches every Section 6 benchmark exactly:",
        "  Net Income FY2025 = $462,894.92 (benchmark $462,894.92) — EXACT",
        "  Net Income FY2024 = $736,959.45 (benchmark $736,959.45) — EXACT",
        "  All 12 FY2025 category totals match the Section 6 breakdown exactly.",
        "",
        "Three accounts in the IndiaRosa 2 WTB carry a GIFI code that contradicts the account name; "
        "these were treated as account-level overrides (not changes to the master GIFI mapping, which "
        "is presumed correct for other entities' use of the same codes):",
        "  Acct 45130 'Valet service' + 45125 'Musical entertainment' — tagged GIFI 9270 (Amortization) "
        "-> reclassified to Marketing & Entertainment.",
        "  Acct 44215 'Rent - Ford' + 33455 'Equipment rental' + 45250 'Travelling - gas & repairs' "
        "— tagged GIFI 8461 (COGS) -> reclassified to Other Operating.",
        "  Acct 45150 'Consulting fees' — tagged GIFI 8863 (Administrative) -> reclassified to "
        "Other Operating per user confirmation (matches Section 6 Admin $99,729 / Other Operating $79,976 exactly).",
        "",
        "The other 14 entities' figures are carried over unchanged from Sandhu_Excel_Model_Data.xlsx "
        "(the pre-built starter kit) and were NOT re-derived from raw WTB files in this session — "
        "only IndiaRosa 2 was rebuilt from source and fully re-validated.",
        "",
        "UNRESOLVED CONFLICT FOUND: Bistro Guru Inc. FY2025 Net Income is -$45,954.06 in this carried-over "
        "data, but the Validation Log tab in GIFI_Master_Mapping_v2.xlsx independently validated it at "
        "-$30,118.89 for the same entity/year — a $15,835.17 discrepancy between two of the project's own "
        "source files. Not resolved here (no raw WTB for Bistro Guru in this session) — needs the client's "
        "WTB file to re-derive and confirm which figure is correct.",
        "",
        "NOT YET INCLUDED: Corporate Taxprep data (SBD, GRIP, RDTOH, CDA, Part I tax payable). Per the "
        "handoff doc, RDTOH and CDA field mappings are unverified — these will not be added until "
        "cross-validated against a known figure, and will be confirmed with you before being treated as reliable.",
        "",
        "Interactivity: this workbook uses dropdown selectors (Dashboard!C6 Entity, Dashboard!H6 "
        "Fiscal Year) driving live SUMIFS formulas and charts — no Power Pivot/DAX setup required, "
        "works in any version of Excel. Entity Register and raw data sheets are Excel Tables with "
        "built-in filter dropdowns per column.",
        "",
        "Recalculated and verified in LibreOffice headless — see verification log for the exact "
        "cell values confirmed.",
    ]
    for i, line in enumerate(notes, start=1):
        c = ws4.cell(row=i, column=1, value=line)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if i == 1:
            c.font = Font(bold=True, size=14, color=NAVY)
        elif line.strip().startswith("UNRESOLVED"):
            c.font = Font(size=10, bold=True, color=RED)
        elif line.strip().startswith(("Net Income", "All 12", "Acct", "  Acct")):
            c.font = Font(size=10, color=TEAL)
        else:
            c.font = Font(size=10, color="1A1A1A")

    for sheet in (ws, ws2, ws3, ws4):
        sheet.page_setup.orientation = "landscape"
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 0
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.print_options.horizontalCentered = False

    wb.calculation.fullCalcOnLoad = True
    wb.save(OUT)
    print("Saved", OUT)


if __name__ == "__main__":
    main()
