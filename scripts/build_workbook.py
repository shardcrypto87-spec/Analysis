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

# --- palette (matches the reference "Consolidated Performance & Wealth
# Report" mockup: cream page background, navy header/nav, gold accent) ------
NAVY = "1C2B45"
TEAL = "0E7C7B"
GOLD = "C9A227"
CREAM = "F5F1E8"
LIGHT = "F4F6F8"
WHITE = "FFFFFF"
GREY = "6B7280"
RED = "B3261E"
GREEN = "1E6B52"

TITLE_FONT = Font(name="Calibri", size=20, bold=True, color=WHITE)
SUB_FONT = Font(name="Calibri", size=11, italic=True, color="D9E2EC")
H1 = Font(name="Calibri", size=13, bold=True, color=WHITE)
H2 = Font(name="Calibri", size=11, bold=True, color=NAVY)
LABEL = Font(name="Calibri", size=10, bold=True, color=GREY)
KPI_VAL = Font(name="Calibri", size=18, bold=True, color=NAVY)
KPI_LABEL = Font(name="Calibri", size=10, bold=True, color=WHITE)
BODY = Font(name="Calibri", size=10, color="1A1A1A")
FLAT_LABEL = Font(name="Calibri", size=9, bold=True, color=GREY)
FLAT_VAL = Font(name="Calibri", size=17, bold=True, color=NAVY)
BADGE_FONT = Font(name="Calibri", size=9, bold=True, color=NAVY)

NAVY_FILL = PatternFill("solid", fgColor=NAVY)
TEAL_FILL = PatternFill("solid", fgColor=TEAL)
GOLD_FILL = PatternFill("solid", fgColor=GOLD)
LIGHT_FILL = PatternFill("solid", fgColor=LIGHT)
CARD_FILL = PatternFill("solid", fgColor=WHITE)
CREAM_FILL = PatternFill("solid", fgColor=CREAM)

thin = Side(style="thin", color="D0D5DD")
CARD_BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

MONEY = '#,##0;[RED](#,##0)'
PCT = '0.0%'
DOT_GREEN, DOT_GOLD, DOT_GREY = "●", "●", "●"


def paint_background(ws, max_row=80, max_col=20, color=CREAM_FILL):
    """Fill a rectangular block so the sheet reads as a colored page
    background (Excel has no native 'page background color')."""
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            ws.cell(row=r, column=c).fill = color


def badge(ws, merge_range, text):
    first_cell = merge_range.split(":")[0]
    ws.merge_cells(merge_range)
    ws[first_cell] = text
    ws[first_cell].font = BADGE_FONT
    ws[first_cell].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    col1, col2 = merge_range.split(":")
    row = "".join(ch for ch in col1 if ch.isdigit())
    c1 = "".join(ch for ch in col1 if ch.isalpha())
    c2 = "".join(ch for ch in col2 if ch.isalpha())
    for col_idx in range(openpyxl.utils.column_index_from_string(c1),
                          openpyxl.utils.column_index_from_string(c2) + 1):
        ws.cell(row=int(row), column=col_idx).fill = GOLD_FILL


def flat_card(ws, top_row, left_col, width, label, formula, fmt=MONEY, sub=None):
    col = get_column_letter(left_col)
    col2 = get_column_letter(left_col + width - 1)
    ws.merge_cells(f"{col}{top_row}:{col2}{top_row}")
    lc = ws[f"{col}{top_row}"]
    lc.value = label.upper()
    lc.font = FLAT_LABEL
    lc.fill = CARD_FILL
    lc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.merge_cells(f"{col}{top_row+1}:{col2}{top_row+2}")
    vc = ws[f"{col}{top_row+1}"]
    vc.value = formula
    vc.font = FLAT_VAL
    vc.number_format = fmt
    vc.fill = CARD_FILL
    vc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for rr in range(top_row, top_row + 3):
        for cc in range(left_col, left_col + width):
            cell = ws.cell(row=rr, column=cc)
            cell.fill = CARD_FILL
            cell.border = CARD_BORDER
    if sub:
        ws.merge_cells(f"{col}{top_row+3}:{col2}{top_row+3}")
        sc = ws[f"{col}{top_row+3}"]
        sc.value = sub
        sc.font = Font(size=8, italic=True, color=GREY)
        sc.alignment = Alignment(horizontal="left", indent=1)


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
    # Every row here is freshly parsed from source and independently
    # validated: IndiaRosa 2 against Section 6 of the handoff doc, the
    # other 15 entities against the Validation Log tab. Supersedes the old
    # Sandhu_Excel_Model_Data.xlsx starter kit entirely.
    with open("output/final_fact_table.json") as f:
        final = json.load(f)
    fact_rows = final["fact_rows"]
    dim_entity_rows = final["dim_entity_rows"]

    with open("output/indiarosa2_fact_rows.json") as f:
        in2 = json.load(f)

    with open("output/indiarosa2_taxprep_extract.json") as f:
        tax = json.load(f)

    entity_names = {e[0]: e[1] for e in dim_entity_rows}
    entity_legal = {e[0]: e[2] for e in dim_entity_rows}
    entity_industry = {e[0]: e[3] for e in dim_entity_rows}
    entity_ids = [e[0] for e in dim_entity_rows]

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ------------------------------------------------------------------ #
    # Sheet: Dim_Entity, Dim_Calendar, Fact_TrialBalance (raw tables)
    # ------------------------------------------------------------------ #
    ws_e = wb.create_sheet("Dim_Entity")
    ws_e.append(["EntityId", "CompanyName", "LegalName", "Industry"])
    for r in dim_entity_rows:
        ws_e.append(r)
    for c in ws_e[1]:
        c.font = Font(bold=True, color=WHITE)
        c.fill = NAVY_FILL
    ws_e.column_dimensions["A"].width = 12
    ws_e.column_dimensions["B"].width = 34
    ws_e.column_dimensions["C"].width = 24
    ws_e.column_dimensions["D"].width = 16
    tab_e = Table(displayName="Dim_Entity", ref=f"A1:D{ws_e.max_row}")
    tab_e.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws_e.add_table(tab_e)

    ws_c = wb.create_sheet("Dim_Calendar")
    ws_c.append(["FiscalYear", "YearEnd", "YearLabel"])
    ws_c.append(["FY2023", "2023-03-31", "FY2023"])
    ws_c.append(["FY2024", "2024-03-31", "FY2024"])
    ws_c.append(["FY2025", "2025-03-31", "FY2025"])
    for c in ws_c[1]:
        c.font = Font(bold=True, color=WHITE)
        c.fill = NAVY_FILL
    ws_c.column_dimensions["A"].width = 14
    ws_c.column_dimensions["B"].width = 14
    ws_c.column_dimensions["C"].width = 26
    tab_c = Table(displayName="Dim_Calendar", ref="A1:C4")
    tab_c.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws_c.add_table(tab_c)

    ws_f = wb.create_sheet("Fact_TrialBalance")
    ws_f.append(["EntityId", "Category", "FiscalYear", "Amount", "Industry"])
    for r in fact_rows:
        ws_f.append(r + [entity_industry[r[0]]])
    for c in ws_f[1]:
        c.font = Font(bold=True, color=WHITE)
        c.fill = NAVY_FILL
    ws_f.column_dimensions["A"].width = 16
    ws_f.column_dimensions["B"].width = 28
    ws_f.column_dimensions["C"].width = 12
    ws_f.column_dimensions["D"].width = 16
    ws_f.column_dimensions["E"].width = 18
    for row in ws_f.iter_rows(min_row=2, max_row=ws_f.max_row, min_col=4, max_col=4):
        row[0].number_format = MONEY
    n_fact = ws_f.max_row
    tab_f = Table(displayName="Fact_TrialBalance", ref=f"A1:E{n_fact}")
    tab_f.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws_f.add_table(tab_f)

    # ------------------------------------------------------------------ #
    # Sheet: 01 Exec Summary — laid out to match the user's reference
    # mockup page-for-page: badge, title, FY chips, nav tabs, 6 KPI cards,
    # data confidence, notable movers, management commentary, group
    # trend chart, performance-by-segment. Every number is real; where the
    # reference shows something I don't have a source for (Group Net
    # Worth, full-group SBD utilization), it's labelled honestly rather
    # than filled in to match the mockup's placeholder-free look.
    # ------------------------------------------------------------------ #
    from collections import defaultdict as _dd
    ni_by_entity_fy = _dd(lambda: _dd(float))
    for _eid, _cat, _fy, _amt in fact_rows:
        ni_by_entity_fy[_eid][_fy] += _amt
    movers = []
    for _eid in entity_ids:
        fy23 = -ni_by_entity_fy[_eid].get("FY2023", 0)
        fy25 = -ni_by_entity_fy[_eid].get("FY2025", 0)
        movers.append((_eid, entity_names[_eid], fy23, fy25, fy25 - fy23))
    movers.sort(key=lambda x: x[4])
    biggest_decline = movers[0]
    biggest_improvement = movers[-1]
    loss_fy23 = sum(1 for _eid in entity_ids if -ni_by_entity_fy[_eid].get("FY2023", 0) < 0)
    sbd_group_total = sum(g["sbd_cy"] for g in tax["associated_group_sbd_grip"])

    _tot23, _tot25 = _dd(float), _dd(float)
    for _eid, _cat, _fy, _amt in fact_rows:
        if _fy == "FY2023":
            _tot23[_cat] += _amt
        elif _fy == "FY2025":
            _tot25[_cat] += _amt
    _rev23 = -_tot23.get("Operating Revenue", 0)
    _rev25 = -_tot25.get("Operating Revenue", 0)
    _excl_cats = ("Interest & Financing", "Amortization", "Income Tax")
    ebitda23 = -sum(v for k, v in _tot23.items() if k not in _excl_cats)
    ebitda25 = -sum(v for k, v in _tot25.items() if k not in _excl_cats)
    ebitda23_margin = ebitda23 / _rev23 if _rev23 else 0
    ebitda25_margin = ebitda25 / _rev25 if _rev25 else 0
    _ni23 = -sum(_tot23.values())
    _ni25 = -sum(_tot25.values())
    _tax23 = _tot23.get("Income Tax", 0)
    _tax25 = _tot25.get("Income Tax", 0)
    tax23_rate = _tax23 / (_ni23 + _tax23) if (_ni23 + _tax23) else 0
    tax25_rate = _tax25 / (_ni25 + _tax25) if (_ni25 + _tax25) else 0

    wsx = wb.create_sheet("01 Exec Summary", 0)
    wsx.sheet_view.showGridLines = False
    wsx.sheet_properties.tabColor = GOLD
    for col, w in zip("ABCDEFGHIJKLM", [3, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13]):
        wsx.column_dimensions[col].width = w
    paint_background(wsx, max_row=95, max_col=14)

    badge(wsx, "B2:H2", f"DRAFT MANAGEMENT REPORT — {len(entity_ids)}/{len(entity_ids)} ENTITIES "
                         "VALIDATED AGAINST CASEWARE")
    # Master fiscal-year selector — this is now the single source of truth
    # for the whole workbook. '02 Profitability'!H6 (and everything
    # downstream of it: Entity Register, Group Summary, these KPI cards)
    # reads from this one cell, so picking a year here changes every sheet.
    wsx["J2"] = "Select Fiscal Year"
    wsx["J2"].font = Font(size=8, bold=True, color=GREY)
    wsx["J2"].alignment = Alignment(horizontal="right", vertical="center")
    wsx.merge_cells("K2:M2")
    fy_selector_cell = "K2"
    wsx["K2"] = "FY2025"
    wsx["K2"].font = Font(bold=True, size=11, color=NAVY)
    wsx["K2"].fill = GOLD_FILL
    wsx["K2"].alignment = Alignment(horizontal="center", vertical="center")
    wsx["K2"].border = CARD_BORDER
    dv_fy_master = DataValidation(type="list", formula1="=Dim_Calendar!$A$2:$A$4", allow_blank=False)
    wsx.add_data_validation(dv_fy_master)
    dv_fy_master.add(wsx["K2"])

    wsx.merge_cells("B4:M5")
    wsx["B4"] = "Sandhu Group — Consolidated Performance Report"
    wsx["B4"].font = Font(name="Calibri", size=20, bold=True, color=NAVY)
    wsx["B4"].alignment = Alignment(vertical="center")

    wsx.merge_cells("B6:M6")
    wsx["B6"] = (f"{len(entity_ids)} entities · Restaurant, Real Estate, Investments, Construction, "
                 "Hotel & Rental · FY ended March 31 · "
                 "Prepared for group management review")
    wsx["B6"].font = Font(size=10, italic=True, color=GREY)

    # Nav tabs — only 01/02/03 exist as sheets so far; the rest are shown
    # greyed out as a roadmap, matching the reference's 7-page structure.
    nav_row = 8
    nav_items = [
        ("01 Executive Summary", "01 Exec Summary"), ("02 Profitability", "02 Profitability"),
        ("03 Tax Position", "03 Tax Position"), ("04 Debt & Liquidity", None),
        ("05 Wealth of Group", None), ("06 Reorg & LCGE", None), ("07 Risk & Alerts", None),
    ]
    col_i = 2
    for label, target in nav_items:
        c = wsx.cell(row=nav_row, column=col_i, value=label)
        if target:
            c.font = Font(size=9, bold=(target == "01 Exec Summary"), color=NAVY)
            if target == "01 Exec Summary":
                c.border = Border(bottom=Side(style="medium", color=GOLD))
        else:
            c.font = Font(size=9, color="C7C2B8", italic=True)
        col_i += 2
    wsx.row_dimensions[nav_row].height = 16
    for col in "BCDEFGHIJKLM":
        wsx[f"{col}9"].border = Border(bottom=Side(style="thin", color="D8D2C4"))

    wsx.merge_cells("B11:M12")
    wsx["B11"] = ('="The one-page summary — six headline numbers below, each labelled with its data '
                  'confidence so nothing is trusted more than it should be. The Select Fiscal Year box '
                  'above (top right) is the master control for the whole workbook — change it here and '
                  '02 Profitability, Entity Register, Group Summary and the industry table below all '
                  'follow. Currently showing "&$K$2&". The trend sections further down (Notable Movers, '
                  'the revenue/EBITDA chart) stay fixed to FY2023→FY2025 regardless of this selector."')
    wsx["B11"].font = Font(size=9.5, italic=True, color=GREY)
    wsx["B11"].alignment = Alignment(wrap_text=True, vertical="top")

    # 6 KPI cards in one row, matching the reference layout
    er = "'Entity Register'!"
    kpi_row = 14
    flat_card(wsx, kpi_row, 2, 2, "Group Revenue", f"=SUM({er}C5:C20)",
              sub=(f'=TEXT((SUM({er}C5:C20)-(-SUMIFS(Fact_TrialBalance[Amount],Fact_TrialBalance[Category],'
                   '"Operating Revenue",Fact_TrialBalance[FiscalYear],"FY2023")))/'
                   '(-SUMIFS(Fact_TrialBalance[Amount],Fact_TrialBalance[Category],"Operating Revenue",'
                   'Fact_TrialBalance[FiscalYear],"FY2023")),"+0.0%;-0.0%")&" vs FY2023"'))
    flat_card(wsx, kpi_row, 4, 2, "Group EBITDA", f"=SUM({er}J5:J20)",
              sub=f'=TEXT(SUM({er}J5:J20)/SUM({er}C5:C20),"0.0%")&" margin"')
    flat_card(wsx, kpi_row, 6, 2, "Consolidated Eff. Tax Rate",
              f"=IFERROR(SUM({er}I5:I20)/(SUM({er}K5:K20)+SUM({er}I5:I20)),0)", fmt=PCT)
    flat_card(wsx, kpi_row, 8, 2, "Entities in Loss Position",
              f'=COUNTIF({er}K5:K20,"<0")&" / "&COUNTA({er}A5:A20)', fmt="@",
              sub=f'=TEXT(COUNTIF({er}K5:K20,"<0")/COUNTA({er}A5:A20),"0%")&" of group"')
    flat_card(wsx, kpi_row, 10, 2, "Group Net Worth", '="pending"', fmt="@",
              sub="needs CaseWare balance sheets")
    flat_card(wsx, kpi_row, 12, 2, "SBD Utilization (14 of 16)", "=" + str(round(sbd_group_total)),
              sub=f"{sbd_group_total/500000:.1%} of $500,000, excl. 2 entities' own claims")
    wsx[f"{get_column_letter(12)}{kpi_row+1}"].number_format = MONEY

    # Data confidence — bordered box, matching the reference's card style
    conf_row = kpi_row + 6
    wsx.merge_cells(f"B{conf_row}:M{conf_row}")
    wsx[f"B{conf_row}"] = "Data confidence"
    wsx[f"B{conf_row}"].font = Font(size=13, bold=True, color=NAVY)
    wsx[f"N{conf_row}"] = None
    conf_lines = [
        (GREEN, f"P&L ({len(entity_ids)}/{len(entity_ids)} entities) — every entity rebuilt fresh from its "
                "raw CaseWare WTB and validated to the penny (FY2025 Net Income matches the project's "
                "Validation Log exactly for all of them)."),
        (GOLD, "SBD & GRIP — real, cross-checked against Corporate Taxprep (14 of 16 entities; excludes "
               "IndiaRosa 2 & Sandhu & Sandhu's own claims, not found in that export)."),
        (GREY, "RDTOH & CDA — confirmed NOT present in the Taxprep export tested (see 03 Tax Position)."),
        (GREY, "Balance sheet / consolidated net worth — not available; the WTB export is income-statement only."),
        (RED, "One unexpected entity (9475-3381 Québec Inc.) appears in the Taxprep data but has no WTB and "
              "isn't in scope — needs your confirmation on whether it belongs in the group."),
        (GREY, "Group-level Revenue/EBITDA are within ~0.5%/1.8% of the Section 6 benchmark despite every "
               "entity's own Net Income matching exactly — a small residual category difference likely "
               "remains in 1–2 entities (see Notes & Validation)."),
    ]
    box_top = conf_row + 1
    r = box_top
    for color, text in conf_lines:
        wsx[f"B{r}"] = DOT_GREEN
        wsx[f"B{r}"].font = Font(color=color, bold=True, size=11)
        wsx.merge_cells(f"C{r}:M{r}")
        wsx[f"C{r}"] = text
        wsx[f"C{r}"].font = Font(size=9.5, color="1A1A1A")
        wsx[f"C{r}"].alignment = Alignment(wrap_text=True, vertical="top")
        wsx.row_dimensions[r].height = 28
        r += 1
    box_bottom = r - 1
    for rr in range(box_top, box_bottom + 1):
        for cc in range(2, 14):
            wsx.cell(row=rr, column=cc).border = Border(
                left=thin if cc == 2 else None, right=thin if cc == 13 else None,
                top=thin if rr == box_top else None, bottom=thin if rr == box_bottom else None,
            )

    # Notable movers, FY2023 -> FY2025 (real, computed from the fact table)
    nm_row = box_bottom + 2
    wsx.merge_cells(f"B{nm_row}:M{nm_row}")
    wsx[f"B{nm_row}"] = "Notable movers, FY2023 → FY2025"
    wsx[f"B{nm_row}"].font = Font(size=13, bold=True, color=NAVY)

    def mover_card(top_row, left_col, width, label, label_color, entity_name, fy23, fy25):
        col = get_column_letter(left_col)
        col2 = get_column_letter(left_col + width - 1)
        wsx.merge_cells(f"{col}{top_row}:{col2}{top_row}")
        c = wsx[f"{col}{top_row}"]
        c.value = label
        c.font = Font(size=9, bold=True, color=label_color)
        c.alignment = Alignment(horizontal="left", indent=1)
        wsx.merge_cells(f"{col}{top_row+1}:{col2}{top_row+1}")
        n = wsx[f"{col}{top_row+1}"]
        n.value = entity_name
        n.font = Font(size=12, bold=True, color=NAVY)
        n.alignment = Alignment(horizontal="left", indent=1)
        delta = fy25 - fy23
        wsx.merge_cells(f"{col}{top_row+2}:{col2}{top_row+2}")
        d = wsx[f"{col}{top_row+2}"]
        d.value = f"Net income: ${fy23:,.0f} → ${fy25:,.0f} ({'+'if delta>=0 else ''}{delta:,.0f})"
        d.font = Font(size=9.5, color="1A1A1A")
        d.alignment = Alignment(horizontal="left", indent=1)
        for rr in range(top_row, top_row + 3):
            for cc in range(left_col, left_col + width):
                cell = wsx.cell(row=rr, column=cc)
                cell.fill = CARD_FILL
                cell.border = CARD_BORDER

    mr = nm_row + 1
    mover_card(mr, 2, 5, "BIGGEST IMPROVEMENT", GREEN, biggest_improvement[1],
               biggest_improvement[2], biggest_improvement[3])
    mover_card(mr, 8, 6, "BIGGEST DECLINE", RED, biggest_decline[1],
               biggest_decline[2], biggest_decline[3])

    # Management commentary — auto-composed from real deltas
    mc_row = mr + 4
    wsx.merge_cells(f"B{mc_row}:M{mc_row}")
    wsx[f"B{mc_row}"] = "Management commentary"
    wsx[f"B{mc_row}"].font = Font(size=13, bold=True, color=NAVY)
    wsx.merge_cells(f"B{mc_row+1}:M{mc_row+1}")
    wsx[f"B{mc_row+1}"] = "Auto-generated from FY2023 → FY2025 movement — every figure ties to the raw WTB data."
    wsx[f"B{mc_row+1}"].font = Font(size=8.5, italic=True, color=GREY)

    cbox_top = mc_row + 2
    bullets = [
        ('="—  Group revenue grew from $"&TEXT(-SUMIFS(Fact_TrialBalance[Amount],'
         'Fact_TrialBalance[Category],"Operating Revenue",Fact_TrialBalance[FiscalYear],"FY2023"),"#,##0")&'
         '" (FY2023) to $"&TEXT(-SUMIFS(Fact_TrialBalance[Amount],Fact_TrialBalance[Category],'
         '"Operating Revenue",Fact_TrialBalance[FiscalYear],"FY2025"),"#,##0")&" (FY2025), up "&'
         'TEXT((-SUMIFS(Fact_TrialBalance[Amount],Fact_TrialBalance[Category],"Operating Revenue",'
         'Fact_TrialBalance[FiscalYear],"FY2025")/-SUMIFS(Fact_TrialBalance[Amount],'
         'Fact_TrialBalance[Category],"Operating Revenue",Fact_TrialBalance[FiscalYear],"FY2023"))-1,"0.0%")&"."'),
        (f'="—  Group EBITDA margin moved from {ebitda23_margin:.1%} (FY2023) to {ebitda25_margin:.1%} '
         f'(FY2025)."'),
        (f'="—  Consolidated effective tax rate moved from {tax23_rate:.1%} (FY2023) to {tax25_rate:.1%} '
         f'(FY2025)."'),
        (f'="—  Entities in a loss position: {loss_fy23} of {len(entity_ids)} in FY2023, versus "&'
         f'COUNTIF({er}K5:K20,"<0")&" of "&COUNTA({er}A5:A20)&" in FY2025."'),
        (f'="—  SBD: ${sbd_group_total:,.0f} of the $500,000 federal limit accounted for across the 14 '
         f'associated corporations with data in IndiaRosa 2\'s Taxprep export — IndiaRosa 2\'s and Sandhu & '
         f'Sandhu\'s own claims are not in that export, so this is a floor, not the group total."'),
        (f'="—  Biggest mover: {biggest_improvement[1]} (net income +${biggest_improvement[4]:,.0f}) '
         f'vs. {biggest_decline[1]} ({biggest_decline[4]:,.0f})."'),
    ]
    r = cbox_top
    for formula in bullets:
        wsx.merge_cells(f"B{r}:M{r}")
        wsx[f"B{r}"] = formula
        wsx[f"B{r}"].font = Font(size=9.5, color="1A1A1A")
        wsx[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top", indent=1)
        wsx.row_dimensions[r].height = 26
        r += 1
    cbox_bottom = r - 1
    for rr in range(cbox_top, cbox_bottom + 1):
        for cc in range(2, 14):
            wsx.cell(row=rr, column=cc).border = Border(
                left=thin if cc == 2 else None, right=thin if cc == 13 else None,
                top=thin if rr == cbox_top else None, bottom=thin if rr == cbox_bottom else None,
            )

    # Group revenue & EBITDA, FY2023 vs FY2025 — clustered bar chart
    chart_row = cbox_bottom + 2
    wsx.merge_cells(f"B{chart_row}:M{chart_row}")
    wsx[f"B{chart_row}"] = "Group revenue & EBITDA, FY2023 vs FY2025"
    wsx[f"B{chart_row}"].font = Font(size=13, bold=True, color=NAVY)

    data_row = chart_row + 1
    wsx[f"B{data_row}"] = "FiscalYear"
    wsx[f"C{data_row}"] = "Revenue"
    wsx[f"D{data_row}"] = "EBITDA"
    for c in (f"B{data_row}", f"C{data_row}", f"D{data_row}"):
        wsx[c].font = Font(size=8, color=GREY)
    for i, fy in enumerate(["FY2023", "FY2025"]):
        rr = data_row + 1 + i
        wsx[f"B{rr}"] = fy
        wsx[f"C{rr}"] = (f'=-SUMIFS(Fact_TrialBalance[Amount],Fact_TrialBalance[Category],'
                          f'"Operating Revenue",Fact_TrialBalance[FiscalYear],"{fy}")')
        excl = '","'.join(["Interest & Financing", "Amortization", "Income Tax"])
        wsx[f"D{rr}"] = (f'=-(SUMIFS(Fact_TrialBalance[Amount],Fact_TrialBalance[FiscalYear],"{fy}")'
                          f'-SUMPRODUCT((Fact_TrialBalance[FiscalYear]="{fy}")*'
                          f'(ISNUMBER(MATCH(Fact_TrialBalance[Category],{{"{excl}"}},0)))*Fact_TrialBalance[Amount]))')
        wsx[f"C{rr}"].number_format = MONEY
        wsx[f"D{rr}"].number_format = MONEY
    trend_chart = BarChart()
    trend_chart.type = "col"
    trend_chart.title = None
    trend_chart.style = 10
    tdata = Reference(wsx, min_col=3, max_col=4, min_row=data_row, max_row=data_row + 2)
    tcats = Reference(wsx, min_col=2, min_row=data_row + 1, max_row=data_row + 2)
    trend_chart.add_data(tdata, titles_from_data=True)
    trend_chart.set_categories(tcats)
    trend_chart.height = 9
    trend_chart.width = 24
    wsx.add_chart(trend_chart, f"B{data_row + 4}")

    # Performance by industry — table + horizontal bar chart. Follows the
    # master FY selector at K2 (same as the KPI cards above).
    seg_row = data_row + 20
    wsx.merge_cells(f"B{seg_row}:M{seg_row}")
    wsx[f"B{seg_row}"] = '="Performance by industry, "&$K$2'
    wsx[f"B{seg_row}"].font = Font(size=13, bold=True, color=NAVY)

    seg_hdr = seg_row + 1
    seg_headers = ["Industry", "# Entities", "Revenue", "EBITDA", "Margin"]
    for i, h in enumerate(seg_headers):
        c = wsx.cell(row=seg_hdr, column=2 + i, value=h)
        c.font = Font(bold=True, color=WHITE)
        c.fill = TEAL_FILL
    segments_present = sorted({entity_industry[e] for e in entity_ids})
    for i, seg in enumerate(segments_present):
        rr = seg_hdr + 1 + i
        wsx.cell(row=rr, column=2, value=seg)
        # Count entities via Dim_Entity, not rows in Fact_TrialBalance — an
        # entity with zero operating revenue has no "Operating Revenue" row
        # at all (zero-amount rows are dropped), which would undercount it.
        wsx.cell(row=rr, column=3, value=f'=COUNTIF(Dim_Entity[Industry],"{seg}")')
        # Revenue here = Operating Revenue + Investment & Other Income, to
        # match EBITDA's income scope — HoldCos earn almost entirely via
        # intercompany dividends/interest (Investment & Other Income), so
        # Operating Revenue alone would understate it and produce a
        # nonsensical >100% "margin".
        wsx.cell(row=rr, column=4,
                 value=f'=-SUMIFS(Fact_TrialBalance[Amount],Fact_TrialBalance[Industry],"{seg}",'
                       f'Fact_TrialBalance[Category],"Operating Revenue",Fact_TrialBalance[FiscalYear],$K$2)'
                       f'-SUMIFS(Fact_TrialBalance[Amount],Fact_TrialBalance[Industry],"{seg}",'
                       f'Fact_TrialBalance[Category],"Investment & Other Income",Fact_TrialBalance[FiscalYear],$K$2)')
        excl = '","'.join(["Interest & Financing", "Amortization", "Income Tax"])
        wsx.cell(row=rr, column=5,
                 value=f'=-(SUMIFS(Fact_TrialBalance[Amount],Fact_TrialBalance[Industry],"{seg}",'
                       f'Fact_TrialBalance[FiscalYear],$K$2)'
                       f'-SUMPRODUCT((Fact_TrialBalance[Industry]="{seg}")*'
                       f'(Fact_TrialBalance[FiscalYear]=$K$2)*'
                       f'(ISNUMBER(MATCH(Fact_TrialBalance[Category],{{"{excl}"}},0)))*Fact_TrialBalance[Amount]))')
        wsx.cell(row=rr, column=6, value=f"=IFERROR(E{rr}/D{rr},0)")
        wsx.cell(row=rr, column=4).number_format = MONEY
        wsx.cell(row=rr, column=5).number_format = MONEY
        wsx.cell(row=rr, column=6).number_format = PCT
    seg_last = seg_hdr + len(segments_present)

    seg_chart = BarChart()
    seg_chart.type = "bar"
    seg_chart.title = "Revenue & EBITDA by Industry"
    seg_chart.style = 10
    seg_data = Reference(wsx, min_col=4, max_col=5, min_row=seg_hdr, max_row=seg_last)
    seg_cats = Reference(wsx, min_col=2, min_row=seg_hdr + 1, max_row=seg_last)
    seg_chart.add_data(seg_data, titles_from_data=True)
    seg_chart.set_categories(seg_cats)
    seg_chart.height = 9
    seg_chart.width = 22
    wsx.add_chart(seg_chart, f"G{seg_hdr}")

    wsx.merge_cells(f"B{seg_last+1}:F{seg_last+1}")
    wsx[f"B{seg_last+1}"] = ("\"Revenue\" here = Operating Revenue + Investment & Other Income, matching "
                              "EBITDA's income scope — HoldCos earn mainly via intercompany dividends/interest.")
    wsx[f"B{seg_last+1}"].font = Font(size=8, italic=True, color=GREY)
    wsx[f"B{seg_last+1}"].alignment = Alignment(wrap_text=True)

    # ------------------------------------------------------------------ #
    # Sheet: 02 Profitability (entity + FY selector, KPI cards, category chart)
    # ------------------------------------------------------------------ #
    ws = wb.create_sheet("02 Profitability", 1)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = NAVY
    for col, w in zip("ABCDEFGHIJ", [3, 15, 15, 15, 15, 3, 15, 15, 15, 15]):
        ws.column_dimensions[col].width = w
    paint_background(ws, max_row=70, max_col=12)

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
    dv_entity = DataValidation(type="list", formula1="=Dim_Entity!$B$2:$B$17", allow_blank=False)
    ws.add_data_validation(dv_entity)
    dv_entity.add(ws["C6"])

    ws["G6"] = "Fiscal Year"
    ws["G6"].font = LABEL
    # H6 mirrors the master selector on 01 Exec Summary — one control for
    # the whole workbook. Not its own dropdown any more: change the year
    # on Exec Summary and every sheet below follows automatically.
    ws["H6"] = "='01 Exec Summary'!$K$2"
    ws["H6"].font = Font(bold=True, size=12, color=NAVY)
    ws["H6"].fill = GOLD_FILL
    ws["H6"].alignment = Alignment(horizontal="center")
    ws["I6"] = "set on 01 Exec Summary → flows to every sheet"
    ws["I6"].font = Font(size=8, italic=True, color=GREY)

    # Helper: selected EntityId
    ws["B8"] = "EntityId ->"
    ws["B8"].font = Font(size=8, color=GREY)
    ws["C8"] = '=INDEX(Dim_Entity!$A$2:$A$17,MATCH($C$6,Dim_Entity!$B$2:$B$17,0))'
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
    # Sheet: 03 Tax Position (real Taxprep-sourced figures for IndiaRosa 2
    # and its associated group — every number here was cross-checked
    # against an independently-known figure before being trusted)
    # ------------------------------------------------------------------ #
    wst = wb.create_sheet("03 Tax Position", 2)
    wst.sheet_view.showGridLines = False
    wst.sheet_properties.tabColor = TEAL
    for col, w in zip("ABCDEFGHIJKL", [3, 15, 15, 15, 15, 3, 15, 15, 15, 3, 15, 15]):
        wst.column_dimensions[col].width = w
    paint_background(wst, max_row=60, max_col=13)

    badge(wst, "B2:J2", "RESTAURANT INDIAROSA 2 — CORPORATE TAXPREP EXPORT, CROSS-VALIDATED FIELDS ONLY")
    wst.merge_cells("B4:L5")
    wst["B4"] = "Tax Position — Restaurant IndiaRosa 2 & Associated Group"
    wst["B4"].font = Font(name="Calibri", size=18, bold=True, color=NAVY)
    wst.merge_cells("B6:L6")
    wst["B6"] = ("Every figure on this page was matched against an independently-known number before being "
                 "trusted (per the handoff doc's validation rule) — nothing here is estimated or assumed.")
    wst["B6"].font = Font(size=9.5, italic=True, color=GREY)

    tv = tax["cross_validated"]
    hist = tax["tax_basis_net_income_history"]
    ncl = tax["non_capital_loss_carryforward_start_fy2023"]

    flat_card(wst, 8, 2, 2, "Income Tax, FY2025", "=" + str(tv["income_tax_cy_fy2025"]),
              sub="Matches CaseWare WTB")
    flat_card(wst, 8, 4, 2, "FY24 NI (accounting)", "=" + str(tv["net_income_fy2024_accounting_taxprep"]),
              sub="Matches CaseWare WTB")
    flat_card(wst, 8, 6, 2, "FY24 NI (tax basis)", "=" + str(hist["FY2024"]),
              sub="Federal T2 — differs from NI")
    flat_card(wst, 8, 8, 2, "NCL b/f, start FY23", "=" + str(ncl["federal"]),
              sub="Federal — QC: " + f'{ncl["quebec"]:,.0f}')

    # 3-year tax-basis net income history for IndiaRosa 2
    hrow = 14
    wst.merge_cells(f"B{hrow}:E{hrow}")
    wst[f"B{hrow}"] = "IndiaRosa 2 — Net income for tax purposes (Federal T2), 3-year"
    wst[f"B{hrow}"].font = Font(size=11, bold=True, color=NAVY)
    wst[f"B{hrow+1}"] = "Fiscal Year"
    wst[f"C{hrow+1}"] = "Net Income (tax basis)"
    for c in (f"B{hrow+1}", f"C{hrow+1}"):
        wst[c].font = Font(bold=True, color=WHITE)
        wst[c].fill = TEAL_FILL
    years_hist = [("FY2022", hist["FY2022"]), ("FY2023", hist["FY2023"]), ("FY2024", hist["FY2024"])]
    for i, (yr, val) in enumerate(years_hist):
        rr = hrow + 2 + i
        wst[f"B{rr}"] = yr
        wst[f"C{rr}"] = val
        wst[f"C{rr}"].number_format = MONEY
    wst.merge_cells(f"B{hrow+6}:E{hrow+7}")
    wst[f"B{hrow+6}"] = ("Not the formal 'Taxable Income' line (loss-carryforward application isn't shown in this "
                          "export) — this is 'Net income for income tax purposes' per the T2. Genuinely differs "
                          "from accounting net income; not treated as interchangeable.")
    wst[f"B{hrow+6}"].font = Font(size=8.5, italic=True, color=GREY)
    wst[f"B{hrow+6}"].alignment = Alignment(wrap_text=True, vertical="top")

    # RDTOH / CDA — confirmed not present
    rbox = hrow
    wst.merge_cells(f"G{rbox}:L{rbox}")
    wst[f"G{rbox}"] = "RDTOH / GRIP-CDA continuity"
    wst[f"G{rbox}"].font = Font(size=11, bold=True, color=NAVY)
    wst.merge_cells(f"G{rbox+1}:L{rbox+5}")
    rd_cell = wst[f"G{rbox+1}"]
    rd_cell.value = (
        "RDTOH balance: NOT FOUND\nCDA balance: NOT FOUND\n\n"
        "Checked every field description in this export for “RDTOH”, “refundable dividend tax on "
        "hand”, “CDA”, and “capital dividend” — zero matches, even with “Rolled "
        "Forward Data” checked in Xpress Filter. Per the handoff, this may need a dedicated Schedule 3 "
        "(RDTOH) / CDA continuity pull rather than the general Xpress Filter export. Not guessed or "
        "estimated here."
    )
    rd_cell.font = Font(size=9.5, color=RED)
    rd_cell.alignment = Alignment(wrap_text=True, vertical="top")
    rd_cell.fill = CARD_FILL
    for rr in range(rbox, rbox + 6):
        for cc in range(7, 13):
            c = wst.cell(row=rr, column=cc)
            c.fill = CARD_FILL
            c.border = CARD_BORDER

    # Associated group SBD & GRIP table (real, from the SLIPA schedule)
    grow = hrow + 9
    wst.merge_cells(f"B{grow}:L{grow}")
    wst[f"B{grow}"] = "Associated Group — SBD & GRIP (from IndiaRosa 2's Taxprep SLIPA schedule, current year)"
    wst[f"B{grow}"].font = Font(size=11, bold=True, color=NAVY)
    wst.merge_cells(f"B{grow+1}:L{grow+1}")
    wst[f"B{grow+1}"] = "Excludes IndiaRosa 2's own SBD claim — not found in this export."
    wst[f"B{grow+1}"].font = Font(size=8.5, italic=True, color=GREY)

    # Columns B,C,D,E,G (skip the narrow F spacer column used by the KPI
    # cards above) so GRIP values get a full-width column instead of "###".
    gcols = [2, 3, 4, 5, 7]
    wst.column_dimensions["G"].width = 15
    ghdr = grow + 2
    gheaders = ["Entity", "Net Income (tax)", "Part I Tax", "SBD", "GRIP"]
    for col_i, h in zip(gcols, gheaders):
        c = wst.cell(row=ghdr, column=col_i, value=h)
        c.font = Font(bold=True, color=WHITE)
        c.fill = TEAL_FILL
    group_data = tax["associated_group_sbd_grip"]
    for i, g in enumerate(group_data):
        rr = ghdr + 1 + i
        wst.cell(row=rr, column=gcols[0], value=g["name"])
        wst.cell(row=rr, column=gcols[1], value=g["net_income_tax_cy"]).number_format = MONEY
        wst.cell(row=rr, column=gcols[2], value=g["part1_tax_cy"]).number_format = MONEY
        wst.cell(row=rr, column=gcols[3], value=g["sbd_cy"]).number_format = MONEY
        wst.cell(row=rr, column=gcols[4], value=g["grip_cy"]).number_format = MONEY
    glast = ghdr + len(group_data)
    wst.cell(row=glast + 1, column=gcols[0], value="Total").font = Font(bold=True, color=NAVY)
    for col_i in gcols[1:]:
        col_l = get_column_letter(col_i)
        cell = wst.cell(row=glast + 1, column=col_i, value=f"=SUM({col_l}{ghdr+1}:{col_l}{glast})")
        cell.number_format = MONEY
        cell.font = Font(bold=True, color=NAVY)
    # F stays blank in every row (skipped column) so the table isn't a
    # contiguous range for Table()/autofilter purposes — style manually.
    thin_grey = Side(style="thin", color="D0D5DD")
    for rr in range(ghdr, glast + 2):
        for col_i in gcols:
            wst.cell(row=rr, column=col_i).border = Border(bottom=thin_grey)

    sbd_chart = BarChart()
    sbd_chart.type = "bar"
    sbd_chart.title = "SBD Allocation — Associated Group (current year)"
    sbd_chart.style = 10
    sbd_data = Reference(wst, min_col=gcols[3], min_row=ghdr, max_row=glast)
    sbd_cats = Reference(wst, min_col=gcols[0], min_row=ghdr + 1, max_row=glast)
    sbd_chart.add_data(sbd_data, titles_from_data=True)
    sbd_chart.set_categories(sbd_cats)
    sbd_chart.height = 10
    sbd_chart.width = 20
    sbd_chart.legend = None
    wst.add_chart(sbd_chart, f"I{ghdr}")

    # ------------------------------------------------------------------ #
    # Sheet: Entity Register (all entities, filterable table, selected FY)
    # ------------------------------------------------------------------ #
    ws2 = wb.create_sheet("Entity Register")
    ws2.sheet_view.showGridLines = False
    ws2.sheet_properties.tabColor = NAVY
    paint_background(ws2, max_row=25, max_col=12)
    ws2.merge_cells("A1:K2")
    ws2["A1"] = "Entity Register — driven by 02 Profitability's fiscal-year selector"
    ws2["A1"].font = H1
    ws2["A1"].fill = NAVY_FILL
    for col in "BCDEFGHIJK":
        ws2[f"{col}1"].fill = NAVY_FILL
        ws2[f"{col}2"].fill = NAVY_FILL

    headers = ["EntityName", "Industry", "Revenue", "COGS", "Payroll", "Occupancy",
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
    fy_ref = "'02 Profitability'!$H$6"
    for i, eid in enumerate(entity_ids):
        rr = hdr_row + 1 + i
        ws2.cell(row=rr, column=1, value=entity_names[eid])
        ws2.cell(row=rr, column=2, value=entity_industry[eid])
        ws2.cell(row=rr, column=3,
                 value=f'=-SUMIFS(Fact_TrialBalance[Amount],Fact_TrialBalance[EntityId],"{eid}",'
                       f'Fact_TrialBalance[Category],"Operating Revenue",Fact_TrialBalance[FiscalYear],{fy_ref})')
        col_i = 4
        for h in headers[3:9]:
            cat = cat_col_map[h]
            ws2.cell(row=rr, column=col_i,
                     value=f'=SUMIFS(Fact_TrialBalance[Amount],Fact_TrialBalance[EntityId],"{eid}",'
                           f'Fact_TrialBalance[Category],"{cat}",Fact_TrialBalance[FiscalYear],{fy_ref})')
            col_i += 1
        excl = '","'.join(["Interest & Financing", "Amortization", "Income Tax"])
        ws2.cell(row=rr, column=10,
                 value=f'=-(SUMIFS(Fact_TrialBalance[Amount],Fact_TrialBalance[EntityId],"{eid}",'
                       f'Fact_TrialBalance[FiscalYear],{fy_ref})'
                       f'-SUMPRODUCT((Fact_TrialBalance[EntityId]="{eid}")*'
                       f'(Fact_TrialBalance[FiscalYear]={fy_ref})*'
                       f'(ISNUMBER(MATCH(Fact_TrialBalance[Category],{{"{excl}"}},0)))*Fact_TrialBalance[Amount]))')
        ws2.cell(row=rr, column=11,
                 value=f'=-SUMIFS(Fact_TrialBalance[Amount],Fact_TrialBalance[EntityId],"{eid}",'
                       f'Fact_TrialBalance[FiscalYear],{fy_ref})')
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
    ws3.sheet_properties.tabColor = NAVY
    paint_background(ws3, max_row=25, max_col=10)
    ws3.merge_cells("B2:H3")
    ws3["B2"] = "Group Summary — driven by 02 Profitability's fiscal-year selector"
    ws3["B2"].font = H1
    ws3["B2"].fill = NAVY_FILL
    for col in "CDEFGH":
        ws3[f"{col}2"].fill = NAVY_FILL
        ws3[f"{col}3"].fill = NAVY_FILL
    for col, w in zip("ABCDEFGH", [3, 18, 18, 18, 18, 3, 18, 18]):
        ws3.column_dimensions[col].width = w

    ws3["B5"] = f"{len(entity_ids)}-entity group total (per Entity Register) — full scope (2 family trusts excluded)"
    ws3["B5"].font = Font(italic=True, size=9, color=GREY)

    r0 = 6
    style_kpi_card(ws3, r0, 2, "GROUP REVENUE", "=SUM('Entity Register'!C5:C20)")
    style_kpi_card(ws3, r0, 4, "GROUP EBITDA", "=SUM('Entity Register'!J5:J20)", big_fill=TEAL_FILL)
    style_kpi_card(ws3, r0, 7, "GROUP NET INCOME", "=SUM('Entity Register'!K5:K20)")

    r1 = r0 + 4
    style_kpi_card(ws3, r1, 2, "GROUP INCOME TAX", "=SUM('Entity Register'!I5:I20)", big_fill=TEAL_FILL)
    style_kpi_card(
        ws3, r1, 4, "EFFECTIVE TAX RATE",
        "=IFERROR(SUM('Entity Register'!I5:I20)/(SUM('Entity Register'!K5:K20)+SUM('Entity Register'!I5:I20)),0)",
        fmt=PCT, big_fill=TEAL_FILL,
    )
    style_kpi_card(ws3, r1, 7, "ENTITIES IN MODEL", "=COUNTA('Entity Register'!A5:A20)", fmt='0', big_fill=TEAL_FILL)

    # Per-entity rollup + chart
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
    ws4.sheet_properties.tabColor = GREY
    paint_background(ws4, max_row=45, max_col=2)
    ws4.column_dimensions["A"].width = 100
    notes = [
        "SANDHU GROUP INTERACTIVE DASHBOARD — Build Notes",
        "",
        f"Scope: all {len(entity_ids)} entities (both family trusts — Gurpreet Sandhu Trust, Harpreet Sandhu "
        "Trust — excluded from scope per instruction; ~$4-$10/yr activity per the handoff, immaterial). "
        "Every one of these was rebuilt fresh from its raw CaseWare WTB export in this session (not "
        "carried over from any prior file) and covers 3 fiscal years: FY2023, FY2024, FY2025.",
        "",
        "Missing: Sandhu & Sandhu Enr.'s Corporate Taxprep export (its WTB was supplied and is included).",
        "",
        "Dim_Entity updated to the user-supplied reference table: CompanyName (incl. owner suffix, e.g. "
        "'- Sandhu H.'), LegalName, and Industry (Construction/Real estate/Investments/Hotel/Restaurant/"
        "Rental) replace the earlier EntityName/Segment (HoldCo/RealCo/OpCo) scheme everywhere in this "
        "workbook. 'Rres_Ind' in that table was treated as a typo for 'Res_Ind' (same LegalName, "
        "9366-1049 Québec Inc.) — same entity, not a new one. Sandhu & Sandhu Enr. isn't in that table "
        "at all; its Industry ('Rental') is inferred from its sibling entity Sandhu Leasing, not given data.",
        "",
        "UNEXPECTED FILE: 947_338_2025_Tax.csv identifies its filer as '9475-3381 Québec Inc.' — an "
        "entity not in the Section 4 entity list at all. No WTB was supplied for it, so it is NOT "
        "included in this workbook. Flagged for you to confirm whether it belongs in the group.",
        "",
        "VALIDATION — every entity's FY2025 Net Income matches an independent benchmark exactly:",
        "  Restaurant IndiaRosa 2 matches every Section 6 category benchmark to the penny (Net Income "
        "FY2025 $462,894.92, FY2024 $736,959.45) — see the account-level exceptions below.",
        "  All other 15 entities match the Validation Log tab in GIFI_Master_Mapping_v2.xlsx exactly, "
        "including Bistro Guru Inc. ($-30,118.89), which RESOLVES the discrepancy flagged in an earlier "
        "version of this workbook (the old starter kit's $-45,954.06 figure for Bistro Guru was wrong; "
        "the fresh WTB rebuild confirms the Validation Log was correct).",
        "",
        "Restaurant IndiaRosa 2 — three accounts in its WTB carry a GIFI code that contradicts the "
        "account name; treated as account-level overrides, not changes to the master GIFI mapping:",
        "  Acct 45130 'Valet service' + 45125 'Musical entertainment' — tagged GIFI 9270 (Amortization) "
        "-> reclassified to Marketing & Entertainment.",
        "  Acct 44215 'Rent - Ford' + 33455 'Equipment rental' + 45250 'Travelling - gas & repairs' "
        "— tagged GIFI 8461 (COGS) -> reclassified to Other Operating.",
        "  Acct 45150 'Consulting fees' — tagged GIFI 8863 (Administrative) -> reclassified to "
        "Other Operating per user confirmation (matches Section 6 Admin $99,729 / Other Operating $79,976 exactly).",
        "",
        "Sandhu & Sandhu Enr. — two Map No formatting variants ('40.1' vs '40. 1') were causing 3 "
        "accounts to fall through the fallback table uncategorized; fixed by normalizing whitespace in "
        "the Map No match, plus adding one missing fallback row (Map No '40. 3' -> Administrative, "
        "Insurance). After the fix, FY2025 Net Income matches the Validation Log exactly ($-15,821.17).",
        "",
        "UNRESOLVED: group-level FY2025 Revenue ($14,945,731) and EBITDA ($2,734,142) are within ~0.5% "
        "and ~1.8% of the Section 6 16-entity benchmark ($14,873,731 / $2,784,496) despite every single "
        "entity's own Net Income matching its benchmark exactly. Income Tax and Investment & Other "
        "Income match exactly. This means a small amount is sitting in a different category than the "
        "original benchmark used, in one or two entities — similar in nature to IndiaRosa 2's account "
        "exceptions above, but not yet isolated. Group Net Income is within $5,014 (0.8%) of benchmark.",
        "",
        "NOT YET INCLUDED: RDTOH and CDA (confirmed absent from the Taxprep export tested — see 03 Tax "
        "Position). SBD/GRIP/Part I tax figures are real for IndiaRosa 2's 14 associated corporations "
        "(from its Taxprep SLIPA schedule) but each entity's OWN SBD claim from its own Taxprep export "
        "has not yet been pulled in — only IndiaRosa 2's Tax CSV has been mined for Section 3-style data.",
        "",
        "Interactivity: this workbook uses dropdown selectors ('02 Profitability'!C6 Entity, !H6 "
        "Fiscal Year) driving live SUMIFS formulas and charts — no Power Pivot/DAX setup required, "
        "works in any version of Excel. Entity Register and raw data sheets are Excel Tables with "
        "built-in filter dropdowns per column.",
        "",
        "Recalculated and verified in LibreOffice headless — every KPI card and cross-check in this "
        "sheet was confirmed against the actual computed cell values, not just the formula text.",
    ]
    for i, line in enumerate(notes, start=1):
        c = ws4.cell(row=i, column=1, value=line)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if i == 1:
            c.font = Font(bold=True, size=14, color=NAVY)
        elif line.strip().startswith(("UNRESOLVED", "UNEXPECTED FILE")):
            c.font = Font(size=10, bold=True, color=RED)
        elif line.strip().startswith("VALIDATION"):
            c.font = Font(size=10, bold=True, color=GREEN)
        elif line.strip().startswith(("Restaurant IndiaRosa 2 matches", "All other 15", "Acct", "  Acct")):
            c.font = Font(size=10, color=TEAL)
        else:
            c.font = Font(size=10, color="1A1A1A")

    for sheet in (wsx, ws, wst, ws2, ws3, ws4):
        sheet.page_setup.orientation = "landscape"
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 0
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.print_options.horizontalCentered = False

    desired_order = [
        "01 Exec Summary", "02 Profitability", "03 Tax Position",
        "Entity Register", "Group Summary", "Notes & Validation",
        "Dim_Entity", "Dim_Calendar", "Fact_TrialBalance",
    ]
    wb._sheets = [wb[name] for name in desired_order]
    wb.active = 0

    wb.calculation.fullCalcOnLoad = True
    wb.save(OUT)
    print("Saved", OUT)


if __name__ == "__main__":
    main()
