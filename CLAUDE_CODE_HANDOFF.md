# Sandhu Group — Project Handoff for Claude Code

**Goal:** Build an interactive Excel workbook (Power Pivot data model + PivotTables + Slicers) from CaseWare Working Trial Balance (WTB) exports and Corporate Taxprep exports, for a multi-entity restaurant/real estate/holding group. This document has everything validated so far — use it as the source of truth rather than re-deriving anything from scratch.

---

## 1. Project Context

**Restaurant Sandhu Group** — 18 entities: restaurants (OpCos), a hospitality property, real estate entities, holding companies, a partnership, and two family trusts. Fiscal year end March 31.

**The end goal:** a self-service interactive Excel workbook the practitioner can rebuild every year without outside help — real financial data (from CaseWare), real tax attributes (from Corporate Taxprep), joined on Entity + Fiscal Year, explorable via PivotTables and Slicers.

**Why Excel, not Power BI:** avoids per-viewer licensing costs (Power BI Pro is ~$14/user/month) when sharing with people outside the firm. Same underlying engine (Power Pivot = Power BI's data model + DAX), just a different container.

---

## 2. Data Source #1: CaseWare Working Trial Balance (WTB) exports

**This is the preferred CaseWare export format going forward** — richer and easier to parse than the older "Trial Balance" leadsheet export (which required fragile row-buffering logic; avoid that approach).

### Column structure (confirmed from real files)
| Col | Field | Notes |
|---|---|---|
| A | Account No | Numeric |
| B | Name | Account description |
| C | Lock | Ignore |
| D | Map No | CaseWare's internal group code (e.g. "40. 1") — **not safe as a universal key across entities**, only use as a last-resort fallback, scoped per-entity |
| E | Type | "Balance sheet" or "Income statement" — use this to filter P&L rows, not digit/letter guessing |
| F | Sign | Debit/Credit |
| G | L/S | Ignore |
| H | Tax Export Code | Ignore |
| I | Class | CaseWare's own plain-English classification (e.g. "Assets - Current - other quick") — potentially useful, not yet used |
| J | Opening Balance | |
| K | Transactions | |
| L | Adjustments | |
| M | Final: [current year end date] | Current year balance — **this is the number to use** |
| N | GIFI | CRA-standard code — **this is the primary categorization key**, universal across all Canadian corporations |
| O | Prior: [prior year end date] | Prior year comparative — gives a second fiscal year for free, no separate export needed |
| P–S | Prior years 2/3/4, Budget | Available but not yet used |
| T | Annotation | Ignore |

**Header row is row 1** — no company-name preamble row (unlike the older leadsheet export). Company/entity identity is **not inside the file** — it must come from the filename, matched against a maintained lookup table (see Section 4).

### Critical known issue: GIFI type mismatch
In the raw file, GIFI is stored as a **number** (e.g. `9060`), not text. Any mapping table keyed on GIFI as text will silently fail to match unless the WTB's GIFI column is explicitly cast to text first. This caused a real, hard-to-diagnose bug earlier — cast early, cast always.

### Critical known issue: Map No is not a safe universal key
The same Map No (e.g. "40. 1") means different things in different entities' books (Payroll in one file, Interest & Bank Charges in another). Any fallback logic using Map No **must be scoped to the specific entity** (composite key: `FilePrefix + "|" + MapNo`), never used as a bare universal lookup.

---

## 3. Data Source #2: Corporate Taxprep exports

**Export format:** Corporate Taxprep's Xpress Filter export → CSV, encoding is `latin-1` (not UTF-8) — 4 columns: `FieldCode, ThisYearValue, LastYearValue, Description`.

**Xpress Filter settings to use** (found via testing, not yet perfect — refine if issues arise):
- ✅ Check: "Data Entered This Year," "Data Entered Last Year," "Rolled Forward Data," "Imported Data"
- ❌ Leave unchecked on the exclude side: everything, especially "Calculated cells" (most useful figures — RDTOH, GRIP, tax payable — are calculated, not manually entered, and would be lost if excluded)

**Important limitation:** Xpress Filter filters by *how a cell got its data* (entered/rolled-forward/calculated), **not by which schedule or line item** it belongs to. There is no way to export "just Schedule 1" or "just RDTOH." The export is broad; filtering down to specific fields happens after export, by field-code pattern matching.

### What's been found and validated so far (from one entity's export — Restaurant IndiaRosa 2)
- **SBD allocation and business limit for the whole associated group** — surprisingly, one entity's Taxprep file contains an array (`SLIPA[1]` through `SLIPA[15]`) listing every associated corporation by name and business number, each with its own SBD dollar amount and business limit %. **This means you likely don't need all 18 entities' Taxprep exports to get group-wide SBD data — one is enough**, assuming the group's association schedule is complete in every file (verify this assumption once more exports are available).
- **GRIP closing balance** — found and appears reliable.
- **Part I tax payable per associated corp** — found within the SLIPA array.
- **Prior-year net income** — found and cross-validated exactly against CaseWare's own prior-year column (see Section 6).

### What was NOT reliably found
- **RDTOH balance** (opening/closing continuity) — not found in the one export tested. Likely requires the "Rolled Forward Data" checkbox (added after the first export attempt) — untested whether this fixes it.
- **CDA balance** — same, not found, same likely fix.
- **This entity's own current-year net income / Part I tax payable** (as opposed to the associated group's figures) — several candidate fields were ambiguous (some explicitly labeled "prior year" in ways that didn't reconcile cleanly). Don't assume a field name without validating the actual number against a known-correct figure.

**Recommendation for Claude Code:** treat every Taxprep field discovery as unverified until cross-checked against a real, independently-known number (e.g., net income already validated from CaseWare). Don't build categorization logic on an assumed field meaning.

---

## 4. Entity List & Segmentation (validated)

| EntityId | EntityName | Segment | Notes |
|---|---|---|---|
| 9357-7427 QI | Harpreet Sandhu (9357-7427) | HoldCo | |
| 9366-1676 QI | Gurpreet Sandhu (9366-1676) | HoldCo | Largest loss entity in group |
| 9425-4547 QI | Harpreet Sandhu (9425-4547) | HoldCo | |
| 9455-4169 QI | Gurpreet Sandhu (9455-4169) | HoldCo | |
| 9455-4177 QI | Harpreet Sandhu (9455-4177) | HoldCo | |
| 9455-4201 QI | Sandhu Holdco | HoldCo | Receives ~$679K/yr intercompany dividends |
| 9108-6876 QI | Auberge St Louis | **Hospitality OpCo** | Reclassified — has room/spa/subcontracted-labor revenue, not a passive RealCo |
| N/A | Bistro Guru Inc. | Restaurant OpCo | FY2024: $0 sales revenue, income came entirely from a $25K intercompany management fee received — confirm with client whether this location operated that year |
| 9096-0436 QI | Lobby Lounge | Restaurant OpCo | |
| 9504-7510 QI | Restaurant Heritage | Restaurant OpCo | Second-largest loss entity |
| 9455-8236 QI | Restaurant IndiaRosa 2 | Restaurant OpCo | Template entity — first validated, most complete GIFI coverage |
| 9525-2854 QI | Restaurant IndiaRosa 3 | Restaurant OpCo | Small/new location, **no rent expense at all** — confirm with client |
| 9366-1049 QI | Restaurant IndiaRosa 1 | Restaurant OpCo | |
| 9098-0558 QI | Restaurant Sandhu | Restaurant OpCo | |
| 2749-8567 QI | Sandhu Leasing | RealCo | |
| N/A | Sandhu & Sandhu Enr. | **RealCo** | Reclassified — legally a Partnership, but business is rental real estate (has rental income, property tax, amortization of "income-producing properties") |
| N/A | Gurpreet Sandhu Trust | Trust | Minimal activity (~$4/year) |
| N/A | Harpreet Sandhu Trust | Trust | Minimal activity (~$10/year) |

---

## 5. GIFI Master Mapping (validated, 57 codes)

Full file: `GIFI_Master_Mapping_v2.xlsx` (already built, attach/reference if available). Two tabs:
- **GIFI Master Mapping** — GIFI Code → Category, 57 rows
- **Fallback - Map No (blank GIFI)** — `FilePrefix, MapNo, Category, Notes` — 6 rows, for the handful of accounts with no GIFI code. **Must be joined on the composite key (FilePrefix + MapNo), never MapNo alone.**

### The 13 standard categories
`Operating Revenue`, `Investment & Other Income`, `COGS`, `Payroll & Benefits`, `Occupancy`, `Marketing & Entertainment`, `Administrative`, `Management Fees (Interco)`, `Interest & Financing`, `Repairs & Maintenance`, `Amortization`, `Other Operating`, `Income Tax`

### Key GIFI codes (most common; see full file for all 57)
```
8000 → Operating Revenue        9060 → Payroll & Benefits
8320 → COGS                     8912 → Occupancy
8710 → Interest & Financing     8670 → Amortization
8860 → Administrative           8871 → Management Fees (Interco)
9990 → Income Tax               8096 → Investment & Other Income (dividends)
8094 → Investment & Other Income (interest income)
```

---

## 6. Validation benchmarks (use these to confirm any new build is working correctly)

Cross-validated across 3 independent extraction methods (original tax workbook, leadsheet CaseWare parsing, WTB parsing) — if a new build doesn't match these, something is wrong in the new build, not in these numbers.

| Entity | FY2025 Net Income | FY2024 Net Income |
|---|---|---|
| Restaurant IndiaRosa 2 | $462,894.92 | $736,959.45 |
| Restaurant IndiaRosa 1 | $253,352.18 | $410,258.04 |
| Restaurant IndiaRosa 3 | -$5,721.66 | $0.00 |
| Gurpreet Sandhu Corp (9366-1676) | -$660,352.13 | -$294,358.42 |
| Sandhu Holdco | $571,535.23 | $1,487,098.69 |

**Group totals, FY2025, 16 entities (excludes 2 trusts):**
Revenue $14,873,731 · Investment & Other Income $1,552,868 · EBITDA $2,784,496 · Net Income $619,496 · Income Tax $148,212 · Effective Tax Rate 19.3%

**IndiaRosa 2 FY2025 category breakdown** (the fullest single-entity validation done):
Revenue $5,213,629 · COGS $1,647,559 · Payroll $1,780,201 · Occupancy $474,172 · Marketing $141,793 · Administrative $99,729 · Interest & Financing $146,738 · Repairs $88,959 · Amortization $199,058 · Mgmt Fees $30,000 · Other Operating $79,976 · Income Tax $62,548 · Net Income $462,895

---

## 7. Data Model Spec (star schema)

```
Dim_Entity (18 rows: EntityId, EntityName, Segment)
Dim_Calendar (FY2024, FY2025: FiscalYear, YearEnd, YearLabel)
        │
        └── Fact_TrialBalance (EntityId, Category, FiscalYear, Amount)
                — from CaseWare WTB exports
        └── Fact_TaxAttributes (EntityId, FiscalYear, SBD, GRIP, BusinessLimitPct, PartITaxPayable, ...)
                — from Corporate Taxprep exports, once field mapping is more fully validated
```

Relationships: both fact tables relate to `Dim_Entity[EntityId]` and `Dim_Calendar[FiscalYear]`, many-to-one, single direction.

---

## 8. DAX Measures (validated, ready to reuse)

```dax
Revenue = CALCULATE(SUM(Fact_TrialBalance[Amount]), Fact_TrialBalance[Category]="Operating Revenue") * -1
Investment Income = CALCULATE(SUM(Fact_TrialBalance[Amount]), Fact_TrialBalance[Category]="Investment & Other Income") * -1
EBITDA = CALCULATE(SUM(Fact_TrialBalance[Amount]), NOT Fact_TrialBalance[Category] IN {"Interest & Financing", "Amortization", "Income Tax"}) * -1
EBITDA Margin = DIVIDE([EBITDA], [Revenue])
Net Income = CALCULATE(SUM(Fact_TrialBalance[Amount])) * -1
Income Tax = CALCULATE(SUM(Fact_TrialBalance[Amount]), Fact_TrialBalance[Category]="Income Tax")
Effective Tax Rate = DIVIDE([Income Tax], [Net Income] + [Income Tax])
% of Group Revenue = DIVIDE([Revenue], CALCULATE([Revenue], ALL(Dim_Entity)))
-- Plus one per category (COGS, Payroll, Occupancy, Marketing, Administrative, Repairs,
-- Amortization, Interest Financing, Management Fees, Other Operating) using the same
-- CALCULATE(...Category="X"...) * -1 pattern as Revenue above.
```

---

## 9. The Interactive Excel Build — what Claude Code needs to actually do

1. **Parse WTB files** (Section 2 spec) into a clean `Fact_TrialBalance` table — filter `Type = "Income statement"`, cast GIFI to text, join against the GIFI master mapping, fall back to the entity-scoped Map No table for blanks.
2. **Parse Taxprep CSVs** (Section 3) into `Fact_TaxAttributes` — start with SBD and GRIP (validated), treat RDTOH/CDA as exploratory until cross-validated against a known figure.
3. **Load everything into Excel's Data Model** (Power Pivot) — via `xlwings`, `openpyxl` + COM automation, or direct Excel automation, since this is the part that couldn't be done outside a real Excel instance.
4. **Build relationships** per Section 7.
5. **Write the DAX measures** from Section 8 into the model.
6. **Build PivotTables and Slicers** — Entity register (PivotTable, EntityName rows, category/measure columns, Segment slicer), Company Profile (a second PivotTable or set of cells driven by an EntityName slicer, using GETPIVOTDATA or CUBEVALUE formulas for a "click one entity, see its detail" experience), FiscalYear slicer synced across sheets.
7. **Validate against Section 6** before considering it done — every number should match those benchmarks exactly.
8. **Test that it actually opens and updates interactively** — this is the step I could never do; Claude Code running locally can and should verify this before calling it finished.

---

## 10. Files already built this project (attach alongside this brief if available)

- `GIFI_Master_Mapping_v2.xlsx` — the master category dictionary
- `Sandhu_Excel_Model_Data.xlsx` — plain data + DAX reference (no working Power Pivot model yet — this is what needs finishing)
- `sandhu_powerquery_simple_loader.txt` / `sandhu_final_loader_GIFI.txt` — Power Query M scripts (Power BI version, adaptable logic reference for the Excel Power Query step)

---

## 11. Known lessons from this project — don't repeat these mistakes

1. **Never assume a code (account number, Map No) is safe as a universal key across entities without checking for collisions first.** GIFI is the one genuinely safe universal key found so far.
2. **Always cast join keys to matching types explicitly** (the GIFI number-vs-text bug cost real debugging time).
3. **Validate every new number against an independently-known figure before trusting it** — this project caught two real, silent bugs this way (mislabeled categories, an entity's revenue accidentally zeroed) that would have gone to a client otherwise.
4. **Don't build tax-concept fields (taxable income, SBD, RDTOH) from accounting data as if they were interchangeable** — net income and taxable income are genuinely different figures; where real tax data isn't available, either omit the field or clearly label it as an accounting-based proxy.
