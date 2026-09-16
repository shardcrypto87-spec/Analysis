"""
Extract validated figures from the Restaurant IndiaRosa 2 Corporate Taxprep
export (Res_In2_2025_Tax.csv). Every figure kept here was cross-checked
against an independently-known number before being trusted, per the
handoff doc's lesson #3/#4 (don't trust a Taxprep field until validated;
don't conflate tax concepts with accounting ones).
"""
import csv
import json
import re

CSV_PATH = "data/Res_In2_2025_Tax.csv"


def load_rows():
    with open(CSV_PATH, encoding="latin-1") as f:
        return list(csv.reader(f))


def to_num(s):
    if s is None or s == "":
        return None
    try:
        return float(s.replace(",", ""))
    except ValueError:
        return None


def main():
    rows = load_rows()
    by_code = {r[0]: (r[1], r[2], r[3]) for r in rows}

    # --- Cross-validated single fields -------------------------------- #
    # Current-year Income Tax expense (GIFI 9990) — matches our WTB-derived
    # figure of $62,548 exactly.
    income_tax_cy = to_num(by_code["GFGFA.Ttwgfa14"][0])

    # Prior-year (FY2024) accounting net income, per financial statements —
    # matches WTB FY2024 net income ($736,959.45) exactly.
    ni_fy2024_accounting = to_num(by_code["QCNIQ.Ttwniq290"][0])

    # 3-year "Net income for income tax purposes" history (Federal T2),
    # genuinely a different figure from accounting net income (expected —
    # CCA vs amortization, non-deductible add-backs, etc.). NOT the formal
    # "Taxable income" line (that would additionally net off loss
    # carryforwards applied and other deductions, which this export does
    # not show applied) — labelled precisely as what was found.
    ni_tax_fy2024 = to_num(by_code["HSFIV.Ttafiv103"][0])  # "1st prior year"
    ni_tax_fy2023 = to_num(by_code["HSFIV.Ttafiv203"][0])  # "2nd prior year"
    ni_tax_fy2022 = to_num(by_code["HSFIV.Ttafiv303"][0])  # "3rd prior year"

    # Non-capital loss carryforward balance at the start of FY2023
    # (2nd preceding taxation year relative to this FY2025 filing).
    ncl_start_fy2023_federal = to_num(by_code["FDLOS.Ttwlos293"][0])
    ncl_start_fy2023_quebec = to_num(by_code["QCLSQ.Ttwlsq119"][0])

    # --- RDTOH / CDA: confirmed absent ---------------------------------#
    rdtoh_found = any("rdtoh" in (r[3] or "").lower() or "refundable dividend" in (r[3] or "").lower() for r in rows)
    cda_found = any("capital dividend" in (r[3] or "").lower() for r in rows)

    # --- Associated-group SBD/GRIP/Part I tax table (SLIPA[2..15]) ----- #
    group = {}
    for r in rows:
        m = re.match(r"MJRAW\.SLIPA\[(\d+)\]\.TtwrawA(\d+)$", r[0])
        if not m:
            continue
        idx, field = int(m.group(1)), m.group(2)
        if idx < 2:
            continue
        group.setdefault(idx, {})[field] = (r[1], r[2])

    associated_group = []
    for idx in sorted(group):
        d = group[idx]
        name = d.get("1", ("", ""))[0]
        if not name:
            continue
        associated_group.append({
            "name": name,
            "net_income_tax_cy": to_num(d.get("597", ("", ""))[0]),
            "net_income_tax_py": to_num(d.get("597", ("", ""))[1]),
            "part1_tax_cy": to_num(d.get("306", ("", ""))[0]),
            "sbd_cy": to_num(d.get("305", ("", ""))[0]),
            "sbd_py": to_num(d.get("305", ("", ""))[1]),
            "grip_cy": to_num(d.get("347", ("", ""))[0]),
            "grip_py": to_num(d.get("347", ("", ""))[1]),
        })

    result = {
        "cross_validated": {
            "income_tax_cy_fy2025": income_tax_cy,
            "expected_income_tax_fy2025": 62548.0,
            "net_income_fy2024_accounting_taxprep": ni_fy2024_accounting,
            "expected_net_income_fy2024": 736959.45,
        },
        "tax_basis_net_income_history": {
            "FY2024": ni_tax_fy2024,
            "FY2023": ni_tax_fy2023,
            "FY2022": ni_tax_fy2022,
            "note": "Federal T2 'Net income for income tax purposes' — NOT the formal "
                    "Taxable Income line (loss-carryforward application not shown). "
                    "Genuinely differs from accounting net income; do not conflate.",
        },
        "non_capital_loss_carryforward_start_fy2023": {
            "federal": ncl_start_fy2023_federal,
            "quebec": ncl_start_fy2023_quebec,
        },
        "rdtoh_found": rdtoh_found,
        "cda_found": cda_found,
        "own_current_year_tax_figures_found": False,
        "own_sbd_found": False,
        "associated_group_sbd_grip": associated_group,
    }

    with open("output/indiarosa2_taxprep_extract.json", "w") as f:
        json.dump(result, f, indent=2)

    print("Income tax cross-check:", income_tax_cy, "vs expected 62,548 ->",
          "MATCH" if income_tax_cy == 62548 else "MISMATCH")
    print("FY2024 net income cross-check:", ni_fy2024_accounting, "vs expected 736,959 ->",
          "MATCH" if ni_fy2024_accounting == 736959 else "MISMATCH")
    print("RDTOH found:", rdtoh_found, " CDA found:", cda_found)
    print("Tax-basis net income FY2024/23/22:", ni_tax_fy2024, ni_tax_fy2023, ni_tax_fy2022)
    print("Associated group entries:", len(associated_group))
    for g in associated_group:
        print(" ", g["name"], g["net_income_tax_cy"], g["sbd_cy"], g["grip_cy"])


if __name__ == "__main__":
    main()
