"""
Parse the Restaurant IndiaRosa 2 WTB export into corrected Fact_TrialBalance rows
(FY2025 + FY2024), applying the GIFI master mapping plus the account-level
exceptions confirmed against the Section 6 benchmarks.
"""
import json
import openpyxl
from collections import defaultdict

GIFI_XLSX = "data/GIFI_Master_Mapping_v2.xlsx"
WTB_XLSX = "data/Res_In2_2025_WTB.xlsx"
FILE_PREFIX = "Res_In2"
ENTITY_ID = "9455-8236 QI"

# Account-level overrides: the source WTB tags these accounts with a GIFI code
# that contradicts the account name (e.g. "Valet service" under GIFI 9270,
# "Amortization of financing costs"). Confirmed against Section 6 benchmarks
# and the user's explicit call on the Consulting fees split.
EXCEPTIONS = {
    45130: "Marketing & Entertainment",  # Valet service (GIFI 9270 = Amortization, wrong)
    45125: "Marketing & Entertainment",  # Musical entertainment (GIFI 9270, wrong)
    44215: "Other Operating",            # Rent - Ford (GIFI 8461 = COGS, wrong)
    33455: "Other Operating",            # Equipment rental (GIFI 8461, wrong)
    45250: "Other Operating",            # Travelling - gas & repairs (GIFI 8461, wrong)
    45150: "Other Operating",            # Consulting fees (GIFI 8863 = Administrative; user confirmed Other Operating)
}


def load_gifi_map():
    wb = openpyxl.load_workbook(GIFI_XLSX, data_only=True)
    gifi_map = {}
    for code, cat, _ in wb["GIFI Master Mapping"].iter_rows(min_row=2, values_only=True):
        if code is not None:
            gifi_map[str(code).strip()] = cat

    fallback = {}
    started = False
    for row in wb["Fallback - Map No (blank GIFI)"].iter_rows(values_only=True):
        if row[0] == "FilePrefix":
            started = True
            continue
        if started and row[0]:
            fp, mapno, cat, _ = row
            fallback[(fp, str(mapno).strip())] = cat
    return gifi_map, fallback


def categorize(acct, gifi, mapno, gifi_map, fallback):
    if acct in EXCEPTIONS:
        return EXCEPTIONS[acct]
    gifi_str = str(gifi).strip() if gifi is not None else ""
    if gifi_str:
        cat = gifi_map.get(gifi_str)
        if cat:
            return cat
    mapno_str = str(mapno).strip() if mapno is not None else ""
    return fallback.get((FILE_PREFIX, mapno_str), "UNCATEGORIZED")


def main():
    gifi_map, fallback = load_gifi_map()
    wb = openpyxl.load_workbook(WTB_XLSX, data_only=True)
    ws = wb["Sheet1"]
    rows = list(ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True))
    inc = [r for r in rows if r[4] == "Income statement"]

    totals = {"FY2025": defaultdict(float), "FY2024": defaultdict(float)}
    uncategorized = []
    for r in inc:
        acct, name, mapno, gifi = r[0], r[1], r[3], r[13]
        final_cy, final_py = r[12], r[14]
        cat = categorize(acct, gifi, mapno, gifi_map, fallback)
        if cat == "UNCATEGORIZED" and (final_cy or final_py):
            uncategorized.append((acct, name, final_cy, final_py))
        totals["FY2025"][cat] += final_cy or 0
        totals["FY2024"][cat] += final_py or 0

    out_rows = []
    for fy in ["FY2025", "FY2024"]:
        for cat, amt in totals[fy].items():
            if cat == "UNCATEGORIZED":
                continue
            if abs(amt) < 0.005:
                continue
            out_rows.append([ENTITY_ID, cat, fy, round(amt, 2)])

    net_income = {fy: round(-sum(totals[fy].values()), 2) for fy in totals}

    result = {
        "entity_id": ENTITY_ID,
        "rows": out_rows,
        "net_income": net_income,
        "uncategorized_with_balance": uncategorized,
    }
    with open("output/indiarosa2_fact_rows.json", "w") as f:
        json.dump(result, f, indent=2)

    print("Net income FY2025:", net_income["FY2025"], " (benchmark 462,894.92)")
    print("Net income FY2024:", net_income["FY2024"], " (benchmark 736,959.45)")
    print("Uncategorized rows with nonzero balance:", uncategorized)
    print()
    for fy in ["FY2025", "FY2024"]:
        print(f"--- {fy} ---")
        for cat, amt in sorted(totals[fy].items()):
            if cat != "UNCATEGORIZED" or amt:
                print(f"  {cat:30s} {amt:15,.2f}")


if __name__ == "__main__":
    main()
