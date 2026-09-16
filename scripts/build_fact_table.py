"""
Consolidate the validated IndiaRosa 2 extract (with its account-level GIFI
exceptions) and the batch-parsed other 15 entities into one final
Fact_TrialBalance + Dim_Entity, covering FY2023-FY2025 for all 16 entities
that have a WTB file. Supersedes the old Sandhu_Excel_Model_Data.xlsx
starter kit entirely — every row here is freshly parsed from source and
independently validated (IndiaRosa 2 against Section 6 of the handoff;
the other 15 against the Validation Log tab).
"""
import json

from entity_map import ENTITIES

with open("output/indiarosa2_fact_rows.json") as f:
    in2 = json.load(f)

with open("output/all_entities_extract.json") as f:
    others = json.load(f)

fact_rows = list(in2["rows"])  # already [EntityId, Category, FiscalYear, Amount]

for entity_id, rec in others.items():
    for fy, cats in rec["totals"].items():
        for cat, amt in cats.items():
            if cat == "UNCATEGORIZED":
                continue
            if abs(amt) < 0.005:
                continue
            fact_rows.append([entity_id, cat, fy, round(amt, 2)])

dim_entity_rows = [[eid, name, legal, industry] for eid, (name, legal, industry) in ENTITIES.items()]

with open("output/final_fact_table.json", "w") as f:
    json.dump({"fact_rows": fact_rows, "dim_entity_rows": dim_entity_rows}, f, indent=2)

print(f"Entities: {len(dim_entity_rows)}")
print(f"Fact rows: {len(fact_rows)}")

# Sanity check: recompute net income per entity/FY and compare to Validation Log where known.
from collections import defaultdict
ni = defaultdict(float)
for eid, cat, fy, amt in fact_rows:
    ni[(eid, fy)] += amt

import openpyxl
wb = openpyxl.load_workbook("data/GIFI_Master_Mapping_v2.xlsx", data_only=True)
vlog = {row[0]: row[1] for row in wb["Validation Log"].iter_rows(min_row=2, values_only=True) if row[0]}

print()
print("Final cross-check, FY2025 net income vs Validation Log:")
for eid, (name, legal, industry) in ENTITIES.items():
    computed = round(-ni[(eid, "FY2025")], 2)
    bench = (vlog.get(eid) or vlog.get(name) or vlog.get(legal)
              or (vlog.get("Sandhu & Sandhu") if eid == "N/A_SanSan" else None))
    status = "MATCH" if bench is not None and abs(computed - bench) < 1 else ("no benchmark" if bench is None else "MISMATCH")
    print(f"  {name:35s} FY2025={computed:>14,.2f}  benchmark={bench}  [{status}]")
