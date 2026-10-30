"""What the tidy table is good for: two questions that take a minute once the data is long,
and an afternoon of copy-paste when it isn't.

    python mcreader.py data && python examples/quick_look.py
"""

from pathlib import Path

import pandas as pd

latest = sorted(Path("output").iterdir())[-1]
t = pd.read_csv(latest / "tidy.csv", dtype={"naics": str})

# 1. Which kinds of store grew fastest from 2019 to 2022? (subsectors only, so totals don't crowd the list)
sales = t[t["file"].str.startswith("Annual retail sales") & (t["level"] == 1) & t["measure"].isin(["2019", "2022"])]
wide = sales.pivot_table(index="kind_of_business", columns="measure", values="value")
wide["growth_%"] = (wide["2022"] / wide["2019"] - 1) * 100
print("Sales growth, 2019 to 2022 ($ millions)")
print(wide.sort_values("growth_%", ascending=False).round(1).to_string(), "\n")

# 2. How much of each subsector's operating cost is payroll?
exp = t[t["file"].str.startswith("Detailed operating expenses") & (t["level"] == 1)
        & (t["measure"] == "Annual payroll | Percent of total")]
print("Payroll as a share of operating expenses (%)")
print(exp.set_index("kind_of_business")["value"].sort_values(ascending=False).to_string())
print("(blank: the Census suppressed it; the reason is in qa/suppressed_cells.csv)")
