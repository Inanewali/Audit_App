"""Generate the synthetic demo population in sample_data/ (fictional names, random balances)."""

from pathlib import Path

import numpy as np
import pandas as pd

rng = np.random.default_rng(2026)
N = 400
words_a = ["Northwind", "Bluefield", "Cedar", "Harbor", "Summit", "Riverbend", "Maple",
           "Ironwood", "Silverline", "Prairie", "Granite", "Evergreen", "Lakeside", "Falcon"]
words_b = ["Traders", "Textiles", "Foods", "Logistics", "Steel", "Printing", "Pharma",
           "Agro", "Ceramics", "Holdings", "Motors", "Packaging"]
names = [f"{rng.choice(words_a)} {rng.choice(words_b)} Ltd" for _ in range(N)]
ids = 3_000_000_000 + rng.choice(1_000_000, size=N, replace=False) * 997
balances = np.round(rng.lognormal(mean=14, sigma=1.6, size=N), 2)
balances[rng.choice(N, 3, replace=False)] *= -0.01  # a few credit balances

loans = pd.DataFrame({"ACCOUNT ID": ids, "CUSTOMER NAME": names,
                      "TOTAL OUTSTANDING": balances,
                      "BRANCH": rng.choice(["Main", "North", "South", "East"], N)})
total = pd.DataFrame({"ACCOUNT ID": [None], "CUSTOMER NAME": ["Grand Total"],
                      "TOTAL OUTSTANDING": [balances.sum()], "BRANCH": [None]})
loans = pd.concat([loans, total], ignore_index=True)

staff_idx = rng.choice(N, 6, replace=False)
staff = pd.DataFrame({"Account No": [str(i) for i in ids[staff_idx]],  # text IDs on purpose
                      "Employee": [f"Employee {i + 1}" for i in range(6)]})

out = Path(__file__).resolve().parent.parent / "sample_data" / "loan_population.xlsx"
out.parent.mkdir(exist_ok=True)
with pd.ExcelWriter(out) as xw:
    loans.to_excel(xw, sheet_name="Loans", index=False)
    staff.to_excel(xw, sheet_name="Staff", index=False)
print(f"wrote {out}")
