"""Creates 3 made-up Excel files in sample_data/ for trying the app.
These are EXAMPLE figures only, not real company data."""
from datetime import date
from pathlib import Path

import pandas as pd

out = Path(__file__).with_name("sample_data")
out.mkdir(exist_ok=True)

suppliers = [
    ("SUP-1001", "Medichem Labs", "Ahmedabad"),
    ("SUP-1002", "BioPack Industries", "Hyderabad"),
    ("SUP-1003", "Sterile Glass Co.", "Vadodara"),
    ("SUP-1004", "API Synth Pvt Ltd", "Visakhapatnam"),
    ("SUP-1005", "ColdChain Logistics", "Pune"),
    ("SUP-1006", "Excipients India", "Mumbai"),
    ("SUP-1007", "LabEquip Traders", "Chennai"),
    ("SUP-1008", "Capsule Shells Ltd", "Bengaluru"),
]

# File 1: supplier master
pd.DataFrame(
    [{"Supplier Code": c, "Supplier Name": n, "City": city, "GST Number": f"24ABCDE{1000 + i}F1Z5"}
     for i, (c, n, city) in enumerate(suppliers)]
).to_excel(out / "1_supplier_master.xlsx", index=False)

# File 2: 3-year contracts (different heading style on purpose)
contracts = [
    ("SUP-1001", date(2023, 4, 1), date(2026, 3, 31), "Extended", date(2029, 3, 31), 1250000.00),
    ("SUP-1002", date(2023, 7, 1), date(2026, 6, 30), "Closed early", None, 480000.50),
    ("SUP-1003", date(2024, 1, 1), date(2026, 12, 31), "Active", None, 915250.75),
    ("SUP-1004", date(2022, 10, 1), date(2025, 9, 30), "Completed", None, 2300000.00),
    ("SUP-1005", date(2024, 4, 1), date(2027, 3, 31), "Active", None, 675000.00),
    ("SUP-1006", date(2023, 11, 15), date(2026, 11, 14), "Extended", date(2029, 11, 14), 340500.25),
    ("SUP-1008", date(2023, 2, 1), date(2026, 1, 31), "Closed early", None, 150000.00),
]
pd.DataFrame(
    [{"supplier code": c, "Contract Start": s, "Contract End": e, "Contract Status": st,
      "Extended Until": x, "Contract Value (Rs)": v} for c, s, e, st, x, v in contracts]
).to_excel(out / "2_contracts.xlsx", index=False)

# File 3: payments made in 2026
payments = [
    ("SUP-1001", 312500.00, 937500.00),
    ("SUP-1002", 120000.25, 360000.25),
    ("SUP-1003", 305083.58, 610167.17),
    ("SUP-1005", 225000.00, 450000.00),
    ("SUP-1006", 113500.08, 227000.17),
    ("SUP-1008", 150000.00, 0.00),
]
pd.DataFrame(
    [{"SUPPLIER_CODE": c, "Paid 2026": p, "Balance Due": b} for c, p, b in payments]
).to_excel(out / "3_payments_2026.xlsx", index=False)

print("Sample files written to", out)
