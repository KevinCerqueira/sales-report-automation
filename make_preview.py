"""Builds preview.png (portfolio image) from report/sales_report.xlsx."""

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

x = pd.read_excel("report/sales_report.xlsx", sheet_name=None)
data, issues = x["Clean data"], x["Issues"]
data["month"] = pd.to_datetime(data["date"]).dt.strftime("%b")
order = pd.to_datetime(data["date"]).dt.month
pv = data.assign(m=order).pivot_table(index=["m", "month"], columns="branch", values="total", aggfunc="sum")
pv.index = pv.index.get_level_values(1)

fig = plt.figure(figsize=(12, 5.2))
ax0 = fig.add_axes([0.02, 0.1, 0.28, 0.75])
ax0.axis("off")
ax0.set_title("Input: 3 messy branch files", loc="left", fontweight="bold")
lines = [
    ("vendas_sao_paulo.csv", "';' separator · R$ 1.234,56 · dd/mm/yyyy · PT headers"),
    ("sales_rio.xlsx", "Excel · English headers · ISO dates"),
    ("curitiba_export.csv", "blank lines · duplicates · typos · bad values"),
]
y = 0.9
for name, desc in lines:
    ax0.text(0, y, name, fontsize=11, family="monospace", fontweight="bold")
    ax0.text(0, y - 0.08, desc, fontsize=9, color="#555")
    y -= 0.25
dups = int((issues["issue"] == "duplicate row (removed)").sum())
ax0.text(0, 0.05, f"→ {len(data):,} clean sales\n→ {dups} duplicates removed\n→ {len(issues) - dups} bad rows flagged "
         "(not dropped silently)", fontsize=10, color="#1F4E78", fontweight="bold")

ax = fig.add_axes([0.38, 0.12, 0.6, 0.73])
bottom = None
for col, color in zip(pv.columns, ["#1F4E78", "#2E86AB", "#A3C4DC"]):
    ax.bar(pv.index, pv[col] / 1000, bottom=bottom, label=col, color=color)
    bottom = pv[col] / 1000 if bottom is None else bottom + pv[col] / 1000
ax.set_title("Output: monthly revenue by branch (R$ thousand)", fontweight="bold")
ax.legend(frameon=False)
for s in ("top", "right"):
    ax.spines[s].set_visible(False)
fig.suptitle("Sales spreadsheet consolidation — one command, one Excel report", fontsize=13)
fig.savefig("preview.png", dpi=130)
print("saved preview.png")
