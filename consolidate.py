"""Consolidates messy sales spreadsheets from several branches into one Excel report.

Drop any number of .csv / .xlsx files in the input folder (each branch can use its own
column names, separators, number and date formats) and run:

    python consolidate.py --input input --output report/sales_report.xlsx

Output workbook:
    Summary          KPIs + revenue by month and branch (with chart)
    Top products     revenue and units per product
    Clean data       standardized, deduplicated rows
    Issues           rows that could not be used and why (nothing is silently dropped)
"""

from __future__ import annotations

import argparse
import logging
import re
import unicodedata
from pathlib import Path

import pandas as pd

log = logging.getLogger("consolidate")

# Any header found in the files is normalized and mapped to the standard schema.
ALIASES = {
    "date": {"data", "date", "sale date", "dt venda", "data venda"},
    "product": {"produto", "item", "product", "descricao produto", "descricao"},
    "quantity": {"qtd", "quantidade", "quantity", "qty"},
    "unit_price": {"preco unitario", "unit price", "valor unit", "valor unitario", "price"},
}
BRANCH_HINTS = {"sao_paulo": "São Paulo", "rio": "Rio de Janeiro", "curitiba": "Curitiba"}


def norm(text: str) -> str:
    t = unicodedata.normalize("NFKD", str(text)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", " ", t).strip()


def map_columns(cols) -> dict:
    out = {}
    for c in cols:
        for std, names in ALIASES.items():
            if norm(c) in names:
                out[c] = std
    missing = set(ALIASES) - set(out.values())
    if missing:
        raise ValueError(f"missing columns {sorted(missing)} in {list(cols)}")
    return out


def parse_number(v) -> float | None:
    """Accepts 1234.5, '1.234,56', 'R$ 1.234,56', '1,234.56'. Returns None if invalid."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d,.\-]", "", str(v))
    if not s or not re.search(r"\d", s):
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def parse_date(v):
    if isinstance(v, pd.Timestamp) or hasattr(v, "year"):
        return pd.Timestamp(v).normalize()
    s = str(v).strip()
    fmt = "%d/%m/%Y" if re.match(r"\d{2}/\d{2}/\d{4}$", s) else None
    d = pd.to_datetime(s, format=fmt, errors="coerce")
    return None if pd.isna(d) else d.normalize()


def branch_name(path: Path) -> str:
    stem = norm(path.stem).replace(" ", "_")
    return next((v for k, v in BRANCH_HINTS.items() if k in stem), path.stem)


def read_file(path: Path) -> pd.DataFrame:
    if path.suffix.lower() in (".xlsx", ".xls"):
        df = pd.read_excel(path, dtype=object)
    else:
        df = pd.read_csv(path, sep=None, engine="python", dtype=str, skip_blank_lines=True)
    df = df.dropna(how="all")
    df = df.rename(columns=map_columns(df.columns))[list(ALIASES)]
    df["branch"], df["source_file"] = branch_name(path), path.name
    df["source_row"] = df.index + 2
    return df


def clean(raw: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    df = raw.copy()
    df["product"] = df["product"].astype(str).str.strip().str.title().str.replace(r"\s+", " ", regex=True)
    df["date"] = df["date"].map(parse_date)
    df["quantity"] = df["quantity"].map(parse_number)
    df["unit_price"] = df["unit_price"].map(parse_number)

    problems = pd.Series("", index=df.index)
    for col, msg in (("date", "invalid date"), ("quantity", "missing/invalid quantity"),
                     ("unit_price", "invalid unit price")):
        problems[df[col].isna()] += msg + "; "
    issues = raw.loc[problems != ""].assign(issue=problems[problems != ""].str.rstrip("; "))

    ok = df[problems == ""].copy()
    key = ["branch", "date", "product", "quantity", "unit_price"]
    dup = ok.duplicated(subset=key, keep="first")
    issues = pd.concat([issues, raw.loc[ok.index[dup]].assign(issue="duplicate row (removed)")])
    ok = ok[~dup]
    ok["quantity"] = ok["quantity"].astype(int)
    ok["total"] = (ok["quantity"] * ok["unit_price"]).round(2)
    ok["month"] = ok["date"].dt.to_period("M").astype(str)
    return ok.sort_values(["date", "branch"]).reset_index(drop=True), issues


def write_report(data: pd.DataFrame, issues: pd.DataFrame, files: list[Path], out: Path) -> None:
    from openpyxl.chart import BarChart, Reference
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    out.parent.mkdir(parents=True, exist_ok=True)
    by_month = data.pivot_table(index="month", columns="branch", values="total", aggfunc="sum").fillna(0).round(2)
    by_month["Total"] = by_month.sum(axis=1)
    top = (data.groupby("product").agg(units=("quantity", "sum"), revenue=("total", "sum"))
           .sort_values("revenue", ascending=False).round(2))
    kpis = pd.DataFrame({"Metric": ["Files processed", "Valid sales", "Duplicates removed", "Rows with issues",
                                    "Total revenue (R$)", "Average ticket (R$)", "Period"],
                         "Value": [len(files), len(data), int((issues["issue"] == "duplicate row (removed)").sum()),
                                   int((issues["issue"] != "duplicate row (removed)").sum()),
                                   round(data["total"].sum(), 2), round(data["total"].mean(), 2),
                                   f"{data['date'].min():%d/%m/%Y} – {data['date'].max():%d/%m/%Y}"]})

    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        kpis.to_excel(xw, sheet_name="Summary", index=False)
        by_month.to_excel(xw, sheet_name="Summary", startrow=len(kpis) + 3)
        top.to_excel(xw, sheet_name="Top products")
        data.drop(columns=["month"]).assign(date=data["date"].dt.date).to_excel(
            xw, sheet_name="Clean data", index=False)
        issues.to_excel(xw, sheet_name="Issues", index=False)

        head_font, head_fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E78")
        for ws in xw.book.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if cell.row == 1 or (ws.title == "Summary" and cell.row == len(kpis) + 4):
                        if cell.value is not None:
                            cell.font, cell.fill = head_font, head_fill
                    if isinstance(cell.value, float):
                        cell.number_format = '"R$" #,##0.00'
            for i, col in enumerate(ws.columns, 1):
                width = max((len(str(c.value)) for c in col if c.value is not None), default=8)
                ws.column_dimensions[get_column_letter(i)].width = min(width + 3, 45)
            if ws.title != "Summary":
                ws.freeze_panes, ws.auto_filter.ref = "A2", ws.dimensions

        ws = xw.book["Summary"]
        r0, n_branches = len(kpis) + 4, len(by_month.columns) - 1
        chart = BarChart()
        chart.type, chart.grouping, chart.overlap = "col", "stacked", 100
        chart.title, chart.y_axis.title = "Monthly revenue by branch (R$)", "R$"
        chart.add_data(Reference(ws, min_col=2, max_col=1 + n_branches, min_row=r0, max_row=r0 + len(by_month)),
                       titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=1, min_row=r0 + 1, max_row=r0 + len(by_month)))
        chart.width, chart.height = 20, 10
        ws.add_chart(chart, "H2")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--input", default="input")
    ap.add_argument("--output", default="report/sales_report.xlsx")
    a = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")

    files = sorted(p for p in Path(a.input).iterdir() if p.suffix.lower() in (".csv", ".xlsx", ".xls"))
    raw = pd.concat([read_file(f) for f in files], ignore_index=True)
    for f in files:
        log.info("read %s", f.name)
    data, issues = clean(raw)
    write_report(data, issues, files, Path(a.output))
    log.info("%d raw rows -> %d clean sales, %d issues flagged -> %s", len(raw), len(data), len(issues), a.output)


if __name__ == "__main__":
    main()
