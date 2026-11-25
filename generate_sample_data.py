"""Generates SYNTHETIC messy sales files (one per branch) to demo the consolidation.

Each branch exports differently, as in real life:
- Sao Paulo: CSV with ';', Brazilian formats (R$ 1.234,56, dd/mm/yyyy), Portuguese headers.
- Rio: Excel with English headers, ISO dates and plain numbers.
- Curitiba: CSV with ',', different header names, blank lines, duplicates, typos, bad values.
"""

import csv
import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook

random.seed(42)
OUT = Path("input")
PRODUCTS = {"Notebook Pro 14": 4899.0, "Mouse Wireless": 89.9, "Teclado Mecanico": 349.0,
            "Monitor 27": 1599.0, "Headset USB": 229.0, "Webcam HD": 279.0, "SSD 1TB": 459.0}
START = date(2026, 1, 1)


def rows(n: int):
    for _ in range(n):
        p = random.choice(list(PRODUCTS))
        q = random.choices([1, 2, 3, 5], weights=[60, 25, 10, 5])[0]
        price = round(PRODUCTS[p] * random.uniform(0.92, 1.05), 2)
        yield START + timedelta(days=random.randint(0, 180)), p, q, price


def br_money(v: float) -> str:
    return "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def main() -> None:
    OUT.mkdir(exist_ok=True)

    with (OUT / "vendas_sao_paulo.csv").open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Data", "Produto", "Qtd", "Preço Unitário"])
        for d, p, q, pr in rows(420):
            w.writerow([d.strftime("%d/%m/%Y"), p, q, br_money(pr)])

    wb = Workbook()
    ws = wb.active
    ws.append(["Sale Date", "Item", "Quantity", "Unit Price"])
    for d, p, q, pr in rows(360):
        ws.append([d, p, q, pr])
    wb.save(OUT / "sales_rio.xlsx")

    with (OUT / "curitiba_export.csv").open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["dt_venda", "descricao_produto", "quantidade", "valor_unit"])
        data = list(rows(300))
        data += random.sample(data, 15)  # duplicated lines (re-exported)
        for i, (d, p, q, pr) in enumerate(data):
            name = p.upper() if i % 7 == 0 else f"  {p} "  # inconsistent casing/spaces
            qty = "" if i % 97 == 0 else q  # missing quantity
            price = "erro" if i % 113 == 0 else f"{pr:.2f}"  # invalid value
            w.writerow([d.strftime("%d/%m/%Y"), name, qty, price])
            if i % 50 == 0:
                w.writerow([])  # blank lines
    print("synthetic files written to", OUT.resolve())


if __name__ == "__main__":
    main()
