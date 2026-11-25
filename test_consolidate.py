import pandas as pd
import pytest

from consolidate import branch_name, clean, map_columns, parse_date, parse_number
from pathlib import Path


@pytest.mark.parametrize("raw,expected", [
    ("R$ 1.234,56", 1234.56), ("1.234,56", 1234.56), ("1,234.56", 1234.56), ("89.90", 89.9),
    ("89,90", 89.9), (349, 349.0), ("erro", None), ("", None), (None, None),
])
def test_parse_number(raw, expected):
    assert parse_number(raw) == expected


def test_parse_date_formats():
    assert parse_date("31/01/2026") == pd.Timestamp("2026-01-31")
    assert parse_date("2026-01-31") == pd.Timestamp("2026-01-31")
    assert parse_date("not a date") is None


def test_map_columns_handles_accents_and_aliases():
    m = map_columns(["Data", "Produto", "Qtd", "Preço Unitário"])
    assert sorted(m.values()) == ["date", "product", "quantity", "unit_price"]
    with pytest.raises(ValueError):
        map_columns(["Data", "Produto"])


def test_branch_name():
    assert branch_name(Path("vendas_sao_paulo.csv")) == "São Paulo"


def test_clean_flags_issues_and_removes_duplicates():
    raw = pd.DataFrame({
        "date": ["01/02/2026", "01/02/2026", "xx", "02/02/2026"],
        "product": ["  mouse wireless ", "MOUSE WIRELESS", "Ssd 1tb", "Webcam Hd"],
        "quantity": ["2", "2", "1", ""],
        "unit_price": ["R$ 89,90", "89.90", "459", "279"],
        "branch": "X", "source_file": "f.csv", "source_row": [2, 3, 4, 5],
    })
    data, issues = clean(raw)
    assert len(data) == 1  # 1 valid; 1 duplicate (same after normalization); 2 invalid
    assert data.loc[0, "product"] == "Mouse Wireless" and data.loc[0, "total"] == 179.8
    assert sorted(issues["issue"]) == ["duplicate row (removed)", "invalid date", "missing/invalid quantity"]
