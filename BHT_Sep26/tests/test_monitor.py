"""Tests use made-up numbers only. They are not market data."""

import pandas as pd
import pytest

from autopulse.monitor import build_report
from autopulse.output import next_version
from autopulse.sales import SalesDataError, load_sales

HEADER = "month,brand,model,segment,units,source\n"


def write(tmp_path, body):
    path = tmp_path / "sales.csv"
    path.write_text(HEADER + body)
    return path


@pytest.fixture
def sales(tmp_path):
    return load_sales(write(tmp_path, (
        "2025-08,BrandA,A1,Hatchback,50,test\n"
        "2025-08,BrandB,B1,Hatchback,50,test\n"
        "2026-07,BrandA,A1,Hatchback,100,test\n"
        "2026-07,BrandB,B1,Hatchback,100,test\n"
        "2026-08,BrandA,A1,Hatchback,150,test\n"
        "2026-08,BrandB,B1,Hatchback,50,test\n"
        "2026-08,BrandB,B2,Sedan,\"1,00\",test\n"
    )))


def test_shares_and_growth(sales):
    r = build_report(sales)
    assert str(r.month) == "2026-08"
    a = r.brands.set_index("brand").loc["BrandA"]
    assert a["units"] == 150
    assert a["share_of_market_pct"] == 50.0
    assert a["mom_change_pct"] == 50.0
    assert a["yoy_change_pct"] == 200.0
    assert a["share_pts_vs_prev_month"] == 0.0
    b2 = r.models.set_index("model").loc["B2"]
    assert b2["share_of_segment_pct"] == 100.0
    assert pd.isna(b2["mom_change_pct"])  # zero last month: growth undefined, not infinite
    assert r.notes == []


def test_missing_base_month_is_flagged(sales):
    r = build_report(sales, "2026-07")
    assert any("2026-06" in n for n in r.notes)
    assert r.brands["mom_change_pct"].isna().all()


def test_validation_reports_every_problem(tmp_path):
    path = write(tmp_path, (
        "2026-13,BrandA,A1,Hatchback,10,test\n"
        "2026-08,BrandA,A1,Hatchback,-5,test\n"
        "2026-08,BrandA,A1,,10,test\n"
    ))
    with pytest.raises(SalesDataError) as exc:
        load_sales(path)
    msg = str(exc.value)
    for expected in ["month not YYYY-MM", "units not a whole number", "blank segment", "duplicate"]:
        assert expected in msg


def test_outputs_are_versioned(tmp_path):
    first = next_version("x", ".xlsx", tmp_path)
    first.write_text("")
    assert next_version("x", ".xlsx", tmp_path).name == "x-v2.xlsx"
