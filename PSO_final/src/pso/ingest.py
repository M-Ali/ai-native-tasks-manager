"""Ingest Agent — load, validate, normalize, classify."""

from __future__ import annotations

import numpy as np
import pandas as pd
from pathlib import Path
from rich.console import Console

from pso.config import (
    HEADER_ROW, REQUIRED_COLUMNS,
    COL_ORG, COL_PRODUCT, COL_CATEGORY, COL_REGION, COL_CITY,
    COL_CORP_GRP, COL_CUST_NAME,
    COL_GRS_CY, COL_GRS_LY, COL_VOL_CY, COL_VOL_LY,
    COL_DISC_CY, COL_DISC_LY, COL_PMGN_CY, COL_PMGN_LY,
    COL_NMGN_CY, COL_NMGN_LY,
    COL_FUEL_SEG, COL_LUBE_CAT, COL_CITY_NORM, COL_IS_RETAIL, COL_IS_INTL,
    PRODUCT_SEGMENTS, LUBE_PRODUCTS, CITY_NORM, INTERNATIONAL_CITIES,
    RETAIL_ORG, ADDITIVE_METRICS,
    COL_PCT_SPLY_VOL, COL_PCT_SPLY_GRS, COL_PCT_SPLY_NMGN,
    COL_PCT_SPLY_DISC, COL_PCT_SPLY_PMGN,
    COL_VOL_SPLY, COL_GRS_SPLY, COL_NMGN_SPLY,
    COL_DISC_SPLY, COL_PMGN_SPLY,
)

console = Console()


def load(path: str | Path) -> tuple[pd.DataFrame, dict]:
    """
    Load the PSO Excel file, validate schema, normalize, and classify.

    Returns (df, quality_report).
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")

    console.print(f"[bold cyan]Loading[/] {path.name} …")

    # The Excel has a blank row 1 and headers in row 2 (0-indexed: row 1)
    raw = pd.read_excel(path, sheet_name=0, header=HEADER_ROW, engine="openpyxl")

    period = _read_period(path)
    console.print(f"  Period: [bold]{period}[/]  |  Raw rows: [bold]{len(raw):,}[/]")

    _validate_schema(raw)

    # Fill numeric NaN with 0 for additive metrics only
    for col in ADDITIVE_METRICS:
        if col in raw.columns:
            raw[col] = pd.to_numeric(raw[col], errors="coerce").fillna(0)

    # String columns — strip whitespace
    for col in [COL_ORG, COL_PRODUCT, COL_CATEGORY, COL_REGION, COL_CITY,
                COL_CORP_GRP, COL_CUST_NAME]:
        if col in raw.columns:
            raw[col] = raw[col].astype(str).str.strip()
            raw[col] = raw[col].replace("nan", "")

    quality = _quality_report(raw)

    raw = _add_derived_columns(raw)
    raw["_Period"] = period

    console.print(f"  Clean rows: [bold green]{len(raw):,}[/]")
    _print_quality(quality)

    return raw, quality


def _read_period(path: Path) -> str:
    """Read sheet name to get period label (e.g. '10M_FY26')."""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True)
    name = wb.sheetnames[0]
    wb.close()
    return name


def _validate_schema(df: pd.DataFrame) -> None:
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            f"Source file is missing required columns:\n  {missing}\n"
            "Check AGENTS.md for the expected schema."
        )


def _add_derived_columns(df: pd.DataFrame) -> pd.DataFrame:
    # FuelSegment
    unmapped = set(df[COL_PRODUCT].unique()) - set(PRODUCT_SEGMENTS.keys()) - {""}
    if unmapped:
        console.print(
            f"  [yellow]Warning:[/] {len(unmapped)} unmapped product codes > "
            f"'Unknown': {unmapped}"
        )
    df[COL_FUEL_SEG] = df[COL_PRODUCT].map(PRODUCT_SEGMENTS).fillna("Unknown")

    # LubeCategory — only meaningful for Lubricant rows
    df[COL_LUBE_CAT] = df.apply(
        lambda r: r[COL_CATEGORY] if r[COL_FUEL_SEG] == "Lubricants" and r[COL_CATEGORY] else "",
        axis=1,
    )

    # IsRetail
    df[COL_IS_RETAIL] = df[COL_ORG] == RETAIL_ORG

    # IsInternational
    df[COL_IS_INTL] = df[COL_CITY].str.upper().isin(
        {c.upper() for c in INTERNATIONAL_CITIES}
    )

    # CityNorm — canonical city name
    df[COL_CITY_NORM] = df[COL_CITY].apply(
        lambda c: CITY_NORM.get(c, c)   # keep original if not in table
    )

    # SPLY absolute columns — derived from CY ÷ (1 + %SPLY/100)
    df = _derive_sply_columns(df)

    return df


def _derive_sply_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Derive Same Period Last Year absolute values from CY and %SPLY columns.

    Formula: SPLY = CY / (1 + %SPLY/100)
    NaN %SPLY means no prior-period data (new station/product) → SPLY = 0.
    """
    pairs = [
        (COL_VOL_CY,  COL_PCT_SPLY_VOL,  COL_VOL_SPLY),
        (COL_GRS_CY,  COL_PCT_SPLY_GRS,  COL_GRS_SPLY),
        (COL_NMGN_CY, COL_PCT_SPLY_NMGN, COL_NMGN_SPLY),
        (COL_DISC_CY, COL_PCT_SPLY_DISC, COL_DISC_SPLY),
        (COL_PMGN_CY, COL_PCT_SPLY_PMGN, COL_PMGN_SPLY),
    ]
    for cy_col, pct_col, sply_col in pairs:
        if pct_col in df.columns and cy_col in df.columns:
            pct    = pd.to_numeric(df[pct_col], errors="coerce")
            factor = 1.0 + pct / 100.0
            df[sply_col] = np.where(
                pct.notna() & (factor.abs() > 0.01),
                df[cy_col] / factor,
                0.0,
            )
        else:
            df[sply_col] = 0.0
    return df


def _quality_report(df: pd.DataFrame) -> dict:
    retail_mask = df[COL_ORG] == RETAIL_ORG

    unnorm_cities = (
        df[COL_CITY]
        .loc[~df[COL_CITY].isin(CITY_NORM) & (df[COL_CITY] != "") & ~df[COL_CITY].str.upper().isin({c.upper() for c in INTERNATIONAL_CITIES})]
        .value_counts()
        .head(30)
        .to_dict()
    )

    return {
        "total_rows":              len(df),
        "retail_rows":             int(retail_mask.sum()),
        "null_org_rows":           int(df[COL_ORG].eq("").sum()),
        "null_region_retail":      int((retail_mask & df[COL_REGION].eq("")).sum()),
        "international_rows":      int(df[COL_CITY].str.upper().isin({c.upper() for c in INTERNATIONAL_CITIES}).sum()),
        "negative_vol_cy":         int((df[COL_VOL_CY] < 0).sum()),
        "negative_margin_cy":      int((df[COL_NMGN_CY] < 0).sum()),
        "unnormalized_cities_top30": unnorm_cities,
        "zero_vol_cy_rows":        int((df[COL_VOL_CY] == 0).sum()),
    }


def _print_quality(q: dict) -> None:
    console.print("  [bold]Data Quality[/]")
    console.print(f"    Total rows           : {q['total_rows']:,}")
    console.print(f"    Retail Business rows : {q['retail_rows']:,}")
    console.print(f"    International rows   : {q['international_rows']:,}")
    console.print(f"    Null Sales Org rows  : {q['null_org_rows']:,}")
    console.print(f"    Null Region (Retail) : {q['null_region_retail']:,}")
    console.print(f"    Negative Vol CY rows : {q['negative_vol_cy']:,}")
    console.print(f"    Zero Vol CY rows     : {q['zero_vol_cy_rows']:,}")
    n_unnorm = len(q["unnormalized_cities_top30"])
    if n_unnorm:
        console.print(f"    Cities not in norm table (top {n_unnorm}): "
                      f"{list(q['unnormalized_cities_top30'].keys())[:10]} …")
