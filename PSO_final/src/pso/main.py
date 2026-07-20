"""CLI entry point — orchestrates all agents."""

from __future__ import annotations

import os
import sys
from pathlib import Path

import click
from dotenv import load_dotenv
from rich.console import Console
from rich.panel import Panel
from rich.table import Table as RichTable

load_dotenv()
console = Console()


@click.group()
def cli():
    """PSO OMC Analytics Pipeline — powered by uv + Claude AI."""
    pass


@cli.command()
@click.option("--input", "-i", "input_path",
              default="data/input/Working File Retail Fuels Data.xlsx",
              show_default=True, help="Path to the source Excel file.")
@click.option("--output", "-o", "out_dir",
              default="reports", show_default=True, help="Output directory for reports.")
@click.option("--ai/--no-ai", default=True, show_default=True,
              help="Enable/disable Claude AI narrative generation.")
def run(input_path: str, out_dir: str, ai: bool):
    """Full pipeline: Ingest > Analyze > Lubes > Premium Fuel > AI Insights > Excel Report."""
    from pso import ingest, analyze, lubes_analyze, premium_fuel_analyze, ai_insights, excel_report

    console.print(Panel.fit(
        "[bold white]PSO OMC Analytics Pipeline[/]\n"
        "[dim]Ingest > Analyze > Lubes > Premium Fuel > AI > Excel[/]",
        border_style="cyan",
    ))

    # -- Agent 1: Ingest --------------------------------------------------------
    console.rule("[cyan]Step 1 / 6 — Ingest Agent[/]")
    df, quality = ingest.load(input_path)

    # -- Agent 2: Analysis ------------------------------------------------------
    console.rule("[cyan]Step 2 / 6 — Analysis Agent[/]")
    analysis_tables = analyze.run_all(df)
    _print_table_summary(analysis_tables)

    # -- Agent 3: Lubes --------------------------------------------------------
    console.rule("[cyan]Step 3 / 6 — Lubes Agent[/]")
    lubes_tables = lubes_analyze.run_lubes(df)
    _print_table_summary(lubes_tables)

    # -- Agent 3b: Premium Fuel --------------------------------------------------
    console.rule("[cyan]Step 4 / 6 — Premium Fuel Agent[/]")
    premium_tables = premium_fuel_analyze.run_premium_fuel(df)
    _print_table_summary(premium_tables)

    # -- Agent 4: AI Narrator --------------------------------------------------
    console.rule("[cyan]Step 5 / 6 — AI Narrator Agent[/]")
    period = df["_Period"].iloc[0] if "_Period" in df.columns else "Unknown"
    if ai:
        import json
        from datetime import date
        Path(out_dir).mkdir(parents=True, exist_ok=True)
        ai_json = Path(out_dir) / f"PSO_AI_{period}_{date.today()}.json"
        # Load any previously completed sections so we resume rather than redo
        existing_ai: dict = {}
        if ai_json.exists():
            try:
                existing_ai = json.loads(ai_json.read_text(encoding="utf-8"))
                console.print(f"  [dim]Resuming from existing JSON: {ai_json.name}[/]")
            except Exception:
                pass
        ai_output = ai_insights.run_ai_insights(
            analysis_tables, lubes_tables, period, existing=existing_ai
        )
        with open(ai_json, "w", encoding="utf-8") as f:
            json.dump(ai_output, f, ensure_ascii=False, indent=2)
        console.print(f"  AI text saved: {ai_json.name}")
    else:
        console.print("[yellow]AI skipped (--no-ai flag)[/]")
        ai_output = {}

    # -- Agent 5: Report --------------------------------------------------------
    console.rule("[cyan]Step 6 / 6 — Report Agent[/]")
    report_path = excel_report.build(
        analysis_tables, lubes_tables, premium_tables, ai_output, quality, period, out_dir
    )

    console.print(Panel.fit(
        f"[bold green]Done![/]  Report saved to:\n[white]{report_path}[/]",
        border_style="green",
    ))


@cli.command()
@click.option("--input", "-i", "input_path",
              default="data/input/Working File Retail Fuels Data.xlsx",
              show_default=True)
@click.option("--output", "-o", "out_dir", default="reports", show_default=True)
def lubes(input_path: str, out_dir: str):
    """Run lubricants analysis only — faster iteration."""
    from pso import ingest, lubes_analyze, excel_report
    import pandas as pd
    from openpyxl import Workbook
    from datetime import date

    console.print(Panel.fit("[bold white]PSO Lubes-Only Analysis[/]", border_style="yellow"))

    df, quality = ingest.load(input_path)
    period = df["_Period"].iloc[0] if "_Period" in df.columns else "Unknown"

    lubes_tables = lubes_analyze.run_lubes(df)
    _print_table_summary(lubes_tables)

    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    fname = out_dir / f"PSO_Lubes_{period}_{date.today()}.xlsx"

    from pso.excel_report import (
        _write_header, _write_section_title, _write_df, _write_ai_block,
        _sheet_lubes_overview, _sheet_lube_category, PCT_CHG_COLS,
    )

    wb = Workbook()
    wb.remove(wb.active)

    # Reuse the lube sheets from full report
    _sheet_lubes_overview(wb, {}, lubes_tables, period)
    _sheet_lube_category(wb, lubes_tables, "DEO",       "DEO",       period)
    _sheet_lube_category(wb, lubes_tables, "PCMO",      "PCMO",      period)
    _sheet_lube_category(wb, lubes_tables, "MCO",       "MCO",       period)
    _sheet_lube_category(wb, lubes_tables, "LOW GRADE", "LOW_GRADE", period)

    wb.save(fname)
    console.print(f"[bold green]Saved:[/] {fname}")


@cli.command()
@click.option("--input", "-i", "input_path",
              default="data/input/Working File Retail Fuels Data.xlsx",
              show_default=True)
def quality(input_path: str):
    """Run data quality check only — prints report to console."""
    from pso import ingest

    df, q = ingest.load(input_path)

    t = RichTable(title="Data Quality Report", show_header=True, header_style="bold cyan")
    t.add_column("Check", style="white")
    t.add_column("Count", justify="right", style="bold")

    rows = [
        ("Total rows",              str(q["total_rows"])),
        ("Retail Business rows",    str(q["retail_rows"])),
        ("International rows",      str(q["international_rows"])),
        ("Null Sales Org",          str(q["null_org_rows"])),
        ("Null Region (Retail)",    str(q["null_region_retail"])),
        ("Negative Vol CY",         str(q["negative_vol_cy"])),
        ("Zero Vol CY",             str(q["zero_vol_cy_rows"])),
        ("Unnorm cities (distinct)", str(len(q["unnormalized_cities_top30"]))),
    ]
    for label, val in rows:
        t.add_row(label, val)

    console.print(t)

    if q["unnormalized_cities_top30"]:
        console.print("\n[yellow]Top unnormalized cities (add to CITY_NORM in config.py):[/]")
        for city, cnt in list(q["unnormalized_cities_top30"].items())[:20]:
            console.print(f"  [dim]{cnt:>4}x[/]  {city}")


def _print_table_summary(tables: dict) -> None:
    t = RichTable(show_header=True, header_style="bold dim")
    t.add_column("Table", style="cyan")
    t.add_column("Rows", justify="right")
    t.add_column("Cols", justify="right")
    for name, df in tables.items():
        t.add_row(name, str(len(df)), str(len(df.columns)))
    console.print(t)


if __name__ == "__main__":
    cli()
