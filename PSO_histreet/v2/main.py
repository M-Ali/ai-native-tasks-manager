"""
Pricing Intelligence CLI
Usage:
  uv run python main.py scrape    — scrape all competitor prices
  uv run python main.py analyze   — run analysis + generate charts
  uv run python main.py report    — generate Word report
  uv run python main.py all       — scrape → analyze → report
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import click
import yaml
from pathlib import Path
from rich.console import Console
from rich.panel import Panel
from rich.progress import Progress, SpinnerColumn, TextColumn

console = Console()
CONFIG_PATH = Path("config/lubricants.yaml")


def load_config() -> dict:
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


@click.group()
def cli():
    """PSO Pricing Intelligence v2 — competitor price scraper and strategy engine."""
    pass


@cli.command()
@click.option("--platform", default="daraz", help="Platform to scrape (default: daraz)")
@click.option("--brands", default=None, help="Comma-separated brand names to scrape (default: all)")
@click.option("--dry-run", is_flag=True, help="Print search terms without scraping")
def scrape(platform, brands, dry_run):
    """Scrape competitor prices from e-commerce platforms."""
    from scrapers import PLATFORM_MAP
    from normalizer import normalize_all
    from db.store import save, init_db

    config = load_config()
    platform_cfg = config.get("platforms", {}).get(platform, {})

    if not platform_cfg.get("enabled", False):
        console.print(f"[red]Platform '{platform}' is not enabled in config.[/red]")
        return

    ScraperClass = PLATFORM_MAP.get(platform)
    if not ScraperClass:
        console.print(f"[red]No scraper implemented for '{platform}'.[/red]")
        return

    # Build search term list
    filter_brands = [b.strip().lower() for b in brands.split(",")] if brands else None
    search_terms = []

    for competitor in config.get("competitors", []):
        if filter_brands and competitor["name"].lower() not in filter_brands:
            continue
        for oil_type, terms in competitor.get("search_terms", {}).items():
            for term in terms:
                search_terms.append((competitor["name"], oil_type, term))

    console.print(Panel(
        f"Platform: [bold]{platform}[/bold]\n"
        f"Search terms: [bold]{len(search_terms)}[/bold]\n"
        f"Config: {CONFIG_PATH}",
        title="Pricing Intelligence Scraper",
        border_style="green"
    ))

    if dry_run:
        console.print("\n[yellow]DRY RUN — search terms that would be used:[/yellow]")
        for brand, oil_type, term in search_terms:
            console.print(f"  [{oil_type}] {brand}: {term}")
        return

    # Install playwright browsers if needed
    import subprocess
    subprocess.run(
        ["uv", "run", "playwright", "install", "chromium", "--with-deps"],
        capture_output=True
    )

    scraper = ScraperClass(config)
    all_products = []

    with Progress(SpinnerColumn(), TextColumn("{task.description}"), console=console) as progress:
        for brand, oil_type, term in search_terms:
            task = progress.add_task(f"[green]{brand}[/green] — {term}", total=None)
            try:
                products = scraper.scrape(term)
                all_products.extend(products)
            except Exception as e:
                console.print(f"  [red]Skipping '{term}': {str(e)[:80]}[/red]")
            finally:
                progress.update(task, completed=True)

    console.print(f"\n[bold green]Scraped {len(all_products)} raw listings[/bold green]")

    # Normalize
    normalized = normalize_all(all_products)
    matched = [p for p in normalized if p.grade_detected and p.pack_size_l]
    console.print(f"[green]Normalized: {len(matched)} listings with grade + pack size[/green]")

    # Save
    init_db()
    save(normalized)
    console.print(f"[green]Saved to db/prices.db[/green]")

    # Quick summary
    from db.store import latest_run
    import pandas as pd
    rows = latest_run()
    if rows:
        df = pd.DataFrame(rows)
        console.print(f"\nBrands captured: {df['brand_detected'].nunique()}")
        console.print(f"Grades captured: {sorted(df['grade_detected'].dropna().unique())}")


@cli.command()
def analyze():
    """Run pricing analysis and generate charts."""
    config = load_config()
    console.print(Panel("Running pricing analysis...", border_style="green"))

    import analyzer
    analyzer.run_all(config)

    console.print("[bold green]Analysis complete. Charts saved to output/charts/[/bold green]")


@cli.command()
def report():
    """Generate Word document pricing report."""
    config = load_config()
    console.print(Panel("Generating pricing report...", border_style="green"))

    import reporter
    reporter.generate(config)

    console.print("[bold green]Report saved to output/reports/[/bold green]")


@cli.command()
@click.option("--platform", default="daraz")
def all(platform):
    """Full pipeline: scrape → analyze → report."""
    ctx = click.get_current_context()
    console.print(Panel(
        "Running full pipeline: scrape → analyze → report",
        title="Pricing Intelligence v2",
        border_style="green"
    ))
    ctx.invoke(scrape, platform=platform, brands=None, dry_run=False)
    ctx.invoke(analyze)
    ctx.invoke(report)


if __name__ == "__main__":
    cli()
