"""sarsabz-pitch command line."""

from __future__ import annotations

import json
from pathlib import Path

import typer
from rich.console import Console
from rich.panel import Panel
from rich.table import Table

from . import brief, extract, submission_doc

app = typer.Typer(
    add_completion=False,
    help="Toolkit for the Fatima Fertilizer (Sarsabz / Bubber Sher) creative-agency pitch.",
)
console = Console()


def _items(title: str, items) -> Table:
    table = Table(title=title, show_lines=False)
    table.add_column("#", justify="right")
    table.add_column("Item")
    table.add_column("p.", justify="right")
    for n, item in enumerate(items, 1):
        table.add_row(str(n), item.text, str(item.page))
    return table


@app.command()
def info() -> None:
    """The brief at a glance: client, mandate, what to present, and what is missing."""
    b = brief.Brief()
    console.print(Panel(
        f"[bold]{b.title}[/bold]\n{b.client}\n\n"
        f"Brands: {', '.join(b.brands)}\n"
        f"Scope: {'; '.join(b.scope_entities)}\n"
        f"Presentation: {b.presentation_minutes} minutes\n"
        f"[yellow]Pitch date:[/yellow] {b.pitch_date or 'NOT STATED'}   "
        f"[yellow]Budget:[/yellow] {b.budget or 'NOT STATED'}",
        title="Fatima Fertilizer creative pitch"))
    console.print(_items("What the presentation must contain", brief.PITCH_CONTENTS))
    console.print(_items("What they expect from the pitch", brief.EXPECTATIONS))
    console.print("\n[bold]Open questions for the client[/bold]")
    for q in brief.OPEN_QUESTIONS:
        console.print(f"  - {q}")


@app.command()
def case2() -> None:
    """Case 2 - Sarsabz brand audit & tactical campaign: every ask, page-traced."""
    console.print(_items("Case 2 task", brief.CASE2_TASK))
    console.print(_items("Case 2: the agency must recommend", brief.CASE2_RECOMMEND))
    console.print(_items("Case 2: the roadmap must", brief.CASE2_ROADMAP_GOALS))
    console.print(f"Channels to cover: {', '.join(brief.CASE2_CHANNELS)}")


@app.command()
def facts(as_json: bool = typer.Option(False, "--json")) -> None:
    """Brand facts and brand-health numbers read from the brief."""
    if as_json:
        console.print_json(json.dumps(brief.as_dict()))
        return
    console.print(_items("Brand facts", brief.BRAND_FACTS))
    table = Table(title="Master-brand awareness, U&A + Brand Health study (pp.29-31)")
    table.add_column("Brand")
    table.add_column("Measure")
    for year in (2013, 2016, 2018, 2022, 2026):
        table.add_column(str(year), justify="right")
    for name, measures in brief.AWARENESS.items():
        for measure, series in measures.items():
            table.add_row(name, measure,
                          *[f"{series[y]}%" if y in series else "" for y in (2013, 2016, 2018, 2022, 2026)])
    console.print(table)


@app.command()
def parse(
    pdf: Path = typer.Option(None, help="Brief PDF (default: the only PDF in data/)."),
    out: Path = typer.Option(Path("workspace/brief"), help="Where to write full.txt and page images."),
    render_images: bool = typer.Option(True, "--images/--no-images",
                                       help="Render image-only pages to PNG."),
) -> None:
    """Extract the brief's text and render its image-only pages so they can be read."""
    path = extract.find_pdf(pdf)
    doc = extract.load(path)
    out.mkdir(parents=True, exist_ok=True)
    (out / "full.txt").write_text(doc.text, encoding="utf-8")
    image_pages = doc.image_only_pages()
    console.print(f"{path.name}: {len(doc.pages)} pages -> {out / 'full.txt'}")
    console.print(f"Image-only pages (no text layer): {image_pages or 'none'}")
    if render_images and image_pages:
        written = extract.render(path, out / "pages", image_pages)
        console.print(f"Rendered {len(written)} page images -> {out / 'pages'}")


@app.command()
def doc(
    markdown: Path = typer.Argument(..., help="Markdown section to render."),
    stem: str = typer.Option(..., help="Output file stem, e.g. Sarsabz_case2_brand_audit."),
    out: Path = typer.Option(Path("workspace/out"), help="Output directory."),
) -> None:
    """Render a Markdown section to a date-stamped .docx (never overwrites)."""
    path = submission_doc.render(markdown, out, stem)
    console.print(f"[green]wrote[/green] {path}")


if __name__ == "__main__":
    app()
