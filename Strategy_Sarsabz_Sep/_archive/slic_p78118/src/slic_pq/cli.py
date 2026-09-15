"""slic-pq command line."""

from __future__ import annotations

import json
from pathlib import Path

import typer
from rich.console import Console
from rich.panel import Panel
from rich.table import Table

from . import criteria, extract, readiness, reports, submission_doc

app = typer.Typer(
    add_completion=False,
    help="Bid-readiness toolkit for the SLIC advertising-agency prequalification (P78118).",
)
console = Console()

STATUS_COLOR = {"ready": "green", "partial": "yellow", "missing": "red", "na": "dim"}


@app.command()
def info() -> None:
    """Show the tender at a glance: who, what, when, how it is scored."""
    t = criteria.Tender()
    console.print(
        Panel(
            f"[bold]{t.title}[/bold]\n"
            f"Reference: {t.reference}   Term: {t.engagement_term}\n"
            f"{t.agency}\n{t.division}\n"
            f"{t.contact_email} / {t.contact_phone}\n\n"
            f"Method: {t.method}\nSelection: {t.selection}\n"
            f"Portal: {t.portal}\n\n"
            f"[yellow]Clarification / pre-application:[/yellow] {t.clarification_deadline}\n"
            f"[yellow]Submission deadline:[/yellow] {t.submission_deadline} (EPADS v2.0 only)\n"
            f"[yellow]Opening:[/yellow] {t.opening}\n"
            f"[yellow]Pass mark:[/yellow] {t.passing_marks} of {t.total_marks}",
            title="SLIC prequalification",
        )
    )
    table = Table(title="Marks by bucket")
    table.add_column("Bucket")
    table.add_column("Marks", justify="right")
    for bucket, marks in criteria.marks_by_bucket().items():
        table.add_row(criteria.BUCKET_LABELS[bucket], str(marks))
    table.add_row("[bold]Total[/bold]", f"[bold]{criteria.total_marks()}[/bold]")
    console.print(table)


@app.command("criteria")
def criteria_cmd(
    as_json: bool = typer.Option(False, "--json", help="Emit the criteria as JSON."),
) -> None:
    """List the eligibility documents and the scored technical criteria."""
    if as_json:
        console.print_json(json.dumps(criteria.as_dict()))
        return

    table = Table(title="Eligibility documents (pass/fail)")
    table.add_column("#", justify="right")
    table.add_column("Requirement")
    table.add_column("PQ page", justify="right")
    for n, item in enumerate(criteria.ELIGIBILITY, 1):
        table.add_row(str(n), item.requirement, str(item.page))
    console.print(table)

    table = Table(title=f"Technical evaluation ({criteria.total_marks()} marks)")
    table.add_column("Bucket")
    table.add_column("Criterion")
    table.add_column("Marks", justify="right")
    table.add_column("Evidence required")
    for item in criteria.SCORED:
        table.add_row(
            criteria.BUCKET_LABELS[item.bucket], item.question, str(item.marks), item.evidence
        )
    console.print(table)


@app.command()
def parse(
    pdf: Path = typer.Option(None, "--pdf", help="Path to the PQ PDF (default: the one in data/)."),
    section: str = typer.Option(None, "--section", help="Print one section only."),
    engine: str = typer.Option("auto", "--engine", help="auto | pdftotext | pypdf."),
    out: Path = typer.Option(None, "--out", help="Write the extracted text to this file."),
) -> None:
    """Extract the prequalification PDF and slice it into named sections."""
    path = extract.find_pdf(pdf)
    doc = extract.load(path, engine=engine)
    sections = doc.sections()
    console.print(f"[dim]{path} - {len(doc.pages)} pages via {doc.engine}[/dim]")

    if section:
        if section not in sections:
            raise typer.BadParameter(
                f"Unknown section {section!r}. Available: {', '.join(sections)}"
            )
        text = sections[section]
    else:
        table = Table(title="Sections found")
        table.add_column("Key")
        table.add_column("Characters", justify="right")
        for key, body in sections.items():
            table.add_row(key, str(len(body)))
        console.print(table)
        dates = extract.key_dates(doc)
        if dates:
            console.print(f"[dim]Dates in the PDS: {dates}[/dim]")
        text = doc.clean_text()

    if out:
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(text, encoding="utf-8")
        console.print(f"[green]Wrote[/green] {out}")
    elif section:
        console.print(text)


@app.command()
def init(
    agency: str = typer.Option("<your agency name>", "--agency", help="Applicant agency name."),
    path: Path = typer.Option(readiness.STATUS_PATH, "--path"),
    force: bool = typer.Option(False, "--force", help="Overwrite an existing status file."),
) -> None:
    """Create workspace/status.yaml, the self-assessment you fill in."""
    try:
        written = readiness.init(path, agency=agency, force=force)
    except FileExistsError as err:
        raise typer.BadParameter(str(err)) from err
    console.print(f"[green]Created[/green] {written}")
    console.print("Set each item to ready / partial / missing / na, then run [bold]slic-pq score[/bold].")


@app.command()
def score(path: Path = typer.Option(readiness.STATUS_PATH, "--path")) -> None:
    """Score the filled-in self-assessment against the 100 technical marks."""
    a = _load(path)
    table = Table(title=f"{a.agency or 'Applicant'} - technical scoring")
    table.add_column("Bucket")
    table.add_column("Criterion")
    table.add_column("Marks", justify="right")
    table.add_column("Status")
    table.add_column("Earned", justify="right")
    for line in a.lines:
        colour = STATUS_COLOR.get(line.status, "white")
        table.add_row(
            criteria.BUCKET_LABELS.get(line.bucket, line.bucket),
            line.label,
            str(line.marks),
            f"[{colour}]{line.status}[/{colour}]",
            f"{line.earned:g}",
        )
    console.print(table)

    summary = Table(title="Summary")
    summary.add_column("Bucket")
    summary.add_column("Earned", justify="right")
    summary.add_column("Available", justify="right")
    for bucket, (got, avail) in a.by_bucket().items():
        summary.add_row(criteria.BUCKET_LABELS.get(bucket, bucket), f"{got:g}", str(avail))
    summary.add_row("[bold]Total[/bold]", f"[bold]{a.earned:g}[/bold]", f"[bold]{a.available}[/bold]")
    console.print(summary)

    pass_mark = criteria.Tender().passing_marks
    if a.passes:
        console.print(f"[green]Above the pass mark[/green] ({a.earned:g} vs {pass_mark}).")
    else:
        console.print(
            f"[red]Below the pass mark[/red] ({a.earned:g} vs {pass_mark}) - "
            f"{pass_mark - a.earned:g} marks short."
        )


@app.command()
def gaps(path: Path = typer.Option(readiness.STATUS_PATH, "--path")) -> None:
    """List what is costing marks and which eligibility documents are outstanding."""
    a = _load(path)
    eg = a.eligibility_gaps()
    if eg:
        table = Table(title="Eligibility documents outstanding (pass/fail)")
        table.add_column("Requirement")
        table.add_column("Status")
        table.add_column("Note")
        for line in eg:
            colour = STATUS_COLOR.get(line.status, "white")
            table.add_row(line.label, f"[{colour}]{line.status}[/{colour}]", line.note)
        console.print(table)
    else:
        console.print("[green]All eligibility documents marked ready.[/green]")

    scored = criteria.scored_by_key()
    rows = a.gaps()
    if not rows:
        console.print("[green]No marks being lost.[/green]")
        return
    table = Table(title="Marks at stake, worst first")
    table.add_column("Criterion")
    table.add_column("At stake", justify="right")
    table.add_column("Status")
    table.add_column("Next action")
    for line in rows:
        colour = STATUS_COLOR.get(line.status, "white")
        table.add_row(
            line.label,
            f"{line.marks - line.earned:g}",
            f"[{colour}]{line.status}[/{colour}]",
            line.note or scored[line.key].evidence,
        )
    console.print(table)


@app.command()
def build(
    path: Path = typer.Option(readiness.STATUS_PATH, "--path"),
    out_dir: Path = typer.Option(reports.OUT_DIR, "--out-dir"),
    forms: bool = typer.Option(True, "--forms/--no-forms", help="Also build the Annexure forms."),
) -> None:
    """Build the deliverables: checklist (Markdown + Excel) and the Annexure forms (Word)."""
    a = _load(path)
    written = [
        reports.checklist_markdown(a, out_dir),
        reports.checklist_excel(a, out_dir),
    ]
    if forms:
        written.append(reports.annexure_forms(a.agency, out_dir))
    for item in written:
        console.print(f"[green]Wrote[/green] {item}")


@app.command()
def doc(
    markdown: Path = typer.Argument(..., help="Markdown section to render."),
    stem: str = typer.Option(None, "--stem", help="Output filename stem. Defaults to the markdown stem."),
    out_dir: Path = typer.Option(reports.OUT_DIR, "--out-dir"),
) -> None:
    """Render a written submission section (Markdown) as a Word document.

    Every scored criterion is (Qualitative)(Doc Required), so each written section has to
    be uploaded as a document. Keeping the Markdown as the source means the working file
    and the submitted file cannot drift apart.
    """
    if not markdown.exists():
        raise typer.BadParameter(f"{markdown} not found")
    written = submission_doc.render(markdown, out_dir, stem or markdown.stem)
    console.print(f"[green]Wrote[/green] {written}")


def _load(path: Path):
    try:
        return readiness.load(path)
    except (FileNotFoundError, ValueError) as err:
        raise typer.BadParameter(str(err)) from err


if __name__ == "__main__":
    app()
