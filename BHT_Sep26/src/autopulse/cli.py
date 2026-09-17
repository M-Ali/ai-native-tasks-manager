"""Command line entry point: uv run autopulse <command>."""

from __future__ import annotations

import argparse
import sys

from autopulse import __version__
from autopulse.monitor import build_report
from autopulse.output import write_monitor
from autopulse.sales import SalesDataError, load_sales


def cmd_validate(args) -> int:
    df = load_sales(args.file)
    months = sorted(df["month"].unique())
    print(f"OK  {len(df)} rows, {df['brand'].nunique()} brands, {df['model'].nunique()} models")
    print(f"    months {months[0]} to {months[-1]} ({len(months)})")
    return 0


def cmd_monitor(args) -> int:
    report = build_report(load_sales(args.file), args.month)
    print(f"Market monitor {report.month}: {report.segments['units'].sum():,} units")
    for note in report.notes:
        print(f"  note: {note}")
    cols = ["brand", "units", "share_of_market_pct", "mom_change_pct", "yoy_change_pct"]
    print(report.brands[cols].to_string(index=False))
    print(f"\nwrote {write_monitor(report, args.file)}")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="autopulse", description=__doc__)
    parser.add_argument("--version", action="version", version=__version__)
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("validate", help="check a monthly sales CSV")
    p.add_argument("file")
    p.set_defaults(func=cmd_validate)

    p = sub.add_parser("monitor", help="share and growth by brand, segment and model")
    p.add_argument("file")
    p.add_argument("--month", help="YYYY-MM; defaults to the latest month in the file")
    p.set_defaults(func=cmd_monitor)

    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except (ValueError, FileNotFoundError) as exc:  # SalesDataError is a ValueError
        print(f"error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
