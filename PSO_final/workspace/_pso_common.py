"""Shared helpers for PSO report-generation scripts.

Centralizes the input file path and period labelling so a new period's
data (e.g. FY27) can be dropped in without editing every script by hand.

Set the PSO_INPUT env var to point at a new source file; otherwise falls
back to the default path. The period tag/label is derived from the source
file's sheet name (captured by pso.ingest as df['_Period']), so report
covers, footers, and output filenames always track whatever file was
actually loaded — not a hardcoded string.
"""
import os
from datetime import date

DEFAULT_INPUT = 'data/input/Working File Retail Fuels Data.xlsx'
INPUT_PATH = os.environ.get('PSO_INPUT', DEFAULT_INPUT)


def get_period(df) -> str:
    """Filesystem-safe period tag, e.g. '10M_FY26'."""
    return df['_Period'].iloc[0] if '_Period' in df.columns else 'UnknownPeriod'


def get_period_label(df) -> str:
    """Human-readable label for report covers/footers, e.g. '10M FY26 (generated June 2026)'."""
    return f"{get_period(df).replace('_', ' ')} (generated {date.today():%B %Y})"


def out_path(stem: str, ext: str, df, out_dir: str = 'reports') -> str:
    """Period-tagged output path, e.g. reports/PSO_Lubricants_Report_10M_FY26.docx.

    Tagging prevents a new period's run from silently overwriting the
    previous period's reports.
    """
    os.makedirs(out_dir, exist_ok=True)
    return f"{out_dir}/{stem}_{get_period(df)}.{ext}"
