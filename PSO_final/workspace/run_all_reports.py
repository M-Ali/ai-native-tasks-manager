"""Push-button regeneration of every PSO lubricants report.

Usage:
    uv run python workspace/run_all_reports.py
    uv run python workspace/run_all_reports.py --input "data/input/Working File Retail Fuels Data FY27.xlsx"

Runs each report-generation script in sequence against the given input
file. Output filenames are auto-tagged with the period read from the
source file's sheet name (see _pso_common.py), so re-runs never silently
overwrite a prior period's reports.
"""
import argparse
import os
import subprocess
import sys

SCRIPTS = [
    'lubes_report.py',
    'lubes_stations_analysis.py',
    'lubes_stations_report.py',
    'city_profiles.py',
    'city_profiles_volume.py',
    'lubes_vol_table.py',
    'lubes_vol_uplift.py',
    'national_vol_slide.py',
]


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--input', '-i', default=None,
                     help='Path to the source Excel file. Defaults to the '
                          'PSO_INPUT env var, or data/input/Working File '
                          'Retail Fuels Data.xlsx if unset.')
    args = ap.parse_args()

    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    env = os.environ.copy()
    if args.input:
        env['PSO_INPUT'] = args.input

    print(f"Input file: {env.get('PSO_INPUT', '(default — see _pso_common.py)')}")

    failures = []
    for script in SCRIPTS:
        path = os.path.join(root, 'workspace', script)
        print(f"\n{'=' * 70}\n  Running {script}\n{'=' * 70}")
        result = subprocess.run([sys.executable, path], cwd=root, env=env)
        if result.returncode != 0:
            failures.append(script)
            print(f"  !! {script} FAILED (exit {result.returncode})")

    print(f"\n{'=' * 70}")
    if failures:
        print(f"  Done with {len(failures)} failure(s): {failures}")
        sys.exit(1)
    print("  All reports regenerated successfully.")


if __name__ == '__main__':
    main()
