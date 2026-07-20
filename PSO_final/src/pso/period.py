"""Period Agent — defines the prior-period comparison basis for the whole pipeline.

CY covers a partial year (e.g. 10 months of FY26). Comparing it to full-year LY
(12 months of FY25) overstates or understates real YoY movement because the periods
are not equal length. Every comparison in this pipeline uses SPLY (Same Period Last
Year — the same 10-month window in FY25) instead. SPLY absolute columns are derived
by pso.ingest from the source file's %SPLY* columns (CY / (1 + %SPLY/100)).

This module is the single place that defines "prior period" — if the comparison
basis ever needs to change again, it changes here only.
"""

from __future__ import annotations

from pso.config import (
    COL_VOL_CY, COL_GRS_CY, COL_NMGN_CY, COL_DISC_CY, COL_PMGN_CY,
    COL_VOL_SPLY, COL_GRS_SPLY, COL_NMGN_SPLY, COL_DISC_SPLY, COL_PMGN_SPLY,
)

# Canonical "prior period" column for each CY metric — SPLY, not full-year LY.
PRIOR_COL: dict[str, str] = {
    COL_VOL_CY:  COL_VOL_SPLY,
    COL_GRS_CY:  COL_GRS_SPLY,
    COL_NMGN_CY: COL_NMGN_SPLY,
    COL_DISC_CY: COL_DISC_SPLY,
    COL_PMGN_CY: COL_PMGN_SPLY,
}

# Display label used in headers/titles/prompts wherever "prior period" is shown.
PRIOR_LABEL = "SPLY"
