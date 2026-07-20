"""AI Narrator Agent — multi-provider: OpenAI / Gemini / Anthropic."""

from __future__ import annotations

import os
import pandas as pd
from rich.console import Console

from pso.config import AI_MAX_TOKENS

console = Console()

_SYSTEM = """You are a senior petroleum industry analyst reviewing PSO (Pakistan State Oil)
OMC performance data for Pakistan. You receive structured summary tables extracted directly
from sales data covering 10 months of FY2026.

Rules you must follow:
1. Base every observation strictly on the numbers provided — no industry benchmarks, no assumptions.
2. If the data does not explain a pattern, say so explicitly rather than guessing.
3. Be specific: name the exact segment, region, or city when making a point.
4. Flag data anomalies if you see them (zeros, extreme outliers, inconsistencies).
5. Lubricants is PSO's primary business problem in this dataset — prioritize those findings.
6. Structure output as: numbered FINDINGS, then numbered RECOMMENDATIONS.
7. Concise, direct language. No filler. Numbers must match what's in the tables.
"""


def _df_to_md(df: pd.DataFrame, max_rows: int = 40) -> str:
    if df.empty:
        return "(no data)"
    return df.head(max_rows).to_markdown(index=False, floatfmt=".3f")


# ── Provider implementations ──────────────────────────────────────────────────

def _call_openai(prompt: str, client) -> str:
    resp = client.chat.completions.create(
        model="gpt-4o",
        max_tokens=AI_MAX_TOKENS,
        messages=[
            {"role": "system", "content": _SYSTEM},
            {"role": "user",   "content": prompt},
        ],
        temperature=0.2,
    )
    return resp.choices[0].message.content


def _call_gemini(prompt: str, client) -> str:
    from google.genai import types as gtypes
    full_prompt = f"{_SYSTEM}\n\n{prompt}"
    resp = client.models.generate_content(
        model="gemini-2.0-flash",
        contents=full_prompt,
        config=gtypes.GenerateContentConfig(
            max_output_tokens=AI_MAX_TOKENS,
            temperature=0.2,
        ),
    )
    return resp.text


def _call_anthropic(prompt: str, client) -> str:
    msg = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=AI_MAX_TOKENS,
        system=_SYSTEM,
        messages=[{"role": "user", "content": prompt}],
    )
    return msg.content[0].text


def _build_caller():
    """Auto-detect available provider from env. Returns (call_fn, provider_name)."""
    provider = os.getenv("AI_PROVIDER", "").lower()

    # Explicit provider override
    if provider == "openai" or (not provider and os.getenv("OPENAI_API_KEY")):
        key = os.getenv("OPENAI_API_KEY")
        if key:
            import openai
            client = openai.OpenAI(api_key=key)
            console.print("  [bold green]AI Provider:[/] OpenAI GPT-4o")
            return lambda p: _call_openai(p, client), "OpenAI GPT-4o"

    if provider == "gemini" or (not provider and os.getenv("GEMINI_API_KEY")):
        key = os.getenv("GEMINI_API_KEY")
        if key:
            # Override system GOOGLE_API_KEY so SDK uses the correct key
            os.environ["GOOGLE_API_KEY"] = key
            from google import genai as google_genai
            client = google_genai.Client(api_key=key)
            console.print("  [bold green]AI Provider:[/] Google Gemini 2.0 Flash")
            return lambda p: _call_gemini(p, client), "Gemini 2.0 Flash"

    if provider == "anthropic" or (not provider and os.getenv("ANTHROPIC_API_KEY")):
        key = os.getenv("ANTHROPIC_API_KEY")
        if key:
            import anthropic
            client = anthropic.Anthropic(api_key=key)
            console.print("  [bold green]AI Provider:[/] Anthropic Claude Sonnet")
            return lambda p: _call_anthropic(p, client), "Claude Sonnet"

    return None, None


# ── Section prompts ───────────────────────────────────────────────────────────

def _prompt_exec_summary(analysis: dict, period: str) -> str:
    return f"""Period: {period}

## PSO Portfolio — All Business Channels (PKR Billions / Million Litres)

{_df_to_md(analysis.get('portfolio_summary', pd.DataFrame()))}

## Retail Business — Segment Split

{_df_to_md(analysis.get('retail_segment_split', pd.DataFrame()))}

Analyze:
1. Which channels are growing vs declining in BOTH volume and margin?
2. Where is PSO gaining revenue but losing margin (price inflation masking volume loss)?
3. What is the overall portfolio health signal for {period}?
4. Top 3 concerns visible from this data.

Numbered FINDINGS then numbered RECOMMENDATIONS."""


def _prompt_diesel(analysis: dict, period: str) -> str:
    return f"""Period: {period}

## Diesel (HSD + LDO) by Region

{_df_to_md(analysis.get('diesel_by_region', pd.DataFrame()))}

## Diesel — Top 25 Cities by Volume (Pareto)

{_df_to_md(analysis.get('diesel_pareto', pd.DataFrame()).head(25))}

## Diesel — HSD vs LDO by Region

{_df_to_md(analysis.get('diesel_by_product', pd.DataFrame()))}

Analyze:
1. Which regions drive diesel volume; which are declining?
2. How many cities = 50% of total diesel volume? How many = 80%?
3. Is LDO trending differently from HSD — where and by how much?
4. Where is Net Margin/Ltr under pressure vs holding?
5. What does city concentration imply for PSO's diesel coverage strategy?

Numbered FINDINGS then numbered RECOMMENDATIONS."""


def _prompt_petrol(analysis: dict, period: str) -> str:
    return f"""Period: {period}

## Petrol (PMG + R95) by Region

{_df_to_md(analysis.get('petrol_by_region', pd.DataFrame()))}

## Petrol — Top 25 Cities by Volume (Pareto)

{_df_to_md(analysis.get('petrol_pareto', pd.DataFrame()).head(25))}

## Petrol — PMG vs R95 by Region

{_df_to_md(analysis.get('petrol_by_product', pd.DataFrame()))}

Analyze:
1. PMG vs R95 mix shift — which grade is growing, which declining, and by how much?
2. Which regions are petrol strongholds vs weak spots?
3. How concentrated is petrol volume — cities covering 50% and 80%?
4. Regions with declining volume but improving margin/ltr — premiumisation signal?
5. What does R95 trend suggest about consumer behaviour (trading up or down)?

Numbered FINDINGS then numbered RECOMMENDATIONS."""


def _prompt_lubes_problem(lubes: dict, period: str) -> str:
    return f"""Period: {period}

## Lubricants — Category Trend (ALL LUBES total + each type)

{_df_to_md(lubes.get('lube_category_trend', pd.DataFrame()))}

## Lubricants — Margin Decomposition per Category (Primary Margin > Discount > Net Margin, per litre)

{_df_to_md(lubes.get('lube_margin_decomp', pd.DataFrame()))}

## Lubricants — Volume by Region x Category Matrix

{_df_to_md(lubes.get('lube_region_category', pd.DataFrame()))}

PSO is experiencing problems selling lubricants. From THIS DATA ONLY:
1. Which lube categories are shrinking and which are growing — by what volume and percentage?
2. Where is margin being eroded per litre — by category — Primary vs Net difference?
3. Which categories show BOTH volume decline AND margin compression (the double problem)?
4. Is PSO giving MORE discounts CY vs SPLY to move lubes? Which categories, how much?
5. Which regions are most problematic and which hold up?
6. What data-driven hypotheses can explain PSO's lube problem?
7. What specific actions does the data point to — be precise with which category and region?

Numbered FINDINGS then numbered RECOMMENDATIONS."""


def _prompt_lubes_geo(lubes: dict, period: str) -> str:
    return f"""Period: {period}

## Top 20 Declining Lube Cities (by absolute volume loss CY vs SPLY)

{_df_to_md(lubes.get('lube_declining_mkts', pd.DataFrame()).head(20))}

## Top 20 Growing Lube Cities (by absolute volume gain)

{_df_to_md(lubes.get('lube_growing_mkts', pd.DataFrame()).head(20))}

## Lubricants — Discount Leakage by City (highest Discount/Ltr first)

{_df_to_md(lubes.get('lube_discount_city', pd.DataFrame()).head(20))}

## Lubricants — Net Margin/Ltr Pockets by Region x Category

{_df_to_md(lubes.get('lube_margin_pockets', pd.DataFrame()))}

From THIS DATA:
1. Are lube declines concentrated in specific regions or spread nationally?
2. Do cities with highest discount/ltr also show the steepest volume declines?
3. Where are the margin pockets (high NMgn/Ltr) — are they being grown or shrinking?
4. What does the growing vs declining city pattern imply about PSO's lube distribution reach?

Numbered FINDINGS then numbered RECOMMENDATIONS."""


def _prompt_regional_gaps(analysis: dict, period: str) -> str:
    return f"""Period: {period}

## Region Performance — All Retail Segments (Volume + Margin CY vs SPLY)

{_df_to_md(analysis.get('region_performance', pd.DataFrame()))}

## Segment x Region Volume Matrix (Million Litres CY)

{_df_to_md(analysis.get('segment_region_matrix', pd.DataFrame()))}

## Underperforming Segments (both volume AND margin declined)

{_df_to_md(analysis.get('underperforming_regions', pd.DataFrame()))}

From THIS DATA:
1. Which regions show the clearest underperformance — across which segments?
2. Is underperformance product-specific (e.g. lubes in one region) or region-wide?
3. Where does volume decline but margin holds — possible mix/saturation signal?
4. What structural imbalances does the segment mix by region reveal?

Numbered FINDINGS then numbered RECOMMENDATIONS."""


def _prompt_city_concentration(analysis: dict, lubes: dict, period: str) -> str:
    return f"""Period: {period}

## All Retail — City Pareto (top 30 cities, cumulative volume and GRS share)

{_df_to_md(analysis.get('city_pareto_all', pd.DataFrame()).head(30))}

## Lubricants — Top 20 Cities x Category Matrix

{_df_to_md(lubes.get('lube_city_category', pd.DataFrame()).head(20))}

From THIS DATA:
1. How many cities = 50% of total retail volume? How many = 80%?
2. Is lube city concentration tighter or broader than fuel concentration?
3. What does the concentration pattern imply about where PSO should focus distribution effort?
4. Are the top fuel cities the same as top lube cities — or is there a geographic mismatch?

Numbered FINDINGS then numbered RECOMMENDATIONS."""


def _prompt_recommendations(period: str) -> str:
    return f"""Period: {period}

You have completed analysis across: portfolio, diesel, petrol, lubricants (category + geography),
regional gaps, and city concentration for PSO OMC data.

Produce a PRIORITIZED ACTION LIST for PSO management across three horizons:

PRIORITY 1 — Immediate (this quarter): active deterioration requiring urgent intervention
PRIORITY 2 — This half-year: structural fixes implied by geographic and category patterns
PRIORITY 3 — Strategic (next FY): portfolio, coverage, and channel decisions

For EACH action:
- State the specific problem with the number that proves it
- State the recommended action (who does what)
- State the KPI to track success and the target direction

No generic advice. Every point must be traceable to data already analyzed.
Maximum 10 actions total across all three horizons."""


# ── Public API ────────────────────────────────────────────────────────────────

def run_ai_insights(
    analysis: dict[str, pd.DataFrame],
    lubes: dict[str, pd.DataFrame],
    period: str,
    existing: dict[str, str] | None = None,
) -> dict[str, str]:
    """Generate AI narrative for each section. Returns dict[key -> text].

    Pass `existing` (loaded from a prior JSON) to skip sections already completed.
    """

    call, provider_name = _build_caller()

    if call is None:
        console.print(
            "[yellow]No AI provider configured.[/]\n"
            "Set one of: OPENAI_API_KEY, GEMINI_API_KEY, or ANTHROPIC_API_KEY in .env\n"
            "Or set AI_PROVIDER=openai|gemini|anthropic"
        )
        return {}

    results: dict[str, str] = dict(existing or {})

    def _already_done(key: str) -> bool:
        v = results.get(key, "")
        return bool(v) and not v.startswith("Rate limit") and not v.startswith("ERROR")

    sections = [
        ("exec_summary",       "Executive summary",          lambda: _prompt_exec_summary(analysis, period)),
        ("diesel_analysis",    "Diesel analysis",            lambda: _prompt_diesel(analysis, period)),
        ("petrol_analysis",    "Petrol analysis",            lambda: _prompt_petrol(analysis, period)),
        ("lubes_problem",      "Lubricants root cause",      lambda: _prompt_lubes_problem(lubes, period)),
        ("lubes_geography",    "Lubricants geography",       lambda: _prompt_lubes_geo(lubes, period)),
        ("regional_gaps",      "Regional performance gaps",  lambda: _prompt_regional_gaps(analysis, period)),
        ("city_concentration", "City concentration (Pareto)",lambda: _prompt_city_concentration(analysis, lubes, period)),
        ("recommendations",    "Prioritized action list",    lambda: _prompt_recommendations(period)),
    ]

    import time
    INTER_SECTION_DELAY = 5  # seconds between sections (pacing for 15 RPM free tier)
    pending = [(k, l, fn) for k, l, fn in sections if not _already_done(k)]
    skipped = len(sections) - len(pending)
    if skipped:
        console.print(f"  [dim]Skipping {skipped} already-complete section(s)[/]")
    for i, (key, label, prompt_fn) in enumerate(pending):
        if i > 0:
            time.sleep(INTER_SECTION_DELAY)
        console.print(f"  [dim]> {label}[/]")
        for attempt in range(4):
            try:
                results[key] = call(prompt_fn())
                break
            except Exception as e:
                err = str(e)
                if "429" in err or "RESOURCE_EXHAUSTED" in err or "quota" in err.lower():
                    wait = 60 * (attempt + 1)
                    console.print(f"    [yellow]Rate limit — waiting {wait}s (attempt {attempt+1}/4)[/]")
                    time.sleep(wait)
                else:
                    console.print(f"  [red]Error on {label}: {e}[/]")
                    results[key] = f"ERROR: {e}"
                    break
        else:
            console.print(f"  [red]Failed after 4 attempts: {label}[/]")
            results[key] = "Rate limit exceeded — re-run to populate this section."

    results["_provider"] = provider_name
    console.print(f"  [bold green]AI insights complete ({provider_name}).[/]")
    return results
