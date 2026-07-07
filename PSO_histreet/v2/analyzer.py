"""
Pricing analysis engine.
Builds a competitive price matrix and surfaces:
  - PSO's price position vs each competitor per grade/pack
  - Tier gap analysis (is the ladder wide enough?)
  - Over/under-priced SKUs (opportunities)
  - Pack-price premium consistency
"""
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import seaborn as sns
from pathlib import Path

from db.store import latest_run, all_runs_summary

OUTPUT = Path("output/charts")
OUTPUT.mkdir(parents=True, exist_ok=True)

PSO_GREEN = "#00703C"
PSO_LIGHT = "#E8F5EE"
COMPETITOR_COLOR = "#2D72B8"
RED = "#C0392B"
AMBER = "#E67E22"

# Tier ordering for display
TIER_ORDER = ["super_premium", "premium", "mainstream", "economy"]
TIER_LABELS = {
    "super_premium": "Super Premium",
    "premium": "Premium",
    "mainstream": "Mainstream",
    "economy": "Economy",
}

PSO_TIER_MAP = {
    "PSO Carient Ultra": "super_premium",
    "PSO Carient FS": "premium",
    "PSO Carient Plus": "mainstream",
    "PSO Carient SPRO": "economy",
}


def load_data() -> pd.DataFrame:
    rows = latest_run()
    if not rows:
        print("No scraped data found. Run: uv run python main.py scrape")
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    df = df[df["price_per_litre"].notna() & df["grade_detected"].notna()]
    df["is_pso"] = df["brand_detected"].str.contains("PSO|Carient", na=False)
    return df


def run_all(config: dict):
    df = load_data()
    if df.empty:
        return

    print("\n=== Pricing Analysis ===")
    chart_price_matrix(df, config)
    chart_tier_gaps(df, config)
    chart_pack_premium(df, config)
    chart_price_index(df, config)
    chart_grade_heatmap(df, config)
    print_gap_table(df, config)
    print("\nAll charts saved to output/charts/")


# ── Chart 1: Competitive Price Matrix ─────────────────────────────

def chart_price_matrix(df: pd.DataFrame, config: dict):
    """Box plot: price/L distribution per brand, per grade — PCMO and HDEO side by side."""
    for oil_type, label in [("pcmo", "Petrol Engines (PCMO)"), ("hdeo", "Diesel Engines (HDEO)")]:
        sub = df[df["oil_type"] == oil_type].copy()
        if sub.empty:
            continue

        grades = config["grades"].get(oil_type, [])
        sub = sub[sub["grade_detected"].isin(grades)]
        if sub.empty:
            continue

        brands = sorted(sub["brand_detected"].dropna().unique())
        n_grades = len(grades)
        n_brands = len(brands)

        fig, axes = plt.subplots(1, n_grades, figsize=(max(4 * n_grades, 12), 6), sharey=True)
        if n_grades == 1:
            axes = [axes]

        fig.suptitle(f"Competitive Price per Litre (PKR) — {label}", fontsize=14, fontweight="bold")

        for ax, grade in zip(axes, grades):
            gdf = sub[sub["grade_detected"] == grade]
            brand_prices = []
            brand_labels = []
            colors = []
            for b in brands:
                bdf = gdf[gdf["brand_detected"] == b]["price_per_litre"].dropna()
                if len(bdf) > 0:
                    brand_prices.append(bdf.values)
                    brand_labels.append(b)
                    colors.append(PSO_GREEN if "PSO" in b or "Carient" in b else COMPETITOR_COLOR)

            if not brand_prices:
                ax.set_title(grade)
                ax.text(0.5, 0.5, "No data", ha="center", va="center", transform=ax.transAxes)
                continue

            bp = ax.boxplot(brand_prices, patch_artist=True, notch=False)
            for patch, color in zip(bp["boxes"], colors):
                patch.set_facecolor(color)
                patch.set_alpha(0.7)

            ax.set_title(grade, fontweight="bold")
            ax.set_xticks(range(1, len(brand_labels) + 1))
            ax.set_xticklabels(brand_labels, rotation=45, ha="right", fontsize=8)
            ax.grid(axis="y", alpha=0.3)

        axes[0].set_ylabel("Price per Litre (PKR)")
        pso_patch = mpatches.Patch(color=PSO_GREEN, alpha=0.7, label="PSO")
        comp_patch = mpatches.Patch(color=COMPETITOR_COLOR, alpha=0.7, label="Competitor")
        fig.legend(handles=[pso_patch, comp_patch], loc="upper right", framealpha=0.9)
        plt.tight_layout()
        fname = OUTPUT / f"01_price_matrix_{oil_type}.png"
        plt.savefig(fname, dpi=150, bbox_inches="tight")
        plt.close()
        print(f"  Saved {fname.name}")


# ── Chart 2: Tier Gap Analysis ─────────────────────────────────────

def chart_tier_gaps(df: pd.DataFrame, config: dict):
    """Bar chart showing median price/L by tier — are PSO's inter-tier gaps wide enough?"""
    # Map brands to tiers using config
    tier_map = {}
    for comp in config.get("competitors", []):
        for variant, tier in comp.get("tier_map", {}).items():
            tier_map[variant.lower()] = tier
    for brand, tier in PSO_TIER_MAP.items():
        tier_map[brand.lower()] = tier

    def assign_tier(brand):
        if pd.isna(brand):
            return None
        bl = brand.lower()
        for key, tier in tier_map.items():
            if key in bl:
                return tier
        return None

    df = df.copy()
    df["tier"] = df["brand_detected"].apply(assign_tier)
    df_t = df[df["tier"].notna()].copy()

    tier_summary = (
        df_t.groupby(["tier", "is_pso"])["price_per_litre"]
        .median()
        .reset_index()
        .rename(columns={"price_per_litre": "median_ppl"})
    )

    fig, ax = plt.subplots(figsize=(10, 6))
    x = np.arange(len(TIER_ORDER))
    width = 0.35

    pso_vals = []
    comp_vals = []
    for tier in TIER_ORDER:
        pso_row = tier_summary[(tier_summary["tier"] == tier) & (tier_summary["is_pso"])]
        comp_row = tier_summary[(tier_summary["tier"] == tier) & (~tier_summary["is_pso"])]
        pso_vals.append(pso_row["median_ppl"].values[0] if not pso_row.empty else 0)
        comp_vals.append(comp_row["median_ppl"].values[0] if not comp_row.empty else 0)

    bars_pso = ax.bar(x - width / 2, pso_vals, width, label="PSO Carient", color=PSO_GREEN, alpha=0.85)
    bars_comp = ax.bar(x + width / 2, comp_vals, width, label="Competitors (median)", color=COMPETITOR_COLOR, alpha=0.85)

    for bar, val in zip(bars_pso, pso_vals):
        if val > 0:
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 10,
                    f"Rs {val:,.0f}", ha="center", va="bottom", fontsize=8, color=PSO_GREEN, fontweight="bold")
    for bar, val in zip(bars_comp, comp_vals):
        if val > 0:
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 10,
                    f"Rs {val:,.0f}", ha="center", va="bottom", fontsize=8, color=COMPETITOR_COLOR, fontweight="bold")

    ax.set_title("Tier Price Positioning: PSO vs Market (Median Price/Litre)", fontsize=13, fontweight="bold")
    ax.set_ylabel("Median Price per Litre (PKR)")
    ax.set_xticks(x)
    ax.set_xticklabels([TIER_LABELS[t] for t in TIER_ORDER])
    ax.legend()
    ax.grid(axis="y", alpha=0.3)
    plt.tight_layout()
    fname = OUTPUT / "02_tier_gaps.png"
    plt.savefig(fname, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"  Saved {fname.name}")


# ── Chart 3: Pack Size Premium Consistency ─────────────────────────

def chart_pack_premium(df: pd.DataFrame, config: dict):
    """Line chart: Price/L vs Pack Size — should decline as pack grows."""
    pack_sizes = config.get("pack_sizes", [1, 3, 4, 5, 10, 20])
    df_p = df[df["pack_size_l"].isin(pack_sizes) & df["price_per_litre"].notna()].copy()

    fig, axes = plt.subplots(1, 2, figsize=(14, 6))
    fig.suptitle("Pack Size Premium Curve: Price per Litre vs Pack Size", fontsize=13, fontweight="bold")

    for ax, (oil_type, label) in zip(axes, [("pcmo", "Petrol (PCMO)"), ("hdeo", "Diesel (HDEO)")]):
        sub = df_p[df_p["oil_type"] == oil_type]
        brands = sub["brand_detected"].dropna().unique()
        for brand in brands:
            bdf = (
                sub[sub["brand_detected"] == brand]
                .groupby("pack_size_l")["price_per_litre"]
                .median()
                .reset_index()
                .sort_values("pack_size_l")
            )
            if len(bdf) < 2:
                continue
            color = PSO_GREEN if "PSO" in brand or "Carient" in brand else None
            lw = 2.5 if color else 1.2
            ax.plot(bdf["pack_size_l"], bdf["price_per_litre"],
                    marker="o", label=brand, color=color, linewidth=lw,
                    alpha=0.9 if color else 0.6)

        ax.set_title(label)
        ax.set_xlabel("Pack Size (Litres)")
        ax.set_ylabel("Median Price per Litre (PKR)")
        ax.legend(fontsize=7, loc="upper right")
        ax.grid(alpha=0.3)

    plt.tight_layout()
    fname = OUTPUT / "03_pack_premium_curve.png"
    plt.savefig(fname, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"  Saved {fname.name}")


# ── Chart 4: Price Index (PSO = 100) ──────────────────────────────

def chart_price_index(df: pd.DataFrame, config: dict):
    """Spider/bar: Index PSO price as 100, show where competitors sit per grade."""
    pso_df = df[df["is_pso"] & df["grade_detected"].notna() & df["price_per_litre"].notna()]
    pso_median = pso_df.groupby("grade_detected")["price_per_litre"].median()

    all_grades = sorted(df["grade_detected"].dropna().unique())
    brands = [b for b in df["brand_detected"].dropna().unique() if not ("PSO" in b or "Carient" in b)]

    matrix = {}
    for brand in brands:
        bdf = df[df["brand_detected"] == brand]
        indices = []
        for grade in all_grades:
            gdf = bdf[bdf["grade_detected"] == grade]["price_per_litre"]
            if gdf.empty or grade not in pso_median:
                indices.append(np.nan)
            else:
                indices.append(round(gdf.median() / pso_median[grade] * 100, 1))
        matrix[brand] = indices

    plot_df = pd.DataFrame(matrix, index=all_grades)
    plot_df = plot_df.dropna(how="all")

    if plot_df.empty:
        return

    fig, ax = plt.subplots(figsize=(max(10, len(all_grades) * 1.5), 6))
    x = np.arange(len(plot_df))
    n_brands = len(plot_df.columns)
    width = 0.8 / n_brands

    cmap = plt.colormaps["tab10"].resampled(n_brands)
    for i, brand in enumerate(plot_df.columns):
        vals = plot_df[brand].values
        offset = (i - n_brands / 2 + 0.5) * width
        bars = ax.bar(x + offset, vals, width, label=brand, color=cmap(i), alpha=0.8)

    ax.axhline(100, color=PSO_GREEN, linewidth=2, linestyle="--", label="PSO = 100")
    ax.fill_between([-0.5, len(plot_df) - 0.5], 90, 110, alpha=0.1, color=PSO_GREEN, label="+/-10% band")

    ax.set_title("Price Index vs PSO (PSO = 100) by Grade", fontsize=13, fontweight="bold")
    ax.set_ylabel("Price Index (PSO = 100)")
    ax.set_xlabel("Grade")
    ax.set_xticks(x)
    ax.set_xticklabels(plot_df.index, rotation=45)
    ax.legend(fontsize=8, loc="upper right", ncol=2)
    ax.grid(axis="y", alpha=0.3)
    plt.tight_layout()
    fname = OUTPUT / "04_price_index.png"
    plt.savefig(fname, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"  Saved {fname.name}")


# ── Chart 5: Price Heatmap (Brand x Grade) ────────────────────────

def chart_grade_heatmap(df: pd.DataFrame, config: dict):
    """Heatmap: median price/L for every brand × grade combo."""
    all_grades = sorted(df["grade_detected"].dropna().unique())
    brands = sorted(df["brand_detected"].dropna().unique())

    matrix = pd.DataFrame(index=brands, columns=all_grades, dtype=float)
    for brand in brands:
        for grade in all_grades:
            vals = df[(df["brand_detected"] == brand) & (df["grade_detected"] == grade)]["price_per_litre"]
            matrix.loc[brand, grade] = vals.median() if not vals.empty else np.nan

    matrix = matrix.dropna(how="all").dropna(axis=1, how="all")

    if matrix.empty:
        return

    fig, ax = plt.subplots(figsize=(max(8, len(matrix.columns) * 1.2), max(5, len(matrix) * 0.6)))
    sns.heatmap(
        matrix.astype(float),
        annot=True, fmt=".0f", cmap="YlOrRd",
        linewidths=0.5, ax=ax,
        cbar_kws={"label": "Median Price/Litre (PKR)"}
    )
    ax.set_title("Price per Litre Heatmap: Brand × Grade", fontsize=13, fontweight="bold")
    ax.set_xlabel("Grade")
    ax.set_ylabel("Brand")
    plt.tight_layout()
    fname = OUTPUT / "05_price_heatmap.png"
    plt.savefig(fname, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"  Saved {fname.name}")


# ── Gap Table (console) ────────────────────────────────────────────

def print_gap_table(df: pd.DataFrame, config: dict):
    """Print a console table showing PSO's gap vs cheapest/median/most expensive competitor per grade."""
    pso = df[df["is_pso"] & df["price_per_litre"].notna()]
    comps = df[~df["is_pso"] & df["price_per_litre"].notna()]

    grades = sorted(df["grade_detected"].dropna().unique())

    print("\n" + "="*80)
    print(f"{'Grade':<10} {'PSO/L':>8} {'Mkt Min':>9} {'Mkt Med':>9} {'Mkt Max':>9} {'vs Med':>8} {'Signal'}")
    print("="*80)

    for grade in grades:
        pso_vals = pso[pso["grade_detected"] == grade]["price_per_litre"]
        comp_vals = comps[comps["grade_detected"] == grade]["price_per_litre"]

        if pso_vals.empty or comp_vals.empty:
            continue

        pso_med = pso_vals.median()
        mkt_min = comp_vals.min()
        mkt_med = comp_vals.median()
        mkt_max = comp_vals.max()
        gap_pct = (pso_med - mkt_med) / mkt_med * 100

        if gap_pct > 10:
            signal = "OVERPRICED vs market"
        elif gap_pct < -10:
            signal = "UNDERPRICED - margin left"
        else:
            signal = "At market"

        print(f"{grade:<10} {pso_med:>8,.0f} {mkt_min:>9,.0f} {mkt_med:>9,.0f} {mkt_max:>9,.0f} "
              f"{gap_pct:>+7.1f}% {signal}")

    print("="*80)
