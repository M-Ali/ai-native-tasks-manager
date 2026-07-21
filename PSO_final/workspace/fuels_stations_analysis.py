"""
PSO Retail Fuels — Station & Volume Analysis. Mirrors workspace/lubes_stations_analysis.py.
"""
import sys
sys.path.insert(0, 'src')
sys.stdout.reconfigure(encoding='utf-8')

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from pso import ingest
from _pso_common import INPUT_PATH, out_path

# ── load ──────────────────────────────────────────────────────────────────────
df, _ = ingest.load(INPUT_PATH)
retail = df[df['IsRetail'] & ~df['IsInternational']].copy()
fuels  = retail[retail['FuelSegment'].isin(['Diesel', 'Petrol', 'Other Fuels', 'LPG'])].copy()

# ── 1. Volume & stations per FuelSegment ──────────────────────────────────────
cat = (fuels.groupby('FuelSegment')
       .agg(vol_cy=('SalesLtr_CY', 'sum'),
            vol_sply=('SalesLtr_SPLY', 'sum'),
            stns=('Customer Number', 'nunique'))
       .assign(vol_chg=lambda d: (d.vol_cy - d.vol_sply) / d.vol_sply.abs().replace(0, float('nan')) * 100,
               vol_sh=lambda d: d.vol_cy / d.vol_cy.sum() * 100,
               vol_per_stn=lambda d: d.vol_cy / d.stns)
       .sort_values('vol_cy', ascending=False))

print("=== VOLUME & STATIONS BY FUEL CATEGORY ===")
print(f"{'Category':<20}  {'Vol CY (KL)':>11}  {'Vol SPLY (KL)':>13}  {'vs SPLY%':>8}  {'Mix%':>6}  {'Stns':>5}  {'KL/Stn':>8}")
print("-" * 85)
for cat_name, r in cat.iterrows():
    chg = f"{r.vol_chg:+.1f}%"
    print(f"{cat_name:<20}  {r.vol_cy/1000:>11,.1f}  {r.vol_sply/1000:>11,.1f}  {chg:>7}  {r.vol_sh:>5.1f}%  {int(r.stns):>5}  {r.vol_per_stn/1000:>8.2f}")

print()

# ── 2. Stations per top 15 cities ─────────────────────────────────────────────
city = (fuels.groupby('CityNorm')
        .agg(vol_cy=('SalesLtr_CY', 'sum'),
             vol_sply=('SalesLtr_SPLY', 'sum'),
             stns=('Customer Number', 'nunique'))
        .assign(vol_chg=lambda d: (d.vol_cy - d.vol_sply) / d.vol_sply.abs().replace(0, float('nan')) * 100,
                vol_per_stn=lambda d: d.vol_cy / d.stns)
        .sort_values('vol_cy', ascending=False)
        .head(15))

print("=== TOP 15 CITIES — STATIONS & VOLUME ===")
print(f"{'#':<3}  {'City':<25}  {'Stns':>5}  {'Vol CY (KL)':>11}  {'Vol SPLY (KL)':>13}  {'vs SPLY%':>8}  {'KL/Stn':>8}")
print("-" * 80)
for i, (city_name, r) in enumerate(city.iterrows(), 1):
    chg = f"{r.vol_chg:+.1f}%"
    print(f"{i:<3}  {city_name:<25}  {int(r.stns):>5}  {r.vol_cy/1000:>11,.1f}  {r.vol_sply/1000:>11,.1f}  {chg:>7}  {r.vol_per_stn/1000:>8.2f}")

# ── CHARTS ────────────────────────────────────────────────────────────────────
fig, axes = plt.subplots(1, 2, figsize=(14, 5.5))
fig.patch.set_facecolor('white')
fig.suptitle('PSO Retail Fuels — Station & Volume Analysis', fontsize=13,
             fontweight='bold', color='#00479D', y=1.01)

BLUE, GREEN, ORANGE, RED, GREY = '#00479D', '#008C4A', '#E46C0A', '#C00000', '#BFBFBF'

# ── Chart 1: Category — dual axis (volume bars + stations line) ───────────────
ax1 = axes[0]
ax1r = ax1.twinx()

cat_labels = cat.index.tolist()
x = np.arange(len(cat))
bw = 0.55
bar_colours = [BLUE, GREEN, ORANGE, RED][:len(cat)]

bars = ax1.bar(x, cat['vol_cy']/1000, bw, color=bar_colours, alpha=0.85, zorder=3, label='Vol CY (KL)')
ax1.bar(x, cat['vol_sply']/1000, bw, color='none', edgecolor=GREY,
        linewidth=1.5, linestyle='--', zorder=2, label='Vol SPLY (KL)')

ax1r.plot(x, cat['stns'], 'o--', color='#333333', linewidth=1.5,
          markersize=7, markerfacecolor='white', markeredgewidth=2, zorder=4, label='Stations')
for xi, (_, r) in zip(x, cat.iterrows()):
    ax1r.annotate(f"{int(r.stns):,}", xy=(xi, r.stns),
                  xytext=(0, 9), textcoords='offset points', ha='center',
                  fontsize=8.5, fontweight='bold', color='#333333')

for bar, (_, r) in zip(bars, cat.iterrows()):
    c = GREEN if r.vol_chg >= 0 else RED
    ax1.annotate(f"{r.vol_chg:+.1f}%",
                 xy=(bar.get_x() + bar.get_width()/2, bar.get_height()),
                 xytext=(0, 5), textcoords='offset points', ha='center',
                 fontsize=8, color=c, fontweight='bold')

ax1.set_xticks(x); ax1.set_xticklabels(cat_labels, fontsize=9)
ax1.set_ylabel('Volume (KL)', fontsize=9)
ax1r.set_ylabel('No. of Stations', fontsize=9, color='#333333')
ax1r.tick_params(axis='y', labelcolor='#333333')
ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f'{v:,.0f}'))
ax1.set_title('Volume & Stations by Fuel Category', fontsize=10, fontweight='bold', color=BLUE, pad=8)
ax1.set_axisbelow(True); ax1.yaxis.grid(True, linestyle='--', alpha=0.4)
ax1.spines[['top', 'right']].set_visible(False)

lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax1r.get_legend_handles_labels()
ax1.legend(lines1 + lines2, labels1 + labels2, fontsize=8, loc='upper right')

# ── Chart 2: Top 15 cities — stations (bar) + volume per station (line) ───────
ax2 = axes[1]
ax2r = ax2.twinx()

city_names = [c[:18] for c in city.index.tolist()]
city_stns = city['stns'].values
city_vol_ps = city['vol_per_stn'].values / 1000
city_chg = city['vol_chg'].values

bar_c2 = [GREEN if c >= 0 else RED for c in city_chg]
bars2 = ax2.bar(range(len(city)), city_stns, 0.6, color=bar_c2, alpha=0.82, zorder=3)

ax2r.plot(range(len(city)), city_vol_ps, 's-', color=BLUE, linewidth=1.8,
          markersize=6, zorder=4, label='KL / Station')
for xi, vps in enumerate(city_vol_ps):
    ax2r.annotate(f'{vps:.1f}', xy=(xi, vps), xytext=(0, 7), textcoords='offset points',
                  ha='center', fontsize=7.5, color=BLUE)

for bar, n in zip(bars2, city_stns):
    ax2.annotate(str(int(n)), xy=(bar.get_x() + bar.get_width()/2, bar.get_height()),
                 xytext=(0, 3), textcoords='offset points', ha='center',
                 fontsize=7.5, fontweight='bold', color='#333333')

ax2.set_xticks(range(len(city)))
ax2.set_xticklabels(city_names, rotation=40, ha='right', fontsize=8)
ax2.set_ylabel('No. of Stations', fontsize=9)
ax2r.set_ylabel('KL per Station', fontsize=9, color=BLUE)
ax2r.tick_params(axis='y', labelcolor=BLUE)
ax2.set_title('Top 15 Cities — Active Stations & Productivity', fontsize=10, fontweight='bold', color=BLUE, pad=8)
ax2.set_axisbelow(True); ax2.yaxis.grid(True, linestyle='--', alpha=0.4)
ax2.spines[['top', 'right']].set_visible(False)

green_patch = plt.matplotlib.patches.Patch(color=GREEN, label='Growing city')
red_patch = plt.matplotlib.patches.Patch(color=RED, label='Declining city')
blue_line = plt.Line2D([0], [0], color=BLUE, marker='s', linewidth=1.5, label='KL / Station')
ax2.legend(handles=[green_patch, red_patch, blue_line], fontsize=7.5, loc='upper right')

plt.tight_layout()
out = out_path('PSO_Fuels_Stations_Analysis', 'png', df)
fig.savefig(out, dpi=160, bbox_inches='tight', facecolor='white')
plt.close()
print(f"\nChart saved: {out}")
