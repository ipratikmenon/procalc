#!/usr/bin/env python3
"""HyTrays — numerical distillation-column tray comparison.

Runs ``column_hydraulics.design_tray`` for each tray family in
``tray_library.ALL_TRAYS`` (Sieve, HyTrays Ripple, Valve, Dualflow,
High-Performance) on a single representative column section, then prints a
comparison table and writes:

    HyTrays/output/tray_comparison.csv
    HyTrays/output/01_capacity_and_sizing.png
    HyTrays/output/02_pressure_drop.png
    HyTrays/output/03_efficiency_and_trays.png
    HyTrays/output/04_operating_window.png
    HyTrays/output/05_downcomer_backup.png
    HyTrays/output/tray_comparison_report.md

Design basis
------------
The vapor load, liquid (reflux) load, vapor mass-flow-weighted molecular
weight and vapor density below come straight from the T801 overhead /
reflux streams in ``Hydraulics/HMB.xlsx`` (T801-OH, T801-REFLUX) -- i.e.
this is sized against the actual top-tray vapor/liquid traffic of T801's
rectifying section.

T801-REFLUX in the HMB is reported at the (cooled) reflux-drum condition
(100 F), not at the ~344 F tray temperature, and the workbook does not carry
tray-level liquid density / viscosity / surface tension / relative
volatility.  Those four properties (rho_l, mu_l, sigma, alpha) are
*engineering estimates* for a light-hydrocarbon liquid at ~344 F / 105 psia
-- replace them with simulator output for your system before using the
results for anything beyond a relative tray-type comparison.

Usage
-----
    python run_tray_comparison.py
"""
from __future__ import annotations

import csv
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from column_hydraulics import ColumnCase, TrayDesignResult, design_tray
from tray_library import ALL_TRAYS

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")

# ── design basis ──────────────────────────────────────────────────────────
CASE = ColumnCase(
    name="T801 rectifying section -- top tray (HMB.xlsx basis)",
    v_mass_lb_hr=52_448.77,   # T801-OH total mass rate           (HMB.xlsx, measured)
    l_mass_lb_hr=14_797.57,   # T801-REFLUX total mass rate       (HMB.xlsx, measured)
    rho_v=1.018,              # T801-OH actual vapor density      (HMB.xlsx, measured)
    rho_l=30.0,               # ESTIMATED liquid density @ ~344 F tray temperature
    sigma=9.0,                # ESTIMATED surface tension, light HC liquid @ ~344 F, dyn/cm
    mu_l=0.12,                # ESTIMATED liquid viscosity, light HC liquid @ ~344 F, cP
    alpha=1.40,               # ESTIMATED light/heavy-key relative volatility
    tray_spacing_in=24.0,
    f_flood=0.80,
    n_theoretical=20,         # illustrative theoretical-stage count for this section
)


def run() -> list[TrayDesignResult]:
    return [design_tray(tray, CASE) for tray in ALL_TRAYS]


# ── console table ────────────────────────────────────────────────────────
def print_table(results: list[TrayDesignResult]) -> None:
    print(f"\nHyTrays distillation-column tray comparison")
    print(f"Case: {CASE.name}")
    print(f"  V = {CASE.v_mass_lb_hr:,.0f} lb/hr  (rho_v = {CASE.rho_v:.3f} lb/ft3)")
    print(f"  L = {CASE.l_mass_lb_hr:,.0f} lb/hr  (rho_l = {CASE.rho_l:.1f} lb/ft3, "
          f"sigma = {CASE.sigma:.1f} dyn/cm, mu_l = {CASE.mu_l:.2f} cP)")
    print(f"  alpha = {CASE.alpha:.2f}, tray spacing = {CASE.tray_spacing_in:.0f} in, "
          f"design flood = {CASE.f_flood:.0%}, N_theoretical = {CASE.n_theoretical}\n")

    hdr = (f"{'Tray':<18}{'D (ft)':>8}{'u_nf (ft/s)':>12}{'dP/tray':>10}"
           f"{'dP/tray':>10}{'DC bkup':>9}{'Range %':>16}{'Turndown':>9}"
           f"{'Eff %':>8}{'N_act':>7}{'Height':>8}{'Tot dP':>9}")
    sub = (f"{'':<18}{'':>8}{'':>12}{'(in liq)':>10}{'(psi)':>10}{'(% lim)':>9}"
           f"{'(min-max)':>16}{'':>9}{'':>8}{'':>7}{'(ft)':>8}{'(psi)':>9}")
    print(hdr)
    print(sub)
    print("-" * len(hdr))
    for r in results:
        dc = "n/a" if r.dc_backup_pct is None else f"{r.dc_backup_pct:.1f}"
        rng = f"{r.min_load_pct:.0f}-{r.max_load_pct:.0f}"
        print(f"{r.tray.name:<18}{r.diameter_ft:>8.2f}{r.u_nf_ft_s:>12.2f}"
              f"{r.dp_tray_in:>10.2f}{r.dp_tray_psi:>10.4f}{dc:>9}{rng:>16}"
              f"{r.turndown_ratio:>9.2f}{r.e_tray_pct:>8.1f}{r.n_actual_trays:>7d}"
              f"{r.column_height_ft:>8.1f}{r.total_dp_psi:>9.3f}")
    print()


# ── CSV ──────────────────────────────────────────────────────────────────
def write_csv(results: list[TrayDesignResult], path: str) -> None:
    fields = [
        "tray", "diameter_ft", "a_active_ft2", "a_col_ft2",
        "flv", "csbf_ft_s", "u_nf_ft_s", "u_design_ft_s", "u_hole_ft_s",
        "dp_dry_in", "h_ow_in", "h_clear_in", "dp_tray_in", "dp_tray_psi",
        "h_dc_in", "h_dc_limit_in", "dc_backup_pct",
        "min_load_pct", "max_load_pct", "turndown_ratio",
        "e_oconnell_pct", "e_tray_pct", "n_actual_trays",
        "column_height_ft", "total_dp_psi",
    ]
    with open(path, "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(fields)
        for r in results:
            w.writerow([
                r.tray.name, f"{r.diameter_ft:.4f}", f"{r.a_active_ft2:.4f}",
                f"{r.a_col_ft2:.4f}", f"{r.flv:.5f}", f"{r.csbf_ft_s:.4f}",
                f"{r.u_nf_ft_s:.4f}", f"{r.u_design_ft_s:.4f}", f"{r.u_hole_ft_s:.4f}",
                f"{r.dp_dry_in:.4f}", f"{r.h_ow_in:.4f}", f"{r.h_clear_in:.4f}",
                f"{r.dp_tray_in:.4f}", f"{r.dp_tray_psi:.5f}",
                "" if r.h_dc_in is None else f"{r.h_dc_in:.4f}",
                "" if r.h_dc_limit_in is None else f"{r.h_dc_limit_in:.4f}",
                "" if r.dc_backup_pct is None else f"{r.dc_backup_pct:.2f}",
                f"{r.min_load_pct:.2f}", f"{r.max_load_pct:.2f}", f"{r.turndown_ratio:.3f}",
                f"{r.e_oconnell_pct:.2f}", f"{r.e_tray_pct:.2f}", r.n_actual_trays,
                f"{r.column_height_ft:.2f}", f"{r.total_dp_psi:.4f}",
            ])


# ── plots ────────────────────────────────────────────────────────────────
def make_plots(results: list[TrayDesignResult], out_dir: str) -> None:
    names = [r.tray.name for r in results]
    colors = plt.get_cmap("tab10").colors

    # 1. Capacity & sizing -------------------------------------------------
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11, 4.5))
    ax1.bar(names, [r.u_nf_ft_s for r in results], color=colors)
    ax1.set_ylabel("Flooding velocity, u_nf (ft/s)")
    ax1.set_title("Capacity (Fair flooding velocity)")
    ax1.tick_params(axis="x", rotation=20)

    ax2.bar(names, [r.diameter_ft for r in results], color=colors)
    ax2.set_ylabel("Column diameter required (ft)")
    ax2.set_title(f"Diameter @ {CASE.f_flood:.0%} of flood, fixed V & L")
    ax2.tick_params(axis="x", rotation=20)
    fig.suptitle("HyTrays tray comparison -- capacity & sizing")
    fig.tight_layout()
    fig.savefig(os.path.join(out_dir, "01_capacity_and_sizing.png"), dpi=140)
    plt.close(fig)

    # 2. Pressure drop -------------------------------------------------------
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11, 4.5))
    dry = [r.dp_dry_in for r in results]
    wet = [r.dp_tray_in - r.dp_dry_in for r in results]
    ax1.bar(names, dry, label="Dry-tray", color="#4C72B0")
    ax1.bar(names, wet, bottom=dry, label="Aerated liquid head", color="#DD8452")
    ax1.set_ylabel("Pressure drop per tray (in. liquid)")
    ax1.set_title("Per-tray pressure drop (design point)")
    ax1.legend()
    ax1.tick_params(axis="x", rotation=20)

    ax2.bar(names, [r.total_dp_psi for r in results], color=colors)
    ax2.set_ylabel("Total column pressure drop (psi)")
    ax2.set_title(f"Total {CASE.name.split(' --')[0]} dP, N_actual trays")
    ax2.tick_params(axis="x", rotation=20)
    fig.suptitle("HyTrays tray comparison -- pressure drop")
    fig.tight_layout()
    fig.savefig(os.path.join(out_dir, "02_pressure_drop.png"), dpi=140)
    plt.close(fig)

    # 3. Efficiency & tray count ---------------------------------------------
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11, 4.5))
    ax1.bar(names, [r.e_tray_pct for r in results], color=colors)
    ax1.axhline(results[0].e_oconnell_pct, color="k", ls="--", lw=1,
                 label=f"O'Connell baseline ({results[0].e_oconnell_pct:.0f}%)")
    ax1.set_ylabel("Effective tray efficiency (%)")
    ax1.set_title("Tray efficiency")
    ax1.legend()
    ax1.tick_params(axis="x", rotation=20)

    ax2b = ax2.twinx()
    bars = ax2.bar(names, [r.n_actual_trays for r in results], color="#55A868",
                    label="Actual trays")
    ax2b.plot(names, [r.column_height_ft for r in results], "o-", color="#C44E52",
              label="Column height")
    ax2.set_ylabel("Actual trays for "
                    f"{CASE.n_theoretical} theoretical stages")
    ax2b.set_ylabel("Column height (ft)")
    ax2.set_title("Actual trays & column height")
    ax2.tick_params(axis="x", rotation=20)
    fig.suptitle("HyTrays tray comparison -- efficiency & tray count")
    fig.tight_layout()
    fig.savefig(os.path.join(out_dir, "03_efficiency_and_trays.png"), dpi=140)
    plt.close(fig)

    # 4. Operating window (turndown) -----------------------------------------
    fig, ax = plt.subplots(figsize=(8, 4.5))
    for i, r in enumerate(results):
        ax.barh(i, r.max_load_pct - r.min_load_pct, left=r.min_load_pct,
                color=colors[i], height=0.5)
        ax.text(r.max_load_pct + 1, i, f"{r.turndown_ratio:.2f}:1 turndown",
                va="center", fontsize=9)
    ax.axvline(100, color="k", ls="--", lw=1, label="Design throughput (100%)")
    ax.set_yticks(range(len(results)))
    ax.set_yticklabels(names)
    ax.set_xlabel("Vapor throughput, % of design")
    ax.set_title("Stable operating window (weep limit -> flood limit)")
    ax.set_xlim(0, max(r.max_load_pct for r in results) * 1.45)
    ax.set_ylim(-0.7, len(results) - 0.3)
    ax.legend(loc="upper right")
    fig.tight_layout()
    fig.savefig(os.path.join(out_dir, "04_operating_window.png"), dpi=140)
    plt.close(fig)

    # 5. Downcomer backup ------------------------------------------------------
    fig, ax = plt.subplots(figsize=(7, 4.5))
    vals = [0 if r.dc_backup_pct is None else r.dc_backup_pct for r in results]
    bars = ax.bar(names, vals, color=colors)
    for b, r in zip(bars, results):
        label = "n/a (no DC)" if r.dc_backup_pct is None else f"{r.dc_backup_pct:.0f}%"
        ax.text(b.get_x() + b.get_width() / 2, b.get_height() + 2, label,
                ha="center", fontsize=9)
    ax.axhline(100, color="r", ls="--", lw=1, label="Downcomer flood limit (100%)")
    ax.set_ylabel("Downcomer backup, % of 50% spacing limit")
    ax.set_title("Downcomer backup")
    ax.set_ylim(0, 120)
    ax.legend()
    ax.tick_params(axis="x", rotation=20)
    fig.tight_layout()
    fig.savefig(os.path.join(out_dir, "05_downcomer_backup.png"), dpi=140)
    plt.close(fig)


# ── markdown report ──────────────────────────────────────────────────────
def write_report(results: list[TrayDesignResult], path: str) -> None:
    r_by_name = {r.tray.name: r for r in results}
    sieve = r_by_name["Sieve"]
    ripple = r_by_name["HyTrays Ripple"]

    lines = []
    a = lines.append
    a("# HyTrays tray comparison -- distillation column design\n")
    a(f"**Case:** {CASE.name}\n")
    a("## Design basis\n")
    a("| Quantity | Value | Source |")
    a("|---|---|---|")
    a(f"| Vapor traffic, V | {CASE.v_mass_lb_hr:,.0f} lb/hr | T801-OH, `Hydraulics/HMB.xlsx` (measured) |")
    a(f"| Vapor density, rho_V | {CASE.rho_v:.3f} lb/ft3 | T801-OH actual vol. rate, `Hydraulics/HMB.xlsx` (measured) |")
    a(f"| Liquid (reflux) traffic, L | {CASE.l_mass_lb_hr:,.0f} lb/hr | T801-REFLUX, `Hydraulics/HMB.xlsx` (measured) |")
    a(f"| Liquid density, rho_L | {CASE.rho_l:.1f} lb/ft3 | **estimated** for ~344 F tray liquid |")
    a(f"| Surface tension, sigma | {CASE.sigma:.1f} dyn/cm | **estimated** |")
    a(f"| Liquid viscosity, mu_L | {CASE.mu_l:.2f} cP | **estimated** |")
    a(f"| Relative volatility, alpha | {CASE.alpha:.2f} | **estimated** |")
    a(f"| Tray spacing | {CASE.tray_spacing_in:.0f} in | assumed |")
    a(f"| Design flood fraction | {CASE.f_flood:.0%} | assumed |")
    a(f"| Theoretical stages (this section) | {CASE.n_theoretical} | illustrative |")
    a("")
    a("Flow parameter FLV = (L/V)*sqrt(rho_V/rho_L) = "
      f"**{sieve.flv:.4f}** -- a low-FLV, vapor-dominated system, typical of an"
      " atmospheric/light-pressure hydrocarbon rectifying section.\n")
    a("> Replace the *estimated* rows with simulator output once available;"
      " everything else in this report is recomputed automatically by"
      " `run_tray_comparison.py`.\n")

    a("## Results\n")
    a("| Tray | Diameter (ft) | u_nf (ft/s) | dP/tray (psi) | dP/tray (in liq) |"
      " DC backup (% lim) | Operating window (% of design) | Turndown | Eff (%) |"
      " N_actual | Height (ft) | Total dP (psi) |")
    a("|---|---|---|---|---|---|---|---|---|---|---|---|")
    for r in results:
        dc = "n/a" if r.dc_backup_pct is None else f"{r.dc_backup_pct:.0f}"
        a(f"| {r.tray.name} | {r.diameter_ft:.2f} | {r.u_nf_ft_s:.2f} |"
          f" {r.dp_tray_psi:.4f} | {r.dp_tray_in:.2f} | {dc} |"
          f" {r.min_load_pct:.0f}-{r.max_load_pct:.0f} | {r.turndown_ratio:.2f}:1 |"
          f" {r.e_tray_pct:.1f} | {r.n_actual_trays} | {r.column_height_ft:.1f} |"
          f" {r.total_dp_psi:.3f} |")
    a("")

    a("## Discussion\n")

    a("**Capacity / column diameter.** At the same vapor and liquid traffic"
      " and the same 80% design flood point, the required column diameter"
      f" ranges from **{min(r.diameter_ft for r in results):.2f} ft**"
      f" ({min(results, key=lambda r: r.diameter_ft).tray.name}) to"
      f" **{max(r.diameter_ft for r in results):.2f} ft**"
      f" ({max(results, key=lambda r: r.diameter_ft).tray.name})."
      f" High-Performance and Dualflow trays need the least active area per"
      " unit of vapor handled (capacity factors of "
      f"{r_by_name['High-Performance'].tray.capacity_factor:.2f}x and"
      f" {r_by_name['Dualflow'].tray.capacity_factor:.2f}x the sieve baseline,"
      " plus little/no downcomer area), so they shrink the shell the most."
      f" The new HyTrays Ripple deck comes in at"
      f" {ripple.diameter_ft:.2f} ft -- about"
      f" {(1 - ripple.diameter_ft / sieve.diameter_ft) * 100:.0f}% smaller"
      " than a plain sieve tray -- from its"
      f" {ripple.tray.capacity_factor:.2f}x capacity factor and slightly"
      " larger active-area fraction.\n")

    a("**Pressure drop.** Per-tray pressure drop tracks hole velocity"
      " (u_hole) and the dry-tray discharge coefficient. Valve trays show"
      " the lowest *clean* dry-tray drop here because of their high"
      f" discharge coefficient (C0={r_by_name['Valve'].tray.c0:.2f}); in"
      " practice their floor (closed-valve weight, modeled as"
      f" {r_by_name['Valve'].tray.dp_dry_floor_in:.2f} in. liquid) dominates"
      " at low loads instead. Dualflow is lowest overall"
      f" ({r_by_name['Dualflow'].dp_tray_psi*1000:.1f} mpsi/tray) because it"
      " has no weir and a large open area. HyTrays Ripple"
      f" ({ripple.dp_tray_psi*1000:.1f} mpsi/tray) sits in the middle of the"
      " pack -- a modest premium over Dualflow/High-Performance for the"
      " efficiency gain described below.\n")

    a("**Downcomer backup.** FLV is low in this case (vapor-dominated), so"
      " every tray with downcomers operates well inside its 50%-of-spacing"
      " backup limit "
      f"({', '.join(f'{r.tray.name} {r.dc_backup_pct:.0f}%' for r in results if r.dc_backup_pct is not None)})."
      " Downcomer area is not the limiting consideration for this section;"
      " flooding/entrainment (Fair capacity) governs sizing instead.\n")

    a("**Turndown / operating window.** Bubble-cap-style decks and the"
      " corrugated HyTrays Ripple deck retain liquid at low vapor rates and"
      " resist weeping, giving the widest stable operating windows"
      f" ({r_by_name['Valve'].turndown_ratio:.2f}:1 for Valve,"
      f" {ripple.turndown_ratio:.2f}:1 for HyTrays Ripple)."
      " Dualflow has the narrowest window"
      f" ({r_by_name['Dualflow'].turndown_ratio:.2f}:1) because its large,"
      " unweired perforations weep heavily as soon as vapor rate drops --"
      " it needs a fairly steady load to stay efficient.\n")

    a("**Efficiency & actual trays.** The O'Connell baseline efficiency for"
      f" this system is **{sieve.e_oconnell_pct:.1f}%**. HyTrays Ripple's"
      f" extra interfacial area from the corrugated deck"
      f" ({ripple.tray.efficiency_factor:.2f}x factor) raises this to"
      f" **{ripple.e_tray_pct:.1f}%**, the highest of the five -- for"
      f" {CASE.n_theoretical} theoretical stages it needs only"
      f" **{ripple.n_actual_trays} actual trays** (column height"
      f" {ripple.column_height_ft:.0f} ft), versus"
      f" **{sieve.n_actual_trays} trays** ({sieve.column_height_ft:.0f} ft)"
      " for a plain sieve tray and"
      f" **{r_by_name['Dualflow'].n_actual_trays} trays**"
      f" ({r_by_name['Dualflow'].column_height_ft:.0f} ft) for Dualflow,"
      " whose lower contacting efficiency"
      f" ({r_by_name['Dualflow'].tray.efficiency_factor:.2f}x) offsets its"
      " low per-tray pressure drop once the whole column is added up:"
      f" total column dP is {r_by_name['Dualflow'].total_dp_psi:.2f} psi for"
      f" Dualflow vs {ripple.total_dp_psi:.2f} psi for HyTrays Ripple and"
      f" {sieve.total_dp_psi:.2f} psi for plain sieve.\n")

    a("## Summary / selection guidance\n")
    a("- **Sieve** -- baseline. Cheapest per-tray hardware, but the"
      " largest shell, the most trays for a given separation and the"
      " worst turndown of the conventional designs.")
    a("- **HyTrays Ripple** -- best overall balance in this case: smaller"
      " shell than sieve/valve, the *highest* efficiency (fewest actual"
      " trays / shortest column), a wide operating window, and total"
      " column dP close to the High-Performance tray -- without the cost"
      " or mechanical complexity of a multi-downcomer/grid design."
      " *(Parameters are placeholders -- refine once the HyTrays Ripple"
      " datasheets are added to `HyTrays/Datasheets/`.)*")
    a("- **Valve** -- same shell diameter as sieve at design, but the"
      " widest operating window of all five -- the right choice when the"
      " column must run efficiently over a wide turndown range.")
    a("- **Dualflow** -- smallest per-tray dP and a compact-ish shell, at"
      " the cost of the lowest efficiency, the most actual trays/tallest"
      " column, and a narrow operating window. Best suited to fouling or"
      " high-liquid-load services that value simplicity over turndown.")
    a("- **High-Performance** -- the smallest shell of all (lowest capex"
      " on the column shell) with efficiency close to sieve; a strong"
      " choice when plot space / shell diameter is the controlling"
      " constraint.")
    a("")
    a("---")
    a("*Generated by `HyTrays/run_tray_comparison.py`. Re-run after editing"
      " `CASE` in that file (operating conditions) or the tray parameters"
      " in `HyTrays/tray_library.py` (e.g. once HyTrays Ripple datasheets"
      " are available).*")

    with open(path, "w") as fh:
        fh.write("\n".join(lines) + "\n")


def main() -> None:
    os.makedirs(OUT_DIR, exist_ok=True)
    results = run()
    print_table(results)
    write_csv(results, os.path.join(OUT_DIR, "tray_comparison.csv"))
    make_plots(results, OUT_DIR)
    write_report(results, os.path.join(OUT_DIR, "tray_comparison_report.md"))
    print(f"Wrote CSV, plots and report to {OUT_DIR}")


if __name__ == "__main__":
    main()
