#!/usr/bin/env python3
"""HyTrays — Apex technical-paper comparison run.

Runs ``column_hydraulics.design_tray`` for ``tray_library.APEX_BENCHMARK_TRAYS``
(Sieve, Dual-Flow, Moving Valve, Fixed Valve, High-Performance MD, Ripple
(Corrugated) and HyTrays Apex) against every service case in
``service_cases.SERVICES``, adds the first-order relative installed-cost
estimate from ``cost_model.py``, and writes the numbers/figures/tables used
by ``HyTrays_Apex_Technical_Paper.md``.

This is independent of ``run_tray_comparison.py`` / ``run_service_comparison.py``
(which compare the HT-series family in ``tray_library.ALL_TRAYS``) -- it does
not change any of their outputs.

Outputs (``HyTrays/output/``)
------------------------------
    apex_paper_comparison.csv          -- full numeric results, long format
    21_apex_diameter_by_service.png     -- column diameter, all services x trays
    22_apex_total_dp_by_service.png      -- total column dP, all services x trays
    23_apex_turndown_and_efficiency.png   -- turndown & efficiency, General case
    24_apex_relative_cost_by_service.png  -- relative installed cost, all services
    25_apex_fouling_sensitivity.png        -- dry-tray dP increase from open-area loss
    apex_paper_tables.md                  -- data tables for the technical paper

Usage
-----
    python run_apex_paper.py
"""
from __future__ import annotations

import csv
import math
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

from column_hydraulics import TrayDesignResult, design_tray
from cost_model import CostResult, relative_installed_cost
from service_cases import GENERAL, SERVICES
from tray_library import APEX, APEX_BENCHMARK_TRAYS, SIEVE

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")

SHORT_NAME = {
    "General Rectification (T801 basis)": "General",
    "Fouling / Heavy-Ends Service": "Fouling",
    "Vacuum Tower Section": "Vacuum",
    "High-Pressure / High-Liquid-Load Service (C3/C4 splitter)": "High-P/High-L",
    "Foaming Service": "Foaming",
}


# ── run ──────────────────────────────────────────────────────────────────
def run_all() -> dict[str, list[TrayDesignResult]]:
    return {svc.name: [design_tray(tray, svc) for tray in APEX_BENCHMARK_TRAYS]
            for svc in SERVICES}


def cost_all(all_results: dict[str, list[TrayDesignResult]]
              ) -> dict[str, dict[str, CostResult]]:
    out: dict[str, dict[str, CostResult]] = {}
    for svc_name, results in all_results.items():
        sieve_r = next(r for r in results if r.tray.name == "Sieve")
        out[svc_name] = {r.tray.name: relative_installed_cost(r, sieve_r)
                          for r in results}
    return out


# ── CSV (long format) ───────────────────────────────────────────────────
def write_csv(all_results: dict[str, list[TrayDesignResult]],
               all_costs: dict[str, dict[str, CostResult]], path: str) -> None:
    fields = [
        "service", "tray", "diameter_ft", "a_col_ft2", "flv", "u_nf_ft_s",
        "u_design_ft_s", "f_hole_eff", "u_hole_ft_s",
        "dp_dry_in", "dp_tray_in", "dp_tray_psi",
        "h_dc_in", "h_dc_limit_in", "dc_backup_pct",
        "min_load_pct", "max_load_pct", "turndown_ratio",
        "e_oconnell_pct", "e_tray_pct", "n_actual_trays",
        "column_height_ft", "total_dp_psi",
        "shell_ratio", "tray_ratio", "relative_installed_cost",
    ]
    with open(path, "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(fields)
        for svc_name, results in all_results.items():
            for r in results:
                c = all_costs[svc_name][r.tray.name]
                w.writerow([
                    svc_name, r.tray.name, f"{r.diameter_ft:.4f}", f"{r.a_col_ft2:.4f}",
                    f"{r.flv:.5f}", f"{r.u_nf_ft_s:.4f}", f"{r.u_design_ft_s:.4f}",
                    f"{r.f_hole_eff:.4f}", f"{r.u_hole_ft_s:.4f}",
                    f"{r.dp_dry_in:.4f}", f"{r.dp_tray_in:.4f}", f"{r.dp_tray_psi:.5f}",
                    "" if r.h_dc_in is None else f"{r.h_dc_in:.4f}",
                    "" if r.h_dc_limit_in is None else f"{r.h_dc_limit_in:.4f}",
                    "" if r.dc_backup_pct is None else f"{r.dc_backup_pct:.2f}",
                    f"{r.min_load_pct:.2f}", f"{r.max_load_pct:.2f}", f"{r.turndown_ratio:.3f}",
                    f"{r.e_oconnell_pct:.2f}", f"{r.e_tray_pct:.2f}", r.n_actual_trays,
                    f"{r.column_height_ft:.2f}", f"{r.total_dp_psi:.4f}",
                    f"{c.shell_ratio:.4f}", f"{c.tray_ratio:.4f}",
                    f"{c.relative_installed_cost:.4f}",
                ])


# ── grouped bar plots (services x trays) ────────────────────────────────
def grouped_bar(ax, all_results, metric_fn, ylabel, title, none_val=0.0, annotate_none="n/a"):
    services = list(all_results.keys())
    trays = [t.name for t in APEX_BENCHMARK_TRAYS]
    colors = plt.get_cmap("tab10").colors
    x = np.arange(len(services))
    n = len(trays)
    width = 0.8 / n
    for i, tray_name in enumerate(trays):
        vals = []
        nones = []
        for svc in services:
            r = next(rr for rr in all_results[svc] if rr.tray.name == tray_name)
            v = metric_fn(r)
            nones.append(v is None)
            vals.append(none_val if v is None else v)
        offset = (i - (n - 1) / 2) * width
        bars = ax.bar(x + offset, vals, width, label=tray_name, color=colors[i])
        for b, is_none in zip(bars, nones):
            if is_none:
                ax.text(b.get_x() + b.get_width() / 2, 0.5, annotate_none,
                        ha="center", va="bottom", fontsize=7, rotation=90)
    ax.set_xticks(x)
    ax.set_xticklabels([SHORT_NAME[s] for s in services], rotation=10)
    ax.set_ylabel(ylabel)
    ax.set_title(title)
    ax.legend(fontsize=8, ncol=4, loc="upper center", bbox_to_anchor=(0.5, -0.18))


def make_plots(all_results, all_costs, out_dir: str) -> None:
    fig, ax = plt.subplots(figsize=(10, 5.5))
    grouped_bar(ax, all_results, lambda r: r.diameter_ft,
                "Column diameter (ft)",
                "Column diameter -- Apex vs. conventional trays, by service")
    fig.tight_layout()
    fig.savefig(os.path.join(out_dir, "21_apex_diameter_by_service.png"), dpi=140)
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(10, 5.5))
    grouped_bar(ax, all_results, lambda r: r.total_dp_psi,
                "Total column dP (psi)",
                "Total column pressure drop -- Apex vs. conventional trays, by service")
    fig.tight_layout()
    fig.savefig(os.path.join(out_dir, "22_apex_total_dp_by_service.png"), dpi=140)
    plt.close(fig)

    # Turndown & efficiency, General case only (single service, 7 trays)
    gen = all_results["General Rectification (T801 basis)"]
    names = [r.tray.name for r in gen]
    colors = plt.get_cmap("tab10").colors
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11, 4.5))
    ax1.bar(names, [r.turndown_ratio for r in gen], color=colors)
    for i, r in enumerate(gen):
        ax1.text(i, r.turndown_ratio + 0.3, f"{r.turndown_ratio:.2f}:1",
                  ha="center", fontsize=8)
    ax1.set_ylabel("Turndown ratio (flood limit / weep limit)")
    ax1.set_title("Operating-range turndown (General Rectification)")
    ax1.tick_params(axis="x", rotation=25)

    ax2.bar(names, [r.e_tray_pct for r in gen], color=colors)
    ax2.axhline(gen[0].e_oconnell_pct, color="k", ls="--", lw=1,
                 label=f"O'Connell baseline ({gen[0].e_oconnell_pct:.0f}%)")
    ax2.set_ylabel("Effective tray efficiency (%)")
    ax2.set_title("Tray efficiency (General Rectification)")
    ax2.legend()
    ax2.tick_params(axis="x", rotation=25)
    fig.suptitle("Apex vs. conventional trays -- flexibility & efficiency")
    fig.tight_layout()
    fig.savefig(os.path.join(out_dir, "23_apex_turndown_and_efficiency.png"), dpi=140)
    plt.close(fig)

    # Relative installed cost by service
    fig, ax = plt.subplots(figsize=(10, 5.5))
    services = list(all_results.keys())
    trays = [t.name for t in APEX_BENCHMARK_TRAYS]
    x = np.arange(len(services))
    n = len(trays)
    width = 0.8 / n
    for i, tray_name in enumerate(trays):
        vals = [all_costs[svc][tray_name].relative_installed_cost for svc in services]
        offset = (i - (n - 1) / 2) * width
        ax.bar(x + offset, vals, width, label=tray_name, color=colors[i])
    ax.axhline(1.0, color="k", ls="--", lw=1, label="Sieve = 1.00")
    ax.set_xticks(x)
    ax.set_xticklabels([SHORT_NAME[s] for s in services], rotation=10)
    ax.set_ylabel("Relative installed cost (Sieve = 1.00)")
    ax.set_title("First-order relative installed cost -- Apex vs. conventional trays")
    ax.legend(fontsize=8, ncol=4, loc="upper center", bbox_to_anchor=(0.5, -0.20))
    fig.tight_layout()
    fig.savefig(os.path.join(out_dir, "24_apex_relative_cost_by_service.png"), dpi=140)
    plt.close(fig)

    # Fouling sensitivity (flow-independent, function of tray.fouling_open_area_retention)
    fig, ax = plt.subplots(figsize=(7.5, 4.5))
    names = [t.name for t in APEX_BENCHMARK_TRAYS]
    pct = [((1.0 / t.fouling_open_area_retention) ** 2 - 1.0) * 100.0
           for t in APEX_BENCHMARK_TRAYS]
    bars = ax.bar(names, pct, color=colors)
    for b, t in zip(bars, APEX_BENCHMARK_TRAYS):
        ax.text(b.get_x() + b.get_width() / 2, b.get_height() + 1,
                f"retention={t.fouling_open_area_retention:.2f}",
                ha="center", fontsize=8)
    ax.set_ylabel("Increase in dry-tray dP (%)")
    ax.set_title("Fouling sensitivity: dry-tray dP increase\nvs. open-area retention")
    ax.tick_params(axis="x", rotation=20)
    fig.tight_layout()
    fig.savefig(os.path.join(out_dir, "25_apex_fouling_sensitivity.png"), dpi=140)
    plt.close(fig)


# ── markdown tables ──────────────────────────────────────────────────────
def results_table_md(results, costs) -> list[str]:
    lines = [
        "| Tray | Diameter (ft) | u_nf (ft/s) | dP/tray (psi) | DC backup (% lim) |"
        " Turndown | Eff (%) | N_actual | Height (ft) | Total dP (psi) |"
        " Rel. installed cost |",
        "|---|---|---|---|---|---|---|---|---|---|---|",
    ]
    for r in results:
        dc = "n/a" if r.dc_backup_pct is None else f"{r.dc_backup_pct:.0f}"
        c = costs[r.tray.name].relative_installed_cost
        lines.append(
            f"| {r.tray.name} | {r.diameter_ft:.2f} | {r.u_nf_ft_s:.2f} |"
            f" {r.dp_tray_psi:.4f} | {dc} | {r.turndown_ratio:.2f}:1 |"
            f" {r.e_tray_pct:.1f} | {r.n_actual_trays} | {r.column_height_ft:.1f} |"
            f" {r.total_dp_psi:.3f} | {c:.2f} |"
        )
    return lines


def write_tables(all_results, all_costs, path: str) -> None:
    a_lines: list[str] = []
    a = a_lines.append

    a("# Apex technical-paper data tables\n")
    a("Generated by `run_apex_paper.py` from `tray_library.APEX_BENCHMARK_TRAYS`"
      " (Sieve + 5 conventional benchmark trays + HyTrays Apex) across all 5"
      " services in `service_cases.SERVICES`. Copied into"
      " `HyTrays_Apex_Technical_Paper.md` -- re-run this script and refresh that"
      " copy after any parameter change.\n")

    # Table 1: TrayType parameters
    a("## Table 1 -- TrayType parameters\n")
    a("| Tray | f_active | f_hole | h_weir (in) | C0 | dp_dry_floor (in) |"
      " Aeration | Capacity factor | Efficiency factor | Turndown |"
      " Fouling retention | Rel. unit cost |")
    a("|---|---|---|---|---|---|---|---|---|---|---|---|")
    for t in APEX_BENCHMARK_TRAYS:
        a(f"| {t.name} | {t.f_active:.2f} | {t.f_hole:.2f} | {t.h_weir_in:.1f} |"
          f" {t.c0:.2f} | {t.dp_dry_floor_in:.2f} | {t.aeration_factor:.2f} |"
          f" {t.capacity_factor:.2f} | {t.efficiency_factor:.2f} |"
          f" {1/t.min_load_frac:.2f}:1 | {t.fouling_open_area_retention:.2f} |"
          f" {t.relative_unit_cost_factor:.2f} |")
    a("")

    # Table 2: Apex geometry breakdown at General
    gen = all_results["General Rectification (T801 basis)"]
    apex_r = next(r for r in gen if r.tray.name == "HyTrays Apex")
    sieve_r = next(r for r in gen if r.tray.name == "Sieve")
    a_col = apex_r.a_col_ft2
    a_active = apex_r.a_active_ft2
    primary_area = 0.70 * a_col
    secondary_area = 0.15 * a_col
    downcomer_area = a_col - a_active
    primary_hole = apex_r.f_hole_eff * primary_area
    secondary_hole = apex_r.f_hole_eff * secondary_area
    a("## Table 2 -- Apex physical geometry at the General Rectification design point\n")
    a(f"Design point: D = {apex_r.diameter_ft:.2f} ft "
      f"({apex_r.diameter_ft*12:.0f} in), A_col = {a_col:.2f} ft2 "
      f"({a_col*144:.0f} in2).\n")
    a("| Zone | Area fraction of A_col | Area (ft2) | Area (in2) | Open"
      " (slot/hole) area (ft2) | Open area (in2) |")
    a("|---|---|---|---|---|---|")
    a(f"| Primary (HT-02/HT-08 swirl-tube cartridges) | 0.70 | {primary_area:.2f} |"
      f" {primary_area*144:.0f} | {primary_hole:.2f} | {primary_hole*144:.0f} |")
    a(f"| Secondary (HT-01C lip-seal trickle) | 0.15 | {secondary_area:.2f} |"
      f" {secondary_area*144:.0f} | {secondary_hole:.2f} | {secondary_hole*144:.0f} |")
    a(f"| Downcomers (both sides) | {1-apex_r.tray.f_active:.2f} |"
      f" {downcomer_area:.2f} | {downcomer_area*144:.0f} | -- | -- |")
    a(f"| **Active total** | {apex_r.tray.f_active:.2f} | {a_active:.2f} |"
      f" {a_active*144:.0f} | {primary_hole+secondary_hole:.2f} |"
      f" {(primary_hole+secondary_hole)*144:.0f} |")
    a("")
    a(f"Hole velocity u_hole = {apex_r.u_hole_ft_s:.2f} ft/s; dry-tray dP ="
      f" {apex_r.dp_dry_in:.3f} in liquid; clear-liquid height ="
      f" {apex_r.h_clear_in:.3f} in (weir {apex_r.tray.h_weir_in:.1f} in + Francis"
      f" crest h_ow {apex_r.h_ow_in:.3f} in).\n")

    # Table 3: per-service results
    a("## Table 3 -- Results by service\n")
    for svc in SERVICES:
        results = all_results[svc.name]
        costs = all_costs[svc.name]
        a(f"### {svc.name}\n")
        a("\n".join(results_table_md(results, costs)))
        a("")

    # Table 4: fouling sensitivity
    a("## Table 4 -- Fouling sensitivity (flow-independent)\n")
    a("| Tray | Open-area retention | Dry-tray dP increase |")
    a("|---|---|---|")
    for t in APEX_BENCHMARK_TRAYS:
        pct = ((1.0 / t.fouling_open_area_retention) ** 2 - 1.0) * 100.0
        a(f"| {t.name} | {t.fouling_open_area_retention:.2f} | +{pct:.0f}% |")
    a("")

    # Table 5: cross-service summary (diameter, total dP, cost) vs sieve
    a("## Table 5 -- Apex vs. Sieve summary, all services\n")
    a("| Service | Diameter Sieve->Apex (ft) | Diameter change | Total dP"
      " Sieve->Apex (psi) | Total dP change | Rel. installed cost (Apex) |")
    a("|---|---|---|---|---|---|")
    for svc in SERVICES:
        results = all_results[svc.name]
        costs = all_costs[svc.name]
        sieve_r = next(r for r in results if r.tray.name == "Sieve")
        apex_r = next(r for r in results if r.tray.name == "HyTrays Apex")
        d_pct = (apex_r.diameter_ft / sieve_r.diameter_ft - 1.0) * 100.0
        dp_pct = (apex_r.total_dp_psi / sieve_r.total_dp_psi - 1.0) * 100.0
        a(f"| {SHORT_NAME[svc.name]} | {sieve_r.diameter_ft:.2f} ->"
          f" {apex_r.diameter_ft:.2f} | {d_pct:.0f}% |"
          f" {sieve_r.total_dp_psi:.3f} -> {apex_r.total_dp_psi:.3f} |"
          f" {dp_pct:.0f}% | {costs['HyTrays Apex'].relative_installed_cost:.2f} |")
    a("")

    # Table 6: cost breakdown detail, General case
    a("## Table 6 -- Relative installed-cost breakdown (General Rectification)\n")
    a("| Tray | Shell proxy ratio (D*H) | Tray-hardware proxy ratio (D2*N*k_cost) |"
      " Relative installed cost |")
    a("|---|---|---|---|")
    for r in gen:
        c = all_costs["General Rectification (T801 basis)"][r.tray.name]
        a(f"| {r.tray.name} | {c.shell_ratio:.2f} | {c.tray_ratio:.2f} |"
          f" {c.relative_installed_cost:.2f} |")
    a("")

    with open(path, "w") as fh:
        fh.write("\n".join(a_lines) + "\n")


def main() -> None:
    os.makedirs(OUT_DIR, exist_ok=True)
    all_results = run_all()
    all_costs = cost_all(all_results)
    write_csv(all_results, all_costs, os.path.join(OUT_DIR, "apex_paper_comparison.csv"))
    make_plots(all_results, all_costs, OUT_DIR)
    write_tables(all_results, all_costs, os.path.join(OUT_DIR, "apex_paper_tables.md"))
    print(f"Wrote CSV, plots and data tables to {OUT_DIR}")


if __name__ == "__main__":
    main()
