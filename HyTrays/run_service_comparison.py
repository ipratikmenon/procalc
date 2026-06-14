#!/usr/bin/env python3
"""HyTrays — tray selection study across services.

Runs every tray family in ``tray_library.ALL_TRAYS`` (the conventional Sieve
baseline plus the six HT-series contacting decks: HT-01A Hinge, HT-01B
Spring, HT-01C LipSeal, HT-02 CVS, HT-03 GRADEX and HT-05 PULSAR) against
every service case in ``service_cases.SERVICES`` (General Rectification,
Fouling/Heavy-Ends, Vacuum, High-Pressure/High-Liquid-Load, Foaming) and
assembles a "paper"-style write-up comparing the tray types service by
service.

Outputs (``HyTrays/output/``)
------------------------------
    service_comparison.csv          -- full numeric results, long format
    11_diameter_by_service.png       -- column diameter, all services x trays
    12_total_dp_by_service.png        -- total column dP, all services x trays
    13_downcomer_backup_by_service.png -- downcomer backup, all services x trays
    14_turndown_by_service.png        -- turndown ratio, all services x trays
    15_fouling_sensitivity.png        -- dry-tray dP increase from open-area loss
    tray_selection_study.md          -- the write-up (tables + discussion)

Usage
-----
    python run_service_comparison.py
"""
from __future__ import annotations

import csv
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

from column_hydraulics import TrayDesignResult, design_tray
from service_cases import SERVICES
from tray_library import ALL_TRAYS

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")

SHORT_NAME = {
    "General Rectification (T801 basis)": "General",
    "Fouling / Heavy-Ends Service": "Fouling",
    "Vacuum Tower Section": "Vacuum",
    "High-Pressure / High-Liquid-Load Service (C3/C4 splitter)": "High-P/High-L",
    "Foaming Service": "Foaming",
}


def run_all() -> dict[str, list[TrayDesignResult]]:
    return {svc.name: [design_tray(tray, svc) for tray in ALL_TRAYS] for svc in SERVICES}


# ── console table ────────────────────────────────────────────────────────
def print_service_table(svc_name: str, results: list[TrayDesignResult]) -> None:
    print(f"\n=== {svc_name} ===")
    hdr = (f"{'Tray':<18}{'D (ft)':>8}{'u_nf':>8}{'dP/tray':>10}{'DC bkup':>9}"
           f"{'Range %':>16}{'Turndown':>9}{'Eff %':>8}{'N_act':>7}{'Tot dP':>9}")
    print(hdr)
    print("-" * len(hdr))
    for r in results:
        dc = "n/a" if r.dc_backup_pct is None else f"{r.dc_backup_pct:.1f}"
        rng = f"{r.min_load_pct:.0f}-{r.max_load_pct:.0f}"
        print(f"{r.tray.name:<18}{r.diameter_ft:>8.2f}{r.u_nf_ft_s:>8.2f}"
              f"{r.dp_tray_psi:>10.4f}{dc:>9}{rng:>16}{r.turndown_ratio:>9.2f}"
              f"{r.e_tray_pct:>8.1f}{r.n_actual_trays:>7d}{r.total_dp_psi:>9.3f}")


# ── CSV (long format) ───────────────────────────────────────────────────
def write_csv(all_results: dict[str, list[TrayDesignResult]], path: str) -> None:
    fields = [
        "service", "tray", "diameter_ft", "flv", "u_nf_ft_s", "u_design_ft_s",
        "f_hole_eff", "dp_dry_in", "dp_tray_in", "dp_tray_psi",
        "h_dc_in", "h_dc_limit_in", "dc_backup_pct",
        "min_load_pct", "max_load_pct", "turndown_ratio",
        "e_oconnell_pct", "e_tray_pct", "n_actual_trays",
        "column_height_ft", "total_dp_psi",
    ]
    with open(path, "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(fields)
        for svc_name, results in all_results.items():
            for r in results:
                w.writerow([
                    svc_name, r.tray.name, f"{r.diameter_ft:.4f}", f"{r.flv:.5f}",
                    f"{r.u_nf_ft_s:.4f}", f"{r.u_design_ft_s:.4f}", f"{r.f_hole_eff:.4f}",
                    f"{r.dp_dry_in:.4f}", f"{r.dp_tray_in:.4f}", f"{r.dp_tray_psi:.5f}",
                    "" if r.h_dc_in is None else f"{r.h_dc_in:.4f}",
                    "" if r.h_dc_limit_in is None else f"{r.h_dc_limit_in:.4f}",
                    "" if r.dc_backup_pct is None else f"{r.dc_backup_pct:.2f}",
                    f"{r.min_load_pct:.2f}", f"{r.max_load_pct:.2f}", f"{r.turndown_ratio:.3f}",
                    f"{r.e_oconnell_pct:.2f}", f"{r.e_tray_pct:.2f}", r.n_actual_trays,
                    f"{r.column_height_ft:.2f}", f"{r.total_dp_psi:.4f}",
                ])


# ── grouped bar plots (services x trays) ────────────────────────────────
def grouped_bar(ax, all_results, metric_fn, ylabel, title, none_val=0.0, annotate_none="n/a"):
    services = list(all_results.keys())
    trays = [t.name for t in ALL_TRAYS]
    colors = plt.get_cmap("tab10").colors
    x = np.arange(len(services))
    n = len(trays)
    width = 0.8 / n
    for i, tray_name in enumerate(trays):
        vals = []
        nones = []
        for j, svc in enumerate(services):
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
    ax.legend(fontsize=8, ncol=5, loc="upper center", bbox_to_anchor=(0.5, -0.18))


def make_plots(all_results: dict[str, list[TrayDesignResult]], out_dir: str) -> None:
    fig, ax = plt.subplots(figsize=(10, 5.5))
    grouped_bar(ax, all_results, lambda r: r.diameter_ft,
                "Column diameter (ft)", "Column diameter by service & tray type")
    fig.tight_layout()
    fig.savefig(os.path.join(out_dir, "11_diameter_by_service.png"), dpi=140)
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(10, 5.5))
    grouped_bar(ax, all_results, lambda r: r.total_dp_psi,
                "Total column dP (psi)", "Total column pressure drop by service & tray type")
    fig.tight_layout()
    fig.savefig(os.path.join(out_dir, "12_total_dp_by_service.png"), dpi=140)
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(10, 5.5))
    grouped_bar(ax, all_results, lambda r: r.dc_backup_pct,
                "Downcomer backup (% of 50% limit)",
                "Downcomer backup by service & tray type")
    ax.axhline(100, color="r", ls="--", lw=1)
    fig.tight_layout()
    fig.savefig(os.path.join(out_dir, "13_downcomer_backup_by_service.png"), dpi=140)
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(10, 5.5))
    grouped_bar(ax, all_results, lambda r: r.turndown_ratio,
                "Turndown ratio (flood limit / weep limit)",
                "Operating-range turndown by service & tray type")
    fig.tight_layout()
    fig.savefig(os.path.join(out_dir, "14_turndown_by_service.png"), dpi=140)
    plt.close(fig)

    # Fouling sensitivity -- purely a function of tray.fouling_open_area_retention
    fig, ax = plt.subplots(figsize=(7, 4.5))
    names = [t.name for t in ALL_TRAYS]
    pct = [((1.0 / t.fouling_open_area_retention) ** 2 - 1.0) * 100.0 for t in ALL_TRAYS]
    colors = plt.get_cmap("tab10").colors
    bars = ax.bar(names, pct, color=colors)
    for b, t in zip(bars, ALL_TRAYS):
        ax.text(b.get_x() + b.get_width() / 2, b.get_height() + 1,
                f"retention={t.fouling_open_area_retention:.2f}",
                ha="center", fontsize=8)
    ax.set_ylabel("Increase in dry-tray dP (%)")
    ax.set_title("Fouling sensitivity: dry-tray dP increase\nfor each tray's assumed open-area retention")
    ax.tick_params(axis="x", rotation=15)
    fig.tight_layout()
    fig.savefig(os.path.join(out_dir, "15_fouling_sensitivity.png"), dpi=140)
    plt.close(fig)


# ── ranking / selection matrix ──────────────────────────────────────────
RANKING_SPECS = [
    ("General Rectification (T801 basis)",
     lambda r: r.total_dp_psi, True,
     "Lower total column dP across the N_actual trays needed for the target "
     "separation (psi)."),
    ("Fouling / Heavy-Ends Service",
     lambda r: r.tray.fouling_open_area_retention, False,
     "Higher fouling open-area retention -> smaller dry-tray dP increase as "
     "deposits build up (see fouling sensitivity figure)."),
    ("Vacuum Tower Section",
     lambda r: r.dp_tray_psi, True,
     "Lower per-tray dP -- directly limits flash-zone temperature rise in "
     "vacuum service."),
    ("High-Pressure / High-Liquid-Load Service (C3/C4 splitter)",
     lambda r: r.dc_backup_pct, True,
     "Lower downcomer backup (% of the 50%-of-spacing limit) at high FLV. "
     "All HT-series decks here are weired; the HT-04 DCX active-downcomer "
     "module (not rated by this engine) is the family's dedicated answer to "
     "high weir loading."),
    ("Foaming Service",
     lambda r: r.tray.aeration_factor, True,
     "Lower aeration factor -> less intense froth generation -> generally "
     "more foam-tolerant."),
]


def rank_service(results, key_func, lower_is_better):
    scored = [(r.tray.name, key_func(r)) for r in results]
    have = [(n, v) for n, v in scored if v is not None]
    none_ = [(n, v) for n, v in scored if v is None]
    have.sort(key=lambda nv: nv[1], reverse=not lower_is_better)
    ordered = have + none_
    return [(i + 1, name, val) for i, (name, val) in enumerate(ordered)]


def build_selection_matrix(all_results):
    matrix = {t.name: {} for t in ALL_TRAYS}
    for svc_name, key_func, lower_is_better, _ in RANKING_SPECS:
        ranked = rank_service(all_results[svc_name], key_func, lower_is_better)
        for rank, tray_name, val in ranked:
            matrix[tray_name][svc_name] = (rank, val)
    return matrix


# ── markdown report ──────────────────────────────────────────────────────
def results_table_md(results: list[TrayDesignResult]) -> list[str]:
    lines = [
        "| Tray | Diameter (ft) | u_nf (ft/s) | dP/tray (psi) | DC backup (% lim) |"
        " Operating window (% design) | Turndown | Eff (%) | N_actual | Height (ft) |"
        " Total dP (psi) |",
        "|---|---|---|---|---|---|---|---|---|---|---|",
    ]
    for r in results:
        dc = "n/a" if r.dc_backup_pct is None else f"{r.dc_backup_pct:.0f}"
        lines.append(
            f"| {r.tray.name} | {r.diameter_ft:.2f} | {r.u_nf_ft_s:.2f} |"
            f" {r.dp_tray_psi:.4f} | {dc} | {r.min_load_pct:.0f}-{r.max_load_pct:.0f} |"
            f" {r.turndown_ratio:.2f}:1 | {r.e_tray_pct:.1f} | {r.n_actual_trays} |"
            f" {r.column_height_ft:.1f} | {r.total_dp_psi:.3f} |"
        )
    return lines


def write_report(all_results, matrix, path: str) -> None:
    a_lines: list[str] = []
    a = a_lines.append

    a("# HyTrays Tray Selection Study -- Numerical Comparison Across Services\n")
    a("The conventional **Sieve** baseline and the six HyTrays HT-series "
      "contacting decks -- **HT-01A Hinge, HT-01B Spring, HT-01C LipSeal, "
      "HT-02 CVS, HT-03 GRADEX and HT-05 PULSAR** -- are sized and rated with "
      "the Fair-correlation tray-hydraulics engine in `column_hydraulics.py` "
      "across five representative distillation services. The goal is to make "
      "the trade-offs that drive tray selection (capacity, pressure drop, "
      "downcomer backup, turndown and fouling resistance) explicit and "
      "reproducible.\n")
    a("> The DCX (HT-04 downcomer module), VortiValve (HT-06 inlet device) and "
      "AEGIS (HT-07 structural overlay) HT-series members are not standalone "
      "contacting decks and so are not sized by this 1-D engine -- see "
      "`tray_library.NON_DECK_MODULES`.\n")

    a("## 1. Tray types compared\n")
    a("| Tray | f_active | f_hole | h_weir (in) | C0 | Aeration factor |"
      " Capacity factor | Efficiency factor | Min load frac | Fouling open-"
      "area retention |")
    a("|---|---|---|---|---|---|---|---|---|---|")
    for t in ALL_TRAYS:
        a(f"| {t.name} | {t.f_active:.2f} | {t.f_hole:.2f} | {t.h_weir_in:.1f} |"
          f" {t.c0:.2f} | {t.aeration_factor:.2f} | {t.capacity_factor:.2f} |"
          f" {t.efficiency_factor:.2f} | {t.min_load_frac:.2f} |"
          f" {t.fouling_open_area_retention:.2f} |")
    a("")
    a("> *HT-series parameters are derived from the datasheets in "
      "`HyTrays/Datasheets/` and are engineering estimates for relative "
      "screening -- refine against detailed vendor/test data before absolute "
      "design.*\n")

    a("## 2. Services evaluated\n")
    a("| Service | V (lb/hr) | L (lb/hr) | rho_V (lb/ft3) | rho_L (lb/ft3) |"
      " sigma (dyn/cm) | mu_L (cP) | alpha | Tray spacing (in) | f_flood |"
      " System factor | Fouling? | N_theoretical |")
    a("|---|---|---|---|---|---|---|---|---|---|---|---|---|")
    for svc in SERVICES:
        a(f"| {svc.name} | {svc.v_mass_lb_hr:,.0f} | {svc.l_mass_lb_hr:,.0f} |"
          f" {svc.rho_v:.3f} | {svc.rho_l:.1f} | {svc.sigma:.1f} | {svc.mu_l:.2f} |"
          f" {svc.alpha:.2f} | {svc.tray_spacing_in:.0f} | {svc.f_flood:.0%} |"
          f" {svc.system_factor:.2f} | {svc.fouling} | {svc.n_theoretical} |")
    a("")
    for svc in SERVICES:
        a(f"- **{svc.name}**: {svc.description}")
    a("")

    a("## 3. Cross-service comparison\n")
    a("![Diameter by service](11_diameter_by_service.png)\n")
    a("![Total dP by service](12_total_dp_by_service.png)\n")
    a("![Downcomer backup by service](13_downcomer_backup_by_service.png)\n")
    a("![Turndown by service](14_turndown_by_service.png)\n")
    a("![Fouling sensitivity](15_fouling_sensitivity.png)\n")

    a("## 4. Results & discussion by service\n")

    # -- General --------------------------------------------------------
    gen = all_results["General Rectification (T801 basis)"]
    by = {r.tray.name: r for r in gen}
    sieve = by.get("Sieve", gen[0])
    hy = [r for r in gen if r is not sieve]
    smallest = min(gen, key=lambda r: r.diameter_ft)
    most_eff = max(gen, key=lambda r: r.e_tray_pct)
    fewest = min(gen, key=lambda r: r.n_actual_trays)
    lowest_total = min(gen, key=lambda r: r.total_dp_psi)
    a("### 4.1 General Rectification (T801 basis)\n")
    a("\n".join(results_table_md(gen)))
    a("")
    a(f"FLV = {gen[0].flv:.3f} (vapor-dominated). The smallest shell is "
      f"**{smallest.tray.name}** ({smallest.diameter_ft:.2f} ft vs "
      f"{sieve.diameter_ft:.2f} ft for the Sieve baseline), and the highest "
      f"efficiency is **{most_eff.tray.name}** ({most_eff.e_tray_pct:.1f}%), "
      f"giving the shortest column ({fewest.tray.name}: "
      f"{fewest.column_height_ft:.0f} ft, {fewest.n_actual_trays} trays). The "
      f"lowest total column dP is **{lowest_total.tray.name}** "
      f"({lowest_total.total_dp_psi:.2f} psi vs {sieve.total_dp_psi:.2f} psi "
      f"for Sieve). Across the HT-series decks the efficiency uplift "
      f"(HT-03 GRADEX / HT-05 PULSAR) and capacity (HT-02 CVS) are the main "
      f"levers in this clean, vapor-dominated service.\n")

    # -- Fouling ----------------------------------------------------------
    foul = all_results["Fouling / Heavy-Ends Service"]
    by = {r.tray.name: r for r in foul}
    a("### 4.2 Fouling / Heavy-Ends Service\n")
    a("\n".join(results_table_md(foul)))
    a("")
    a("Sized at 65% of flood (margin for deposit buildup) with each tray's "
      "hole area derated by `fouling_open_area_retention`. The standalone "
      "fouling-sensitivity result (Fig. 15, independent of flow rates) shows "
      "the underlying mechanism -- for a tray whose open area survives at "
      "fraction *r* of as-new, dry-tray dP rises by (1/r)^2 - 1:\n")
    a("| Tray | Open-area retention | Dry-tray dP increase |")
    a("|---|---|---|")
    for t in ALL_TRAYS:
        pct = ((1.0 / t.fouling_open_area_retention) ** 2 - 1.0) * 100.0
        a(f"| {t.name} | {t.fouling_open_area_retention:.2f} | +{pct:.0f}% |")
    a("")
    best_foul = max(ALL_TRAYS, key=lambda t: t.fouling_open_area_retention)
    worst_foul = min(ALL_TRAYS, key=lambda t: t.fouling_open_area_retention)
    sieve_t = next((t for t in ALL_TRAYS if t.name == "Sieve"), ALL_TRAYS[0])
    a(f"The most fouling-tolerant tray here is **{best_foul.name}** "
      f"(retention {best_foul.fouling_open_area_retention:.2f}, only "
      f"+{((1/best_foul.fouling_open_area_retention)**2-1)*100:.0f}% dry-tray "
      f"dP as deposits build), keeping it usable far longer between cleanings "
      f"-- by design for HT-05 PULSAR, whose self-sweeping jet and lack of "
      f"moving parts resist plugging. The least tolerant is "
      f"**{worst_foul.name}** "
      f"(+{((1/worst_foul.fouling_open_area_retention)**2-1)*100:.0f}%; "
      f"crevices/moving parts), while the plain Sieve baseline rises by "
      f"roughly {((1/sieve_t.fouling_open_area_retention)**2-1)*100:.0f}%. "
      f"The adaptive HT-01 hinge/lip decks sit lower than PULSAR because "
      f"their moving flaps and lips offer more crevices for deposits.\n")

    # -- Vacuum -------------------------------------------------------------
    vac = all_results["Vacuum Tower Section"]
    by = {r.tray.name: r for r in vac}
    a("### 4.3 Vacuum Tower Section\n")
    a("\n".join(results_table_md(vac)))
    a("")
    low_dp = min(vac, key=lambda r: r.dp_tray_psi)
    high_dp = max(vac, key=lambda r: r.dp_tray_psi)
    a(f"With rho_V = {VACUUM_RHO_V} lb/ft3, superficial velocities and "
      f"diameters are large for all trays ({min(r.diameter_ft for r in vac):.1f}-"
      f"{max(r.diameter_ft for r in vac):.1f} ft). Per-tray dP is what "
      f"matters most here, because every inch of tray dP raises the flash-zone "
      f"temperature. The lowest per-tray dP is **{low_dp.tray.name}** "
      f"({low_dp.dp_tray_psi*1000:.1f} mpsi, C0 = {low_dp.tray.c0:.2f}) and "
      f"the highest is **{high_dp.tray.name}** "
      f"({high_dp.dp_tray_psi*1000:.0f} mpsi, C0 = {high_dp.tray.c0:.2f}). The "
      f"HT-series decks with higher discharge coefficients (HT-02 CVS swirl "
      f"tubes, HT-03 GRADEX push valves) hold the per-tray dP down, while "
      f"low-C0 sieve-like decks run highest -- multiplied over a real vacuum "
      f"tower's tray count, that gap is what drives vacuum-service designs "
      f"toward higher-C0 or high-capacity internals.\n")

    # -- High-P/High-L ------------------------------------------------------
    hpl = all_results["High-Pressure / High-Liquid-Load Service (C3/C4 splitter)"]
    by = {r.tray.name: r for r in hpl}
    a("### 4.4 High-Pressure / High-Liquid-Load Service (C3/C4 splitter)\n")
    a("\n".join(results_table_md(hpl)))
    a("")
    sieve_hpl = by.get("Sieve", hpl[0])
    sieve_gen = next((r for r in gen if r.tray.name == "Sieve"), gen[0])
    worst_dc = max((r for r in hpl if r.dc_backup_pct is not None),
                   key=lambda r: r.dc_backup_pct)
    a(f"FLV = {hpl[0].flv:.3f} (high, liquid-dominated). Downcomer backup "
      f"margins shrink markedly versus the General case "
      f"({sieve_hpl.dc_backup_pct:.0f}% vs {sieve_gen.dc_backup_pct:.0f}% for "
      f"the Sieve baseline) because the smaller diameters needed for "
      f"high-capacity trays shorten the weir, increasing the Francis-weir "
      f"crest (h_ow) for the same liquid rate. The tightest downcomer margin "
      f"here is **{worst_dc.tray.name}** ({worst_dc.dc_backup_pct:.0f}% of the "
      f"50% limit) -- the high-capacity HT-series decks (e.g. HT-02 CVS) buy "
      f"the smallest shell but pay for it in downcomer loading at high FLV. A "
      f"real design at this FLV would likely need wider downcomers or a larger "
      f"diameter than the flood-only sizing shown; the HT-04 DCX active "
      f"downcomer module (see `NON_DECK_MODULES`) is aimed squarely at this "
      f"high-weir-loading regime.\n")

    # -- Foaming --------------------------------------------------------------
    foam = all_results["Foaming Service"]
    by = {r.tray.name: r for r in foam}
    gen_by = {r.tray.name: r for r in gen}
    a("### 4.5 Foaming Service\n")
    a("\n".join(results_table_md(foam)))
    a("")
    sieve_foam = by.get("Sieve", foam[0])
    sieve_gen2 = gen_by.get("Sieve", gen[0])
    low_aer = min(foam, key=lambda r: r.tray.aeration_factor)
    high_aer = max(foam, key=lambda r: r.tray.aeration_factor)
    a(f"Applying Kister's moderate-foam system factor (Fp = "
      f"{FOAMING_SYSTEM_FACTOR:.2f}) derates every tray's flooding velocity "
      f"equally, so diameters grow by 1/sqrt(Fp) = "
      f"{1/FOAMING_SYSTEM_FACTOR**0.5:.2f}x relative to the General case "
      f"(e.g. Sieve {sieve_gen2.diameter_ft:.2f} -> "
      f"{sieve_foam.diameter_ft:.2f} ft) for every tray type -- the relative "
      f"ranking by capacity is unchanged. The differentiator in foaming "
      f"service is froth intensity: trays with a lower aeration factor "
      f"generate less aerated froth for the same clear-liquid height and are "
      f"generally more foam-tolerant. The lowest here is **{low_aer.tray.name}** "
      f"(aeration {low_aer.tray.aeration_factor:.2f}) and the highest is "
      f"**{high_aer.tray.name}** ({high_aer.tray.aeration_factor:.2f}); the "
      f"capacity-oriented HT-series decks (HT-02 CVS, HT-03 GRADEX) run leaner "
      f"froth than the adaptive sieve-like HT-01 decks.\n")

    a("## 5. Tray selection guidance matrix\n")
    a("Rank 1 = best, 5 = worst, by the metric noted for each service "
      "(footnotes below).\n")
    header = ["Tray"] + [f"{SHORT_NAME[s]} [{i+1}]" for i, (s, *_rest) in enumerate(RANKING_SPECS)]
    a("| " + " | ".join(header) + " |")
    a("|" + "---|" * len(header))
    for t in ALL_TRAYS:
        row = [t.name]
        for svc_name, *_ in RANKING_SPECS:
            rank, val = matrix[t.name][svc_name]
            if val is None:
                row.append(f"{rank} (n/a)")
            else:
                row.append(f"{rank} ({val:.3g})")
        a("| " + " | ".join(row) + " |")
    a("")
    for i, (svc_name, _key, _dir, note) in enumerate(RANKING_SPECS):
        a(f"[{i+1}] **{SHORT_NAME[svc_name]}** ({svc_name}): {note}")
    a("")

    a("## 6. Conclusions & limitations\n")
    a("- **HT-02 CVS** (centrifugal swirl) wins on raw capacity -- the "
      "smallest shell across services and capacity decoupled from tray "
      "spacing -- making it the pick where shell diameter / plot space is the "
      "controlling constraint, at the cost of narrower turndown and tighter "
      "downcomer margins at high FLV.")
    a("- **HT-03 GRADEX** (radially-graded push valves) is the efficiency "
      "lever, especially on large-diameter trays (Peclet/plug-flow gain), "
      "shortening the column where many stages are needed.")
    a("- **HT-05 PULSAR** (fluidic oscillator) is the fouling specialist: the "
      "best open-area retention in the family with no moving parts, plus a "
      "useful efficiency bump -- the choice for deposit-forming service.")
    a("- **HT-01A Hinge / HT-01B Spring** (adaptive flap decks) deliver the "
      "widest turndown in the family for columns that must run efficiently "
      "over a very wide load range; the spring variant trades a little "
      "turndown for robustness and better fouling resistance.")
    a("- **HT-01C LipSeal** (check-valve lips) is the straightforward "
      "weep-resistant upgrade from plain sieve where mild turndown is the "
      "issue.")
    a("- **Sieve** remains the simplest/cheapest baseline, included here only "
      "as the reference the HT-series factors are measured against.")
    a("- The **DCX (HT-04)**, **VortiValve (HT-06)** and **AEGIS (HT-07)** "
      "family members are not standalone decks and are not rated here; DCX in "
      "particular targets the high-weir-loading regime exposed by the "
      "High-P/High-L service above.")
    a("")
    a("**Model limitations** -- read before using these numbers for "
      "anything beyond relative tray-type screening:")
    a("- Fair's C_SBF curve fit is valid for FLV = 0.01-1.0 and tray "
      "spacing 6-24 in (the Vacuum case uses 30 in, slightly outside the "
      "fitted range, and the High-P/High-L case has FLV at the top of the "
      "range).")
    a("- Turndown/weep is represented by each tray's datasheet-derived "
      "`min_load_frac`, not a first-principles weep-point correlation; "
      "entrainment is not explicitly modeled.")
    a("- The HT-series decks are modeled as single bubbling decks on a "
      "standard weir/downcomer layout; device-specific physics (centrifugal "
      "swirl, fluidic oscillation, adaptive-flap dynamics) are folded into "
      "the capacity / efficiency / turndown / open-area factors rather than "
      "resolved mechanistically.")
    a("- All HT-series parameters are engineering estimates derived from the "
      "datasheets in `HyTrays/Datasheets/` for *relative* screening -- refine "
      "against detailed vendor/test data before absolute design.")
    a("")
    a("---")
    a("*Generated by `HyTrays/run_service_comparison.py`. Edit "
      "`service_cases.py` (operating conditions per service) or "
      "`tray_library.py` (tray parameters) and re-run to update every "
      "table, figure and number in this report.*")

    with open(path, "w") as fh:
        fh.write("\n".join(a_lines) + "\n")


VACUUM_RHO_V = SERVICES[2].rho_v
FOAMING_SYSTEM_FACTOR = SERVICES[4].system_factor


def main() -> None:
    os.makedirs(OUT_DIR, exist_ok=True)
    all_results = run_all()
    for svc_name, results in all_results.items():
        print_service_table(svc_name, results)
    write_csv(all_results, os.path.join(OUT_DIR, "service_comparison.csv"))
    make_plots(all_results, OUT_DIR)
    matrix = build_selection_matrix(all_results)
    write_report(all_results, matrix, os.path.join(OUT_DIR, "tray_selection_study.md"))
    print(f"\nWrote CSV, plots and report to {OUT_DIR}")


if __name__ == "__main__":
    main()
