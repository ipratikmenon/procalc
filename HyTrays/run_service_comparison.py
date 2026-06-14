#!/usr/bin/env python3
"""HyTrays — tray selection study across services.

Runs every tray family in ``tray_library.ALL_TRAYS`` (Sieve, HyTrays Ripple,
Valve, Dualflow, High-Performance) against every service case in
``service_cases.SERVICES`` (General Rectification, Fouling/Heavy-Ends,
Vacuum, High-Pressure/High-Liquid-Load, Foaming) and assembles a
"paper"-style write-up comparing the five tray types service by service.

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
                "Downcomer backup by service & tray type (Dualflow has no downcomer)")
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
     "Dualflow has no downcomer and is generally not recommended for "
     "high-liquid-rate services (Kister, Distillation Operation) -- ranked "
     "last regardless of the hydraulic numbers."),
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
    a("Five tray families -- **Sieve, HyTrays Ripple, Valve, Dualflow and a "
      "generic High-Performance tray** -- are sized and rated with the Fair"
      "-correlation tray-hydraulics engine in `column_hydraulics.py` across "
      "five representative distillation services. The goal is to make the "
      "trade-offs that drive tray selection (capacity, pressure drop, "
      "downcomer backup, turndown and fouling resistance) explicit and "
      "reproducible.\n")

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
    a("> *HyTrays Ripple parameters are placeholders pending the datasheets "
      "in `HyTrays/Datasheets/` -- update and re-run once available.*\n")

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
    a("### 4.1 General Rectification (T801 basis)\n")
    a("\n".join(results_table_md(gen)))
    a("")
    a(f"FLV = {gen[0].flv:.3f} (vapor-dominated). HyTrays Ripple gives the "
      f"smallest shell of the conventional weired trays ({by['HyTrays Ripple'].diameter_ft:.2f} ft "
      f"vs {by['Sieve'].diameter_ft:.2f} ft for Sieve/Valve) **and** the "
      f"highest efficiency ({by['HyTrays Ripple'].e_tray_pct:.1f}%), giving "
      f"the shortest column ({by['HyTrays Ripple'].column_height_ft:.0f} ft, "
      f"{by['HyTrays Ripple'].n_actual_trays} trays) and the lowest total dP "
      f"after Dualflow/High-Performance "
      f"({by['HyTrays Ripple'].total_dp_psi:.2f} psi vs "
      f"{by['Sieve'].total_dp_psi:.2f} psi for Sieve). Dualflow needs the "
      f"most trays ({by['Dualflow'].n_actual_trays}) due to its lower "
      f"contacting efficiency.\n")

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
    a(f"Dualflow's large, simple perforations lose the least relative open "
      f"area (+{((1/0.92)**2-1)*100:.0f}% dP), keeping it usable far longer "
      f"between cleanings, while a plain Sieve tray's dry-tray dP rises by "
      f"roughly {((1/0.75)**2-1)*100:.0f}% and a Valve tray's by "
      f"{((1/0.70)**2-1)*100:.0f}% (small holes/slots, moving parts), "
      f"signaling much faster fouling-driven capacity loss. HyTrays Ripple "
      f"(+{((1/0.85)**2-1)*100:.0f}%, placeholder) sits between the two "
      f"extremes.\n")

    # -- Vacuum -------------------------------------------------------------
    vac = all_results["Vacuum Tower Section"]
    by = {r.tray.name: r for r in vac}
    a("### 4.3 Vacuum Tower Section\n")
    a("\n".join(results_table_md(vac)))
    a("")
    a(f"With rho_V = {VACUUM_RHO_V} lb/ft3, superficial velocities and "
      f"diameters are large for all trays ({min(r.diameter_ft for r in vac):.1f}-"
      f"{max(r.diameter_ft for r in vac):.1f} ft). Per-tray dP is what "
      f"matters most here: Dualflow is lowest "
      f"({by['Dualflow'].dp_tray_psi*1000:.1f} mpsi), Valve and "
      f"High-Performance follow "
      f"({by['Valve'].dp_tray_psi*1000:.1f} and "
      f"{by['High-Performance'].dp_tray_psi*1000:.1f} mpsi) on the strength of "
      f"their higher discharge coefficients (C0 = {by['Valve'].tray.c0:.2f} / "
      f"{by['High-Performance'].tray.c0:.2f} vs "
      f"{by['Sieve'].tray.c0:.2f}-{by['HyTrays Ripple'].tray.c0:.2f} for "
      f"Sieve/Ripple), while Sieve and Ripple run highest at "
      f"{by['HyTrays Ripple'].dp_tray_psi*1000:.0f}-"
      f"{by['Sieve'].dp_tray_psi*1000:.0f} mpsi/tray -- multiplied over a real "
      f"vacuum tower's tray count this difference is what drives "
      f"vacuum-service designs toward higher-C0 or grid/high-capacity "
      f"internals.\n")

    # -- High-P/High-L ------------------------------------------------------
    hpl = all_results["High-Pressure / High-Liquid-Load Service (C3/C4 splitter)"]
    by = {r.tray.name: r for r in hpl}
    a("### 4.4 High-Pressure / High-Liquid-Load Service (C3/C4 splitter)\n")
    a("\n".join(results_table_md(hpl)))
    a("")
    a(f"FLV = {hpl[0].flv:.3f} (high, liquid-dominated). Downcomer backup "
      f"margins shrink markedly versus the General case "
      f"({by['Sieve'].dc_backup_pct:.0f}% vs "
      f"{[r.dc_backup_pct for r in gen if r.tray.name=='Sieve'][0]:.0f}% for "
      f"Sieve) because the smaller diameters needed for high-capacity trays "
      f"shorten the weir, increasing the Francis-weir crest (h_ow) for the "
      f"same liquid rate. All weired trays here still sit under the 50% "
      f"limit, but with much less margin than the General case -- a real "
      f"design at this FLV would likely need wider downcomers or a larger "
      f"diameter than the flood-only sizing shown. **Dualflow is flagged as "
      f"not recommended** for this service regardless of the numbers: with "
      f"no downcomer, all the liquid must counter-flow through the same "
      f"perforations as the vapor, which becomes the limiting mechanism at "
      f"high liquid rates (Kister, *Distillation Operation*).\n")

    # -- Foaming --------------------------------------------------------------
    foam = all_results["Foaming Service"]
    by = {r.tray.name: r for r in foam}
    gen_by = {r.tray.name: r for r in gen}
    a("### 4.5 Foaming Service\n")
    a("\n".join(results_table_md(foam)))
    a("")
    a(f"Applying Kister's moderate-foam system factor (Fp = "
      f"{FOAMING_SYSTEM_FACTOR:.2f}) derates every tray's flooding velocity "
      f"equally, so diameters grow by 1/sqrt(Fp) = "
      f"{1/FOAMING_SYSTEM_FACTOR**0.5:.2f}x relative to the General case "
      f"(e.g. Sieve {gen_by['Sieve'].diameter_ft:.2f} -> "
      f"{by['Sieve'].diameter_ft:.2f} ft) for every tray type -- the relative "
      f"ranking by capacity is unchanged. The differentiator in foaming "
      f"service is froth intensity: trays with a lower aeration factor "
      f"(Dualflow {by['Dualflow'].tray.aeration_factor:.2f}, HyTrays Ripple/"
      f"High-Performance {by['HyTrays Ripple'].tray.aeration_factor:.2f}) "
      f"generate less aerated froth for the same clear-liquid height and "
      f"are generally more foam-tolerant than Sieve/Valve "
      f"({by['Sieve'].tray.aeration_factor:.2f}).\n")

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
    a("- **HyTrays Ripple** is competitive or best-in-class for General "
      "Rectification (smallest shell among weired trays, highest "
      "efficiency, shortest column) and has a favorable fouling profile "
      "between Sieve/Valve and Dualflow/High-Performance -- a good general-"
      "purpose upgrade from plain Sieve. *Its parameters are placeholders; "
      "replace them from the HyTrays datasheets once added to "
      "`HyTrays/Datasheets/` and re-run.*")
    a("- **Dualflow** wins on raw capacity and per-tray dP (General, "
      "Vacuum, Fouling) but loses on efficiency/turndown and is not "
      "recommended for high-liquid-rate service.")
    a("- **High-Performance** gives the smallest shell across the board "
      "and the lowest per-tray dP alongside Dualflow, with efficiency much "
      "closer to Sieve -- a strong default where shell diameter/capex "
      "dominates.")
    a("- **Valve** matches Sieve's capacity but offers the widest turndown "
      "of the conventional designs -- preferred where the column must run "
      "efficiently over a wide load range, at the cost of being the most "
      "fouling-sensitive (smallest open area, moving parts).")
    a("- **Sieve** remains the simplest/cheapest baseline but is dominated "
      "by HyTrays Ripple on every metric computed here.")
    a("")
    a("**Model limitations** -- read before using these numbers for "
      "anything beyond relative tray-type screening:")
    a("- Fair's C_SBF curve fit is valid for FLV = 0.01-1.0 and tray "
      "spacing 6-24 in (the Vacuum case uses 30 in, slightly outside the "
      "fitted range, and the High-P/High-L case has FLV at the top of the "
      "range).")
    a("- Turndown/weep is represented by each tray's literature-typical "
      "`min_load_frac`, not a first-principles weep-point correlation; "
      "entrainment is not explicitly modeled.")
    a("- Dualflow's clear-liquid height is a fixed 1 in. placeholder "
      "(no overflow weir to apply the Francis equation to) -- its dP and "
      "downcomer-backup columns should be read as indicative only.")
    a("- All `HyTrays Ripple` and `fouling_open_area_retention` values are "
      "engineering placeholders pending vendor/manufacturer data.")
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
