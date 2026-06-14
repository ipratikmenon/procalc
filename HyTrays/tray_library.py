#!/usr/bin/env python3
"""HyTrays — tray-type parameter library.

Defines the geometric / hydraulic "fingerprint" of each tray family used by
the distillation-column comparison engine (``column_hydraulics.py``) and run
script (``run_tray_comparison.py``).

Each :class:`TrayType` carries the parameters needed to evaluate Fair's
flooding correlation, dry/wet tray pressure drop, downcomer backup,
turndown and a relative (O'Connell-based) tray efficiency -- see
``column_hydraulics.py`` for the equations.

Parameter sources
------------------
``SIEVE``, ``VALVE``, ``DUALFLOW`` and ``HIGH_PERFORMANCE`` use values
representative of the published ranges for these well-known tray families
(Kister, *Distillation Design* (1992) & *Distillation Operation* (1990);
Perry's Chemical Engineers' Handbook, Sec. 18; Lockett, *Distillation Tray
Fundamentals* (1986)).

``RIPPLE`` represents the new HyTrays corrugated-deck sieve tray (the
products referenced by the datasheets in ``HyTrays/Datasheets/``).  Its
capacity / efficiency / turndown factors below are *placeholders* set from
the qualitative behaviour described for corrugated/rippled decks (extra
interfacial area, better low-rate liquid retention).  **Replace
``capacity_factor``, ``efficiency_factor``, ``min_load_frac``,
``f_active``, ``f_hole``, ``c0`` and ``h_weir_in`` with the values from the
manufacturer datasheet once it is added to ``HyTrays/Datasheets/``.**
"""
from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class TrayType:
    """One tray family's hydraulic / performance fingerprint.

    All geometric fractions are relative to the column cross-section
    (``A_col``) unless noted otherwise.
    """

    name: str
    description: str

    # ── geometry ────────────────────────────────────────────────────────
    f_active: float    # bubbling/active area / column area  (A_a / A_col)
    f_hole: float      # hole (or valve) open area / active area (A_h / A_a)
    h_weir_in: float   # outlet weir height, inches (0 for dualflow)

    # ── dry-tray pressure drop ─────────────────────────────────────────
    c0: float              # orifice discharge coefficient
    dp_dry_floor_in: float  # minimum dry DP, in. liquid (valve-weight floor)

    # ── wet-tray / downcomer ────────────────────────────────────────────
    aeration_factor: float  # beta: froth density factor on clear-liquid head

    # ── capacity / efficiency / turndown, relative to plain sieve ───────
    capacity_factor: float    # multiplier on Fair's C_SBF (system factor)
    efficiency_factor: float  # multiplier on O'Connell point efficiency
    min_load_frac: float      # min stable vapor load / flood-limited design load

    has_downcomer: bool = True


SIEVE = TrayType(
    name="Sieve",
    description="Conventional perforated sieve tray (baseline).",
    f_active=0.78, f_hole=0.10, h_weir_in=2.0,
    c0=0.73, dp_dry_floor_in=0.00,
    aeration_factor=0.55,
    capacity_factor=1.00, efficiency_factor=1.00, min_load_frac=0.50,
)

RIPPLE = TrayType(
    name="HyTrays Ripple",
    description="HyTrays corrugated-deck sieve tray (new product family -- "
                 "see HyTrays/Datasheets/ for source PDFs; parameters below "
                 "are placeholders pending those datasheets).",
    f_active=0.80, f_hole=0.11, h_weir_in=1.5,
    c0=0.74, dp_dry_floor_in=0.00,
    aeration_factor=0.50,
    capacity_factor=1.10, efficiency_factor=1.08, min_load_frac=0.35,
)

VALVE = TrayType(
    name="Valve",
    description="Conventional moving (round-cap) valve tray.",
    f_active=0.78, f_hole=0.13, h_weir_in=2.0,
    c0=0.85, dp_dry_floor_in=0.40,
    aeration_factor=0.55,
    capacity_factor=1.00, efficiency_factor=1.00, min_load_frac=0.30,
)

DUALFLOW = TrayType(
    name="Dualflow",
    description="Perforated grid / dualflow tray -- no downcomers, "
                 "vapor and liquid share the same perforations.",
    f_active=1.00, f_hole=0.20, h_weir_in=0.0,
    c0=0.73, dp_dry_floor_in=0.00,
    aeration_factor=0.40,
    capacity_factor=1.15, efficiency_factor=0.85, min_load_frac=0.65,
    has_downcomer=False,
)

HIGH_PERFORMANCE = TrayType(
    name="High-Performance",
    description="Generic high-capacity multi-downcomer / fixed-valve tray "
                 "(e.g. MVG, V-Grid, ConSep-class designs).",
    f_active=0.85, f_hole=0.14, h_weir_in=1.0,
    c0=0.80, dp_dry_floor_in=0.00,
    aeration_factor=0.50,
    capacity_factor=1.25, efficiency_factor=0.97, min_load_frac=0.40,
)

ALL_TRAYS: list[TrayType] = [SIEVE, RIPPLE, VALVE, DUALFLOW, HIGH_PERFORMANCE]
