#!/usr/bin/env python3
"""HyTrays — tray-type parameter library (HT-series invention family).

Defines the geometric / hydraulic "fingerprint" of each tray family used by
the distillation-column comparison engine (``column_hydraulics.py``) and the
run scripts (``run_tray_comparison.py`` / ``run_service_comparison.py``).

Each :class:`TrayType` carries the parameters needed to evaluate Fair's
flooding correlation, dry/wet tray pressure drop, downcomer backup,
turndown and a relative (O'Connell-based) tray efficiency -- see
``column_hydraulics.py`` for the equations. All capacity / efficiency /
turndown factors are expressed *relative to the plain sieve baseline*
(``SIEVE``), which is included as the conventional reference, not as a
HyTrays product.

The HyTrays HT-series (datasheets in ``HyTrays/Datasheets/``)
------------------------------------------------------------------
The family spans several distinct functional classes. Seven are *standalone
contacting decks* that this 1-D Fair-correlation engine can size and rate,
and they are represented as :class:`TrayType` entries below:

    HT-01A  Hinge      -- coined-land living-hinge adaptive deck
    HT-01B  Spring     -- leaf-spring bent-tab adaptive deck
    HT-01C  LipSeal    -- depending check-valve lip dual-flow deck
    HT-02   CVS        -- centrifugal vapour-swirl tube tray
    HT-03   GRADEX     -- radially-graded push-valve tray
    HT-05   PULSAR     -- fluidic-oscillator self-sweeping tray
    HT-08   VORTEXA    -- helical-indexed variable-slot swirl tray
                          (load-tracking turndown extension of HT-02)

Three further family members are **not** standalone contacting decks and so
are deliberately *not* modeled as :class:`TrayType` entries -- the engine
sizes a bubbling deck, and these change boundary conditions rather than the
deck itself (see ``NON_DECK_MODULES`` for the descriptive catalog):

    HT-04   DCX        -- active downcomer module (structured-packing insert
                          + vapour slipstream); a bolt-on that adds NTS to
                          *any* deck rather than being a deck in its own right.
    HT-06   VortiValve -- column inlet/feed device (Tesla-cascade momentum
                          killer + swirl de-aeration + drip distributor); a
                          feed-conditioning internal, not a contacting tray.
    HT-07   AEGIS      -- surge-tolerant structural overlay (hook clamps,
                          rib stiffening, relief flaps); leaves the normal
                          hydraulics unchanged, so it has no distinct
                          hydraulic fingerprint to rate.

Parameter sources & status
--------------------------
``SIEVE`` uses values representative of the published ranges for the plain
perforated sieve tray (Kister, *Distillation Design* (1992) & *Distillation
Operation* (1990); Perry's Chemical Engineers' Handbook, Sec. 18; Lockett,
*Distillation Tray Fundamentals* (1986)).

The six HT-series decks are parameterised from the qualitative/quantitative
behaviour stated in their datasheets (capacity envelope, turndown range,
efficiency uplift, open-area and fouling notes). Where a datasheet gives a
range, the value below is a representative design-point pick; per-tray notes
record the datasheet figure each parameter is anchored to. These remain
engineering estimates for *relative* screening -- refine against detailed
vendor/test data before using them for absolute design.
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

    # ── fouling resistance ───────────────────────────────────────────────
    # Fraction of as-new hole/slot open area expected to remain available
    # once typical service deposits have built up (qualitative, from the
    # tray-selection guidance in Kister, *Distillation Operation* Ch.1:
    # large simple round holes/grids resist plugging best, small holes and
    # tray hardware with moving parts/crevices resist it least).
    fouling_open_area_retention: float = 1.00

    has_downcomer: bool = True


# ── conventional reference baseline ─────────────────────────────────────────
SIEVE = TrayType(
    name="Sieve",
    description="Conventional perforated sieve tray -- the conventional "
                 "reference baseline (not a HyTrays product); all HT-series "
                 "factors below are relative to this tray.",
    f_active=0.78, f_hole=0.10, h_weir_in=2.0,
    c0=0.73, dp_dry_floor_in=0.00,
    aeration_factor=0.55,
    capacity_factor=1.00, efficiency_factor=1.00, min_load_frac=0.50,
    fouling_open_area_retention=0.75,
)

# ── HyTrays HT-series contacting decks ──────────────────────────────────────
HT01A_HINGE = TrayType(
    name="HT-01A Hinge",
    description="Coined-land living-hinge adaptive deck: flaps open with "
                 "vapour load so aggregate open area swings from ~3-6% at low "
                 "rate to ~14-18% at high rate, giving very wide turndown "
                 "(~12-15:1). Pivots are crevice/moving-part features, so "
                 "fouling resistance is modest.",
    # f_hole = representative design-point open area (upper end of the 3-18%
    # adaptive swing); turndown ~12-15:1 -> min_load_frac ~ 1/13 (datasheet HT-01A).
    f_active=0.80, f_hole=0.14, h_weir_in=2.0,
    c0=0.73, dp_dry_floor_in=0.00,
    aeration_factor=0.55,
    capacity_factor=1.05, efficiency_factor=1.05, min_load_frac=0.08,
    fouling_open_area_retention=0.70,
)

HT01B_SPRING = TrayType(
    name="HT-01B Spring",
    description="Leaf-spring bent-tab adaptive deck: same load-following open "
                 "area as the hinge but with distributed (~12x lower) spring "
                 "stress and zero corrodible/separate parts, so it is more "
                 "robust and slightly less crevice-prone. Crack-open dP set by "
                 "the spring family (1.2/2.4/4.8 mbar).",
    # Adaptive like HT-01A but the stiffer, more robust spring families give a
    # slightly tighter (~11:1) but more reliable window (datasheet HT-01B).
    f_active=0.80, f_hole=0.13, h_weir_in=2.0,
    c0=0.73, dp_dry_floor_in=0.00,
    aeration_factor=0.55,
    capacity_factor=1.05, efficiency_factor=1.05, min_load_frac=0.09,
    fouling_open_area_retention=0.78,
)

HT01C_LIPSEAL = TrayType(
    name="HT-01C LipSeal",
    description="Depending check-valve lip deck (dual-flow, directional "
                 "bias): lip valves (12/22/35 deg families) seal against weep "
                 "at low rate, giving a >=30% wider stable window than a plain "
                 "sieve and 5-10x higher weep head.",
    # Check-valve lips lift the weep limit: window ~30% wider than sieve
    # (sieve 2:1 -> ~2.6:1, min_load_frac ~0.38). Lip valves -> higher C0,
    # small valve-weight floor (datasheet HT-01C).
    f_active=0.82, f_hole=0.13, h_weir_in=1.5,
    c0=0.78, dp_dry_floor_in=0.10,
    aeration_factor=0.52,
    capacity_factor=1.08, efficiency_factor=1.02, min_load_frac=0.38,
    fouling_open_area_retention=0.78,
)

HT02_CVS = TrayType(
    name="HT-02 CVS",
    description="Centrifugal vapour-swirl tube deck: swirl tubes spin the "
                 "froth at 30-60 g, raising Fair capacity by ~50-80% (Csb "
                 "0.14-0.18 m/s) and decoupling capacity from tray spacing "
                 "(usable down to ~300 mm). High point efficiency (EOG ~0.86), "
                 "but needs a minimum vapour rate to sustain swirl (narrower "
                 "turndown).",
    # capacity_factor ~1.65 from the stated +50-80% Csb; smooth tubes ->
    # higher C0 and good fouling resistance; swirl needs minimum velocity ->
    # min_load_frac ~0.40 (~2.5:1) (datasheet HT-02).
    f_active=0.75, f_hole=0.15, h_weir_in=2.0,
    c0=0.80, dp_dry_floor_in=0.20,
    aeration_factor=0.50,
    capacity_factor=1.65, efficiency_factor=1.00, min_load_frac=0.40,
    fouling_open_area_retention=0.80,
)

HT03_GRADEX = TrayType(
    name="HT-03 GRADEX",
    description="Radially-graded push-valve deck: open area graded 8%->13% "
                 "across the radius with an azimuth map from a flow-field "
                 "solver and a chord-tuned picket weir. Pushes the deck toward "
                 "plug flow (Peclet ~10 -> ~30), adding +10-18 Murphree points "
                 "on large (D>3.5 m) trays.",
    # Big efficiency lever on large diameters from the Peclet/plug-flow gain;
    # push valves give valve-like capacity and turndown (datasheet HT-03).
    f_active=0.82, f_hole=0.11, h_weir_in=1.5,
    c0=0.82, dp_dry_floor_in=0.30,
    aeration_factor=0.50,
    capacity_factor=1.20, efficiency_factor=1.12, min_load_frac=0.30,
    fouling_open_area_retention=0.75,
)

HT05_PULSAR = TrayType(
    name="HT-05 PULSAR",
    description="Bi-stable fluidic-oscillator deck: Coanda-island + feedback "
                 "channels generate a self-sweeping jet (5-30 Hz) with no "
                 "moving parts, giving long fouling run-lengths (~5-6 yr) and "
                 "+8 EOG points. Turndown ~3.5-4.5:1; oscillator discharge "
                 "coefficient Co,osc ~0.68-0.74.",
    # Self-cleaning, no moving parts -> best fouling retention in the family;
    # +8 EOG points -> efficiency_factor ~1.10; turndown ~4:1 -> min_load_frac
    # ~0.25; C0 from Co,osc ~0.71 (datasheet HT-05).
    f_active=0.80, f_hole=0.12, h_weir_in=2.0,
    c0=0.71, dp_dry_floor_in=0.00,
    aeration_factor=0.55,
    capacity_factor=1.05, efficiency_factor=1.10, min_load_frac=0.25,
    fouling_open_area_retention=0.95,
)

HT08_VORTEXA = TrayType(
    name="HT-08 VORTEXA",
    description="Helical-indexed variable-slot swirl deck: a load-tracking "
                 "sleeve retrofit to HT-02's swirl tubes, kinematically "
                 "coupling axial lift to rotation via a back-drivable helical "
                 "constraint so tangential slot area opens/closes passively "
                 "with vapour load. Preserves HT-02's Csb ~0.14-0.18 m/s "
                 "capacity and <0.02 entrainment at 90% flood, at a slightly "
                 "lower per-tray dP (~11 vs ~12 mbar), while extending "
                 "turndown from HT-02's ~2.5:1 to a predicted ~8:1 (cell). "
                 "Failure mode is benign: a stuck sleeve simply reverts to "
                 "fixed-geometry HT-02 behaviour. Not recommended for "
                 "heavy-fouling duty -- HT-02 (no moving parts) remains "
                 "preferred there.",
    # Same tube/slot layout and capacity as HT-02 (Table 2: Csb, entrainment
    # unchanged); slightly higher C0 from the ~11 vs ~12 mbar dP/stage figure.
    f_active=0.75, f_hole=0.15, h_weir_in=2.0,
    c0=0.82, dp_dry_floor_in=0.20,
    aeration_factor=0.50,
    # EOG 0.85 vs HT-02's 0.86 (Table 2) -> efficiency_factor scaled down
    # proportionally from HT-02's 1.00. Turndown ~8:1 (cell) at f_flood=0.80
    # -> max_load_pct=125% -> min_load_frac = (125/8)/125 = 1/8.
    capacity_factor=1.65, efficiency_factor=0.98, min_load_frac=0.125,
    fouling_open_area_retention=0.75,
)


# Engine-rateable trays: the conventional baseline + the seven HT-series
# contacting decks. The three non-deck modules (below) are documented but not
# sized here.
ALL_TRAYS: list[TrayType] = [
    SIEVE,
    HT01A_HINGE,
    HT01B_SPRING,
    HT01C_LIPSEAL,
    HT02_CVS,
    HT03_GRADEX,
    HT05_PULSAR,
    HT08_VORTEXA,
]


# ── HyTrays HT-series members that are NOT standalone contacting decks ───────
# These change the column's boundary conditions (downcomer behaviour, feed
# conditioning, structural integrity) rather than the bubbling deck the 1-D
# Fair engine sizes, so they are catalogued here for completeness rather than
# rated as TrayType entries.
@dataclass(frozen=True)
class NonDeckModule:
    code: str
    name: str
    klass: str          # functional class
    description: str
    why_not_rated: str   # why it is not a TrayType in this engine


NON_DECK_MODULES: list[NonDeckModule] = [
    NonDeckModule(
        code="HT-04", name="DCX", klass="Active downcomer module",
        description="Bolt-on downcomer with a structured-packing insert plus a "
                     "3-8% vapour slipstream; adds +0.10-0.25 NTS per tray and "
                     "raises weir loading ~26%. Fits onto ANY deck type.",
        why_not_rated="It augments the downcomer of whatever deck it is bolted "
                       "to (an add-on stage), not a bubbling deck the Fair "
                       "engine can size on its own.",
    ),
    NonDeckModule(
        code="HT-06", name="VortiValve", klass="Column inlet / feed device",
        description="Tesla-cascade momentum killer + gentle swirl de-aeration "
                     "+ drip-tube distributor; <5% maldistribution at <20 mbar "
                     "dP. A feed-conditioning internal, not a contacting tray.",
        why_not_rated="It conditions the inlet stream; it does no vapour/liquid "
                       "contacting, so it has no flooding/efficiency fingerprint "
                       "to rate.",
    ),
    NonDeckModule(
        code="HT-07", name="AEGIS", klass="Surge-tolerant structural overlay",
        description="Hook clamps + rib stiffening + relief flaps that let a "
                     "tray survive ~5 psi uplift; normal hydraulics unchanged.",
        why_not_rated="A structural/mechanical overlay -- by design it leaves "
                       "the normal tray hydraulics unchanged, so it has no "
                       "distinct hydraulic fingerprint.",
    ),
]
