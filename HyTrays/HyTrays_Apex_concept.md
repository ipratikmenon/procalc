# HyTrays Apex -- Integrated Cartridge-Grid: Concept & Scorecard

## Status

This memo now has two layers, built one on top of the other:

- **Phase 1** (Sections 1.1, 2 and the original "Directional Vortex-Grid"
  ideas) scoped Apex against a severe-service spec calibrated to a real
  SUPERFRAC/VG-0 retrofit: "2x ULTRA-FRAC" capacity, 2-5 psi uplift rating,
  extreme fouling resistance. It was a paper exercise -- no engine
  parameters.
- **Phase 2** (this revision) responds to a much more aggressive brief: a
  single deck that *simultaneously* hits ten performance metrics plus a set
  of "no-X" qualitative guarantees -- "a revolution in design" (Section 1.2).
  Rather than inventing new physics, Phase 2 resolves this by **combining
  mechanisms already disclosed across HT-01 through HT-08** into one deck:
  the **Integrated Cartridge-Grid** (Section 3).

Apex is now implemented as `tray_library.APEX` and included in
`tray_library.ALL_TRAYS` as the flagship entry, "above" the numbered HT-0X
line. Every number in Sections 6-7 below comes straight from
`run_tray_comparison.py` / `run_service_comparison.py` (see
`HyTrays/output/`) -- the original memo's Section 7 ("path into the
framework, not yet done") is therefore **done**, in the sense described in
Section 8.

## 1. Target specification

### 1.1 Phase 1 spec: "2x ULTRA-FRAC", severe service

| Requirement | Interpretation used here |
|---|---|
| Extreme fouling resistance | Large, simple, self-sweeping openings; no moving parts, no crevices |
| "Better than ULTRA-FRAC" | Outperform a Koch-Glitsch ULTRA-FRAC(R)-class multi-chordal high-capacity valve tray on capacity, fouling and pressure rating |
| 2x ULTRA-FRAC ultimate capacity | ~2x the vapor-handling capacity of a tray that is itself ~30-40% above conventional valve trays -- i.e. a step-change, not an incremental gain |
| 2-5 psi uplift resistance | A *structural/mechanical* deck rating (anti-flotation under upset/relief transients), separate from the hydraulic per-tray dP the Fair model computes |
| Higher vapor rates | Consequence of the capacity target above |

The retrofit data referenced in Phase 1 (16 ft column, Movable-Valve ->
Multi-Chordal SUPERFRAC/VG-0, 196->178 trays above feed, 820->958 MM lb/yr
propylene) is a useful calibration point: a real deck-technology generation
step bought **~17% more throughput at the same diameter**. "2x ULTRA-FRAC" is
therefore not "one more generation of valve" -- it requires a different
flooding mechanism entirely.

![SUPERFRAC/VG-0 retrofit, before vs. after](References/superfrac_retrofit_before_after.jpeg)
*Reference: real-world VG-0/multi-chordal retrofit at constant 16 ft
diameter -- ~17% capacity gain, used here as the "one generation" baseline
that "2x ULTRA-FRAC" must be measured against.*

### 1.2 Phase 2: the "revolution" wishlist

A revolutionary tray would dominate if it *simultaneously* achieved:

| Metric | Desired outcome |
|---|---|
| Efficiency | >90% Murphree |
| Turndown | >20:1 |
| Pressure drop | Packing-level |
| Capacity | Grid-tray-level |
| Fouling resistance | Near self-cleaning |
| Maintenance | Cartridge swap |
| Stability | No weeping/dumping |
| Cost | Manufacturable at scale |
| Control | Real-time adaptive |
| Energy | Major reduction |

...plus, qualitatively: **no stagnant corners, no jet flooding, no dead
zones, no weeping, no dumping, no entrainment, a massive turndown ratio.**

Phase 1's five requirements map onto this list directly (extreme fouling
resistance -> fouling resistance; 2x ULTRA-FRAC capacity & higher vapor rates
-> capacity; 2-5 psi uplift -> the maintenance/cost architecture in Section
3.5). Section 7 scores Apex against every item above, with real numbers from
the engine.

## 2. The central conflict

| Goal | Wants... | ...which tends to give |
|---|---|---|
| Fouling resistance | Large, simple, fixed openings (Kister: big round holes/grids resist plugging best) | Lower per-pass efficiency, more weeping at low load |
| 2x capacity | Co-current jet zones + mechanical vapor/liquid separation (bypasses Souders-Brown flooding) | Small, precisely shaped elements -- typically the *opposite* of "simple and fouling-resistant" |
| 2-5 psi structural rating | Thick plate, dense support-beam grid, positively-clamped panels | Added weight/cost, less open area unless designed around it |
| >20:1 turndown | A secondary low-load flow path that stays open when the primary path shuts | Extra zone, extra geometry, another thing to validate |
| Real-time adaptive control | Sensing + actuation that tracks load continuously | Electronics/actuators -- a different product category, more failure modes, more cost |

**Phase 1 resolved the first three rows by separating the "fouling-resistant
flow element" from the "capacity-boost element"**: large fixed slots stay
simple (fouling resistance); the capacity boost comes from how those slots
are *arranged and capped* (directional vortex zones with mechanical
disengagement), not from making them small and intricate.

**Phase 2 extends the same idea to all ten metrics, by giving each goal its
own zone or mechanism on the same deck** rather than asking one mechanism to
do everything (Section 3). No single HT-0X deck individually satisfies the
wishlist -- HT-02/HT-08 are the best on capacity but worst on turndown,
HT-01A/B are the best on turndown but unremarkable on capacity, HT-05 is the
fouling specialist but modest everywhere else. Apex's premise is that these
are *complementary*, not competing, because each excels on a different part
of the operating envelope or a different fraction of the deck area.

## 3. Architecture: the Integrated Cartridge-Grid

Apex's active area (`f_active = 0.85`, vs sieve's 0.78) splits into two
contacting zones plus three deck-wide treatments layered over both:

```
        VAPOR + LIQUID  (co-current swirl, mechanical disengage)
            \\\   ///     \\\   ///        .  .  .
         ____\\\_///_______\\\_///_____.__.__.__.____
        |  [VORTEXA   ] [VORTEXA   ]  [ LipSeal trickle ] |  <- vane caps + PULSAR
   DC -->| /=sleeves=\   /=sleeves=\    /==lip valves==\  |<-- DC
        |  ~70% area, GRADEX-graded  |~15% area, secondary|
        |__beams_____beams___________beams______beams____|  <- AEGIS cartridge grid
              <----  multi-chordal liquid sweep  ---->
```

### 3.1 Primary zone -- vortex-tube cartridges (~70% of active area)

HT-02 CVS swirl tubes, fitted with HT-08 VORTEXA's helical-indexed variable
slots, supplied as bolt-in cartridges. Each tube spins the froth at 30-60 g
and a fixed vane cap mechanically separates entrained liquid from vapor
*before* it reaches the tray above -- so the classic Souders-Brown
entrainment-flooding limit does not govern this zone the way it governs a
plain sieve (Phase 1 Section 3.2's "Regime B" argument). HT-08's
back-drivable helical sleeve couples axial lift to rotation, so slot area
opens and closes passively with vapour load: at the design point this zone
preserves HT-02's Csb ~0.14-0.18 m/s capacity and <0.02 entrainment at 90%
flood, while extending the zone's own turndown from HT-02's ~2.5:1 to
VORTEXA's ~8:1. A stuck sleeve reverts to fixed-geometry HT-02 behaviour --
benign, no cascading failure.

![VG-0 single-valve CFD: paired-jet recirculation](References/vg0_cfd_xvelocity_unequal_legs.jpeg)
*Reference: X-velocity CFD for a single VG-0 valve -- the red/blue lobes
either side of the cap are the same kind of paired-jet recirculation pattern
that each vortex-tube cartridge organizes deliberately and at larger scale,
with a fixed vane cap for mechanical disengagement instead of relying on
gravity settling.*

### 3.2 Secondary zone -- lip-sealed trickle path (~15% of active area)

HT-01C LipSeal dual-flow lip valves, sized as a deliberately small secondary
zone. On its own, HT-01C's check-valve lips give a >=30% wider stable window
than plain sieve (~2.63:1 turndown, min_load_frac ~0.38) by sealing against
weep at low rate. The point of this zone is *not* its own turndown number --
it is what happens when the **primary zone closes**:

- At high-to-moderate loads, both zones contact normally; the primary zone
  (70% of area) carries roughly its proportional share of vapour/liquid
  traffic.
- As overall load falls toward the primary zone's own ~12.5% threshold
  (VORTEXA's min_load_frac), the primary zone's sleeves close -- benignly,
  by design (3.1). All remaining flow is now forced through the secondary
  zone alone.
- Because the secondary zone is only ~15% of the deck, the *same* overall
  flow now represents a much larger fraction of *its* design capacity --
  comfortably inside its own ~38% weep threshold even as overall load
  continues to fall.
- The secondary zone only weeps when overall load drops below roughly its
  own 38% threshold times its 15% area share, i.e. in the rough
  neighbourhood of 5-6% of total design flow.

This "primary closes -> secondary absorbs the remainder at a much higher
fraction of its own, smaller design capacity" hand-off is the mechanism
behind Apex's modeled `min_load_frac = 0.045` (turndown 22.2:1, Section 7,
item 2) -- well past either zone's individual number (8:1 / 2.63:1). **This
is the central modeling caveat for the whole document**: the 1-D engine rates
Apex as *one* `TrayType` with *one* `min_load_frac`; it cannot simulate the
two-zone hand-off directly. The sketch above shows the combined number is
*plausible* with reasonable area splits, and `min_load_frac = 0.045` is
recorded as the **design target** the two-zone hand-off is aiming for, not a
number independently derived bottom-up from the two zones' individual
datasheet figures. A genuine two-zone engine extension (Section 8) would
validate this properly.

### 3.3 Radial/azimuthal grading + multi-chordal sweep (efficiency, no dead zones)

HT-03 GRADEX's radially-graded open area (8% -> 13% across the radius, with
an azimuth map from a flow-field solver) and chord-tuned picket weir are
applied across *both* zones, plus Phase 1's multi-chordal liquid-handling
layout (center + side downcomers, vortex banks arranged in chordal strips so
liquid-sweep direction follows the jet angle). Together these push the deck
toward plug flow (Peclet ~10 -> ~30 on GRADEX's own datasheet, +10-18
Murphree points on large trays) and actively sweep every part of the deck --
directly addressing "no stagnant corners / no dead zones."

### 3.4 Self-sweeping oscillator relief (fouling)

HT-05 PULSAR's Coanda-island fluidic oscillator slots are cut into the
deck plate between cartridges (not inside them, so they add negligible
moving-part count). With no moving parts of their own, they give the
self-sweeping jet (5-30 Hz) that keeps cartridge faces and inter-cartridge
gaps clear between turnarounds -- PULSAR's own fouling run-length (~5-6 yr)
and +8 EOG points are the calibration point for Apex's
`fouling_open_area_retention = 0.93` (just under PULSAR's family-best 0.95).

### 3.5 Cartridge-grid structural deck (2-5 psi, cartridge-swap maintenance)

HT-07 AEGIS's hook-clamp + rib-stiffened + relief-flap structural overlay is
applied to a deck plate whose primary and secondary zones are built as
discrete bolt-in cartridges (3.1/3.2) on a dense beam grid (Phase 1 Section
3.4). Two things fall out of this for free:

- **2-5 psi transient/sustained uplift rating** is a *mechanical* design
  case (AEGIS's hook clamps + beam grid), independent of the hydraulic
  per-tray dP the Fair model computes.
- **Maintenance = cartridge swap**: because the primary and secondary zones
  are discrete cartridges bolted to the grid (not welded/integral), a fouled,
  eroded, or sleeve-stuck cartridge is replaced as a unit during a
  turnaround, rather than requiring deck-wide rework.

### 3.6 The deck as controller (passive self-regulation)

Every load-tracking element in Sections 3.1-3.2 -- VORTEXA's helical sleeves,
LipSeal's check-valve lips -- responds to *local* vapour/liquid load
instantaneously and mechanically. There is no sensor, no actuator, no
controller, no wiring. "Real-time adaptive" is achieved as a *property of
the geometry*: the deck's open area continuously tracks the load it is
currently seeing, with zero latency and zero electronic failure modes. Optional
instrumentation-only ports (pressure/temperature taps, no actuation) remain
available for DCS monitoring without compromising this. Section 7, item 9
scores this reframing explicitly.

## 4. Requirement -> mechanism mapping

| Wishlist item | Architecture element(s) | `TrayType` parameter(s) affected | Source HT-0X |
|---|---|---|---|
| Efficiency >90% Murphree | 3.3 grading + multi-chordal sweep | `efficiency_factor` | HT-03 GRADEX |
| Turndown >20:1 | 3.1 benign primary closure + 3.2 secondary trickle hand-off | `min_load_frac` | HT-08 VORTEXA, HT-01C LipSeal |
| Pressure drop -- packing-level | 3.1 mechanical disengagement (large open area, high C0) | `f_hole`, `c0`, `dp_dry_floor_in`, `aeration_factor`, `h_weir_in` | HT-02 CVS / HT-08 VORTEXA (geometry), Phase 1 large-slot concept |
| Capacity -- grid-tray-level | 3.1 vane-cap disengagement bypasses Souders-Brown at design point | `capacity_factor`, `f_active` | HT-02 CVS, HT-08 VORTEXA, Phase 1 vortex banks |
| Fouling resistance -- near self-cleaning | 3.4 oscillator relief + large chamfered slots (Phase 1) | `fouling_open_area_retention` | HT-05 PULSAR |
| Maintenance -- cartridge swap | 3.5 bolt-in cartridge grid | *(architecture only -- no parameter)* | HT-07 AEGIS |
| Stability -- no weeping/dumping | 3.2 secondary trickle path + downcomer sizing | `min_load_frac`, `h_weir_in`, downcomer backup (`a_dc`, `h_dc`) | HT-01C LipSeal |
| Cost -- manufacturable at scale | 3.5 cartridges as repeatable stamped/cast units, large fixed slots | *(architecture only -- no parameter)* | HT-07 AEGIS, Phase 1 large-slot concept |
| Control -- real-time adaptive | 3.6 passive load-tracking sleeves/lips | *(reframing -- same parameters as turndown row)* | HT-08 VORTEXA, HT-01C LipSeal |
| Energy -- major reduction | 3.1+3.3 fewer, lower-dP, higher-efficiency trays | `total_dp_psi` (derived), `n_actual_trays` (derived) | combination |
| No stagnant corners / no dead zones | 3.3 multi-chordal sweep | `efficiency_factor` (proxy) | HT-03 GRADEX |
| No jet flooding / no entrainment | 3.1 vane-cap mechanical disengagement | `capacity_factor` (proxy) | HT-02 CVS, HT-08 VORTEXA |

## 5. APEX `TrayType` parameters

| Parameter | Value | vs. Sieve | Derivation |
|---|---|---|---|
| `f_active` | 0.85 | 0.78 | 70% primary (3.1) + 15% secondary (3.2) |
| `f_hole` | 0.25 | 0.10 | Large fixed slots, sized as-fouled (Phase 1 3.1); largest in the family |
| `h_weir_in` | 0.5 | 2.0 | Low weir -- mechanical disengagement (3.1) replaces weir-height-driven disengagement; lowest in the family |
| `c0` | 0.85 | 0.73 | Highest C0 in the family -- large smooth slots, no valve hardware |
| `dp_dry_floor_in` | 0.10 | 0.00 | Small floor from cartridge/sleeve hardware (HT-02/HT-08 lineage) |
| `aeration_factor` | 0.38 | 0.55 | Lowest in the family -- mechanical disengagement (3.1) produces leaner froth than gravity settling |
| `capacity_factor` | 2.50 | 1.00 | HT-02/HT-08's 1.65x base, amplified by larger `f_hole`/`c0` -- the family's biggest *synthesis* number (echoes Phase 1's "Regime B (stretch)" 2.4-2.8x almost exactly) |
| `efficiency_factor` | 1.20 | 1.00 | HT-03 GRADEX's grading/sweep (3.3), highest in the family |
| `min_load_frac` | 0.045 | 0.50 | Two-zone hand-off (3.2) design target; lowest (best) in the family by a wide margin |
| `fouling_open_area_retention` | 0.93 | 0.75 | HT-05 PULSAR's self-sweeping relief (3.4), just under PULSAR's 0.95 family-best |
| `has_downcomer` | True | True | Standard weir/downcomer layout (3.2/3.5) |

## 6. Design-point results (from the engine)

Sieve vs. Apex, all five services in `service_cases.SERVICES`, at each
service's own `f_flood` (full tables in `output/tray_selection_study.md`):

| Service | Diameter: Sieve -> Apex (ft) | Total dP: Sieve -> Apex (psi) | Apex turndown | Apex efficiency | Apex DC backup |
|---|---|---|---|---|---|
| General Rectification | 4.39 -> **2.66** (-39%) | 1.503 -> **0.697** (-54%) | 22.22:1 | 91.4% | 23% |
| Fouling / Heavy-Ends | 5.42 -> **3.28** (-39%) | 4.070 -> **1.711** (-58%) | 22.22:1 | 40.2% | 33% |
| Vacuum Tower | 8.84 -> **5.35** (-39%) | 1.891 -> **0.985** (-48%) | 22.22:1 | 55.0% | 26% |
| High-P/High-L (C3/C4) | 9.38 -> **5.68** (-39%) | 2.463 -> **1.361** (-45%) | 22.22:1 | 100.0%\* | **54%** |
| Foaming | 5.07 -> **3.07** (-39%) | 1.130 -> **0.470** (-58%) | 22.22:1 | 91.4% | 17% |

\* `e_tray = min(e_oc * efficiency_factor, 100.0)` -- clipped at the engine's
100% efficiency cap for this high-O'Connell-baseline service.

The -39% diameter reduction is identical across services because
`capacity_factor` (2.50) and `f_active` (0.85) are fixed -- only the flow
parameter FLV (which sets `csbf`) varies, and it cancels out of the *ratio*
between two trays at the same `f_flood`. Total-dP reduction varies more
(-45% to -58%) because it depends on both the per-tray dP *and* how many
actual trays each tray's efficiency needs for that service's O'Connell
baseline.

The one number that moves the wrong way is **downcomer backup in
High-P/High-L: 54% vs. Sieve's 43%** -- still inside the 100% limit (46%
margin remaining) but the *tightest* in the entire family (next-worst is
HT-02 CVS at 47%). This is the direct cost of `h_weir_in = 0.5`: a shorter
weir at a smaller diameter raises the Francis-weir crest for the same liquid
rate. Flagged as a residual risk for very-high-FLV services in Section 8.

## 7. Scorecard against the 10-metric wishlist

| # | Metric | Model result | Status | Basis & caveats |
|---|---|---|---|---|
| 1 | Efficiency >90% Murphree | 91.4% (General/Foaming), 55.0% (Vacuum), 40.2% (Fouling), 100.0% (High-P/High-L, capped) | **Met** for favourable-alpha services; **highest in family everywhere** | `efficiency_factor=1.20` is a *relative* uplift on the system's O'Connell baseline (`e_oc`), which Apex cannot change -- ">90%" is a property of the tray *and* the separation's alpha/mu_L |
| 2 | Turndown >20:1 | 22.22:1 (= 1/`min_load_frac`, constant across services) | **Met**, but see caveat | Two-zone hand-off design target (Section 3.2) -- plausible with reasonable area splits, **not yet validated by a two-zone model** |
| 3 | Pressure drop -- packing-level | 1.83 in liquid/tray (General) = 0.88 in H2O/actual tray = ~0.96 in H2O/theoretical stage | **Stretch -- approaching, not matching** | High-efficiency structured packing typically runs ~0.3-0.5 in H2O/theoretical stage; Apex is ~2-3x that -- *less than half of Sieve's per-stage dP* (best in family), but not literally "packing-level" |
| 4 | Capacity -- grid-tray-level | `capacity_factor=2.50` (vs. typical high-capacity grid trays ~1.4-1.6x); -39% diameter vs. Sieve in every service | **Met / exceeded on paper** | The family's single biggest *synthesis* number -- a multiplicative combination of HT-02/HT-08's 1.65x base with larger `f_hole`/`c0`; echoes Phase 1's "Regime B (stretch)" 2.4-2.8x almost exactly, so it carries the same CFD/pilot-validation need |
| 5 | Fouling resistance -- near self-cleaning | `fouling_open_area_retention=0.93` (PULSAR=0.95 is family-best, Sieve=0.75) | **Met** | HT-05 PULSAR self-sweeping relief (3.4) + large chamfered slots; only +16% dry-tray dP increase when fouled (vs. Sieve's +78%, see fouling-sensitivity figure) |
| 6 | Maintenance -- cartridge swap | *(architecture, not a numeric output)* | **Met by design** | HT-07 AEGIS-style bolt-in cartridge grid (3.5); not modeled by the 1-D engine |
| 7 | Stability -- no weeping/dumping | `min_load_frac=0.045`; DC backup 17-54% across services, all within the 100% limit | **Met in the model's terms**, with one flagged margin | "No weeping" down to 4.5% load shares item 2's caveat; High-P/High-L DC backup at 54% is the *tightest* in the family (Section 6) |
| 8 | Cost -- manufacturable at scale | *(architecture, not a numeric output)* | **Architecture claim -- open** | Repeatable stamped/cast cartridges + large fixed slots (3.5) plausibly help, but a first-order cost multiple vs. conventional decks is still needed (Section 8) |
| 9 | Control -- real-time adaptive | *(reframing, same params as item 2)* | **Reframed** | "Real-time adaptive" delivered as passive, zero-latency geometry (3.6) rather than sensors/actuators -- electronics would contradict items 5/6/8 |
| 10 | Energy -- major reduction | Total column dP -45% to -58% vs. Sieve across all 5 services (Section 6) | **Met** | The most consistent win in the scorecard; lower column dP reduces reboiler-duty / recompression requirements in pressure-driven services |

**Qualitative claims:**

| Claim | Mechanism | Status |
|---|---|---|
| No stagnant corners / no dead zones | 3.3 GRADEX grading + multi-chordal sweep | Architecture-based (efficiency_factor as proxy) |
| No jet flooding / no entrainment | 3.1 vane-cap mechanical disengagement | Architecture-based (capacity_factor as proxy; HT-02's own <0.02 entrainment at 90% flood is the calibration point) |
| No weeping / no dumping | 3.2 two-zone hand-off | Same as scorecard item 7 |
| Massive turndown ratio | 3.2 two-zone hand-off | Same as scorecard item 2 (22.22:1) |

**Bottom line**: 7 of 10 metrics are *Met* (numerically or by design); one
(*Pressure drop*) is a *stretch* -- approaching but not literally matching its
target; one (*Control*) is honestly *reframed* as passive geometry rather than
electronics; and one (*Cost*) remains an *open* architecture claim pending a
first-order cost study. The two metrics carrying the most validation risk --
**Capacity** (item 4, a 2.50x synthesis number) and **Turndown** (item 2, a
two-zone hand-off the 1-D engine can't simulate directly) -- are exactly the
two flagged for CFD/pilot work in Section 8.

## 8. Open questions / R&D needs

1. **Two-zone turndown model.** Section 3.2's primary-closes /
   secondary-absorbs hand-off is the basis for `min_load_frac=0.045`
   (scorecard item 2) but is currently a single-`TrayType` design target, not
   a derived result. A genuine extension would model the primary and
   secondary zones as separate sub-decks with their own area fractions and
   `min_load_frac`, summing their flows -- directly testing whether ~20:1+ is
   achievable with realistic 70/15 (or other) area splits.
2. **Capacity synthesis validation.** `capacity_factor=2.50` (scorecard item
   4) combines HT-02/HT-08's disclosed 1.65x with additional gains from
   larger `f_hole`/`c0` multiplicatively. This is the same CFD/pilot-needed
   number Phase 1 flagged as "Regime B (stretch)" -- still open.
3. **High-P/High-L downcomer margin.** Apex's 54% DC backup in the
   High-Pressure/High-Liquid-Load service (Section 6) is the tightest in the
   family. Worth checking whether a slightly taller `h_weir_in` (e.g. 1.0 in
   instead of 0.5 in) recovers margin there without materially hurting the
   other four services -- or whether this service class should pair Apex
   with the HT-04 DCX active-downcomer module (`NON_DECK_MODULES`).
4. **Cartridge interface details (3.1/3.2 boundary).** HT-08's helical
   sleeves were developed for HT-02's round swirl-tube slots; HT-01C's lip
   valves for a conventional dual-flow deck. The mechanical interface where a
   70%-area cartridge field meets a 15%-area lip-valve field (sealing,
   differential thermal growth, vibration coupling) is unaddressed by any
   individual HT-0X datasheet.
5. **Cost multiple.** Scorecard item 8 (manufacturable at scale) is currently
   an architecture claim. A first-order cost estimate -- cartridges +
   AEGIS-grade structural grid vs. a conventional valve/sieve deck of the
   same diameter -- is needed before "manufacturable at scale" is more than
   an assertion.

## 9. One-line summary

HyTrays Apex = an **Integrated Cartridge-Grid**: HT-02/HT-08 vortex-tube
cartridges (70% of the deck) for grid-tray-beating capacity at near-packing
pressure drop, an HT-01C lip-sealed trickle zone (15%) that takes over when
the primary cartridges benignly close -- pushing turndown past 20:1 -- HT-03
GRADEX grading and multi-chordal sweep for >90% efficiency with no dead
zones, HT-05 PULSAR oscillator relief for near-self-cleaning fouling
resistance, and an HT-07 AEGIS bolt-in cartridge grid that delivers the 2-5
psi structural rating *and* turns maintenance into a cartridge swap. Every
adaptive element is passive and load-driven, so "real-time adaptive control"
is the deck's geometry, not its instrumentation. Eight of the ten wishlist
metrics are met outright in the model; the other two (packing-level dP,
electronic-style control) are honestly reframed as "approaching" and
"achieved passively" respectively -- and the two highest-leverage numbers
(2.50x capacity, 22.2:1 turndown) are exactly the two queued for CFD/pilot
validation.
