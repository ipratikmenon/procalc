#!/usr/bin/env python3
"""HyTrays — service-specific column cases for the tray selection study.

Each :class:`column_hydraulics.ColumnCase` below represents a distinct
*service* commonly used as a tray-selection discriminator (Kister,
*Distillation Operation*, Ch. 1): general hydrocarbon rectification,
fouling/heavy-ends, vacuum, high-pressure/high-liquid-load and foaming.

Only the "General Rectification" case is anchored to a real stream
(T801-OH / T801-REFLUX in ``Hydraulics/HMB.xlsx``); the others use
representative property sets for the named service so that the five tray
types in ``tray_library.py`` can be compared on a common, reproducible
basis. Edit the numbers below to match a specific column once real data is
available -- the run script recomputes everything from these cases.
"""
from __future__ import annotations

from column_hydraulics import ColumnCase

GENERAL = ColumnCase(
    name="General Rectification (T801 basis)",
    description="Mid-pressure hydrocarbon rectifying section -- vapor/liquid "
                 "loads and vapor density from T801-OH / T801-REFLUX "
                 "(Hydraulics/HMB.xlsx); liquid density, viscosity, surface "
                 "tension and alpha are estimates for ~344 F / 105 psia.",
    v_mass_lb_hr=52_448.77, l_mass_lb_hr=14_797.57,
    rho_v=1.018, rho_l=30.0, sigma=9.0, mu_l=0.12, alpha=1.40,
    tray_spacing_in=24.0, f_flood=0.80, n_theoretical=20,
)

FOULING = ColumnCase(
    name="Fouling / Heavy-Ends Service",
    description="Heavy, viscous, deposit-forming liquid (e.g. coker main "
                 "fractionator wash/slurry section): high liquid load, "
                 "high viscosity, close relative volatility. Sized at a "
                 "reduced 65% of flood to leave margin for deposit buildup; "
                 "evaluated with tray.fouling_open_area_retention applied "
                 "(fouling=True).",
    v_mass_lb_hr=45_000.0, l_mass_lb_hr=90_000.0,
    rho_v=0.40, rho_l=42.0, sigma=18.0, mu_l=3.00, alpha=1.15,
    tray_spacing_in=24.0, f_flood=0.65, n_theoretical=15,
    fouling=True,
)

VACUUM = ColumnCase(
    name="Vacuum Tower Section",
    description="Low-pressure HVGO-type section: very low vapor density "
                 "drives large diameters and high superficial velocities; "
                 "every inch of tray dP raises the flash-zone temperature, "
                 "so dP/tray is the key discriminator. 30 in. tray spacing "
                 "(common practice in vacuum towers).",
    v_mass_lb_hr=80_000.0, l_mass_lb_hr=40_000.0,
    rho_v=0.06, rho_l=38.0, sigma=14.0, mu_l=1.20, alpha=1.20,
    tray_spacing_in=30.0, f_flood=0.80, n_theoretical=8,
)

HIGH_P_HIGH_L = ColumnCase(
    name="High-Pressure / High-Liquid-Load Service (C3/C4 splitter)",
    description="High-pressure light-ends splitter: high reflux ratio "
                 "(L/V ~ 2), low surface tension, low relative volatility "
                 "drives many stages. High FLV stresses downcomer/liquid "
                 "handling rather than vapor capacity.",
    v_mass_lb_hr=150_000.0, l_mass_lb_hr=300_000.0,
    rho_v=2.00, rho_l=28.0, sigma=5.0, mu_l=0.08, alpha=1.15,
    tray_spacing_in=24.0, f_flood=0.80, n_theoretical=40,
)

FOAMING = ColumnCase(
    name="Foaming Service",
    description="Same traffic/properties as General Rectification but with "
                 "a moderate foaming system factor (Fp = 0.75, Kister "
                 "Table 1-2) applied uniformly to every tray's flooding "
                 "velocity -- representative of amine/glycol-type foaming "
                 "tendencies.",
    v_mass_lb_hr=52_448.77, l_mass_lb_hr=14_797.57,
    rho_v=1.018, rho_l=30.0, sigma=9.0, mu_l=0.12, alpha=1.40,
    tray_spacing_in=24.0, f_flood=0.80, n_theoretical=20,
    system_factor=0.75,
)

SERVICES: list[ColumnCase] = [GENERAL, FOULING, VACUUM, HIGH_P_HIGH_L, FOAMING]
