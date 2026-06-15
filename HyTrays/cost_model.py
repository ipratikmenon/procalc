#!/usr/bin/env python3
"""HyTrays — first-order relative installed-cost model.

A column's installed cost for tray internals splits, very roughly, into two
pieces that scale differently with tray-type:

* **Shell + heads + supports** -- the dominant cost driver for most tray
  columns (Towler & Sinnott, *Chemical Engineering Design*, give purchased
  vessel cost as the majority of a tray column's bare-module cost, with
  trays/internals typically ~15-30% of the total). For a *fixed pressure
  class* (constant wall thickness), shell metal weight -- and so cost --
  scales roughly with the cylindrical shell's surface area, ``D * H``.
* **Tray hardware** -- punched/stamped/cast deck panels, valves, downcomers
  or cartridges. Cost per tray scales with deck area (``D**2``); the
  *fabrication complexity* multiplier is ``tray.relative_unit_cost_factor``
  (see ``tray_library.TrayType``); and the number of trays needed is
  ``n_actual_trays``.

``relative_installed_cost`` combines these two proxies with a fixed split
(``SHELL_COST_FRACTION``) and normalises everything to the Sieve baseline
(=1.00). This is a *first-order, illustrative* estimate -- it answers
``HyTrays_Apex_concept.md`` Section 8's open cost-multiple question with a
directional result, not a quote. Real costing needs vendor pricing for the
cartridge hardware and a proper vessel-costing correlation.
"""
from __future__ import annotations

from dataclasses import dataclass

from column_hydraulics import TrayDesignResult

# Fraction of installed cost attributed to the shell/heads/supports, with the
# remainder to tray hardware. 70/30 is a representative split for a tray
# column per Towler & Sinnott-style guidance; the result is not very
# sensitive to this choice because the shell and tray proxies move together
# for most trays (see ``HyTrays_Apex_Technical_Paper.md`` Section 7).
SHELL_COST_FRACTION = 0.70


@dataclass(frozen=True)
class CostResult:
    """Relative installed-cost breakdown for one ``TrayDesignResult``."""

    shell_proxy: float    # D * H            (relative shell metal/area)
    tray_proxy: float      # D^2 * N * k_cost  (relative tray hardware)
    shell_ratio: float      # shell_proxy / sieve shell_proxy
    tray_ratio: float        # tray_proxy / sieve tray_proxy
    relative_installed_cost: float  # weighted combination, sieve = 1.00


def _proxies(result: TrayDesignResult) -> tuple[float, float]:
    shell_proxy = result.diameter_ft * result.column_height_ft
    tray_proxy = (result.diameter_ft ** 2) * result.n_actual_trays \
        * result.tray.relative_unit_cost_factor
    return shell_proxy, tray_proxy


def relative_installed_cost(
    result: TrayDesignResult,
    sieve_result: TrayDesignResult,
    shell_fraction: float = SHELL_COST_FRACTION,
) -> CostResult:
    """Relative installed cost of ``result`` vs. ``sieve_result`` (=1.00)."""
    shell, tray = _proxies(result)
    shell_s, tray_s = _proxies(sieve_result)
    shell_ratio = shell / shell_s
    tray_ratio = tray / tray_s
    total = shell_fraction * shell_ratio + (1.0 - shell_fraction) * tray_ratio
    return CostResult(
        shell_proxy=shell, tray_proxy=tray,
        shell_ratio=shell_ratio, tray_ratio=tray_ratio,
        relative_installed_cost=total,
    )
