#!/usr/bin/env python3
"""Composition reader for PRO/II "Case N" transposed HMB exports.

Thin adapter over ``hmb_proii_reader.parse_case`` exposing exactly the
interface ``hydraulics_XOM._from_casesheet`` expects:

    parse_proii(path, case) -> (streams, comp_data, names)

where ``comp_data[stream]`` is a dict with the per-component maps used by
``_build_feed``:
    MASS_FLOW, MOLE_FLOW, MOLE_FRAC (total z),
    LIQ_MOLE_FRAC (x), VAP_MOLE_FRAC (y)
and ``names`` is the component list in feed order.
"""
from __future__ import annotations

import hmb_proii_reader as _R

_KEYS = ("MASS_FLOW", "MOLE_FLOW", "MOLE_FRAC", "LIQ_MOLE_FRAC", "VAP_MOLE_FRAC")


def parse_proii(path, case="Case 1"):
    data = _R.parse_case(path, case)
    comp = data["comp"]
    comp_data: dict[str, dict] = {}
    for stream in data["streams"]:
        comp_data[stream] = {k: comp.get(k, {}).get(stream, {}) for k in _KEYS}
    return data["streams"], comp_data, data["components"]
