# -*- coding: utf-8 -*-
# =============================================================================
# PyInstaller runtime hook  --  Procalc Hydraulics
#
# Purpose
# -------
# The Procalc app package (``procalc_app``) imports the hydraulics engine that
# lives OUTSIDE the package -- the modules from ``Hydraulics/`` and ``common/``
# in the source tree (hydraulics_XOM, pms_classes, units, style, ...).  At
# source-run time ``procalc_app/engine_api.py`` and ``__main__.py`` put those
# directories on ``sys.path`` themselves.
#
# Under a PyInstaller *freeze* those engine modules are collected into the
# frozen import archive (the PYZ) as top-level modules, so ``import
# hydraulics_XOM`` already resolves without any path juggling.  This hook is a
# defensive belt-and-suspenders layer for the case where any engine module (or
# a data file it looks for next to itself) was instead shipped as a *data*
# file inside a sub-folder of the one-dir bundle -- e.g. ``Hydraulics/`` or
# ``common/`` under ``sys._MEIPASS``.  It prepends every plausible engine
# location to ``sys.path`` so those imports and companion-file lookups resolve
# no matter which layout the .spec produced.
#
# Runtime hooks run *before* any of the app's own code (including
# ``engine_api``'s own ``sys.path.insert`` calls), so the paths are in place by
# the time the first ``import hydraulics_XOM`` happens.
#
# This file must stay import-safe and ``python -m py_compile`` clean: it is
# executed verbatim at process start, so it must never raise.
# =============================================================================

import os
import sys


def _add_path(path):
    """Prepend *path* to sys.path if it is a real directory not already there."""
    try:
        if path and os.path.isdir(path) and path not in sys.path:
            sys.path.insert(0, path)
    except Exception:
        # Never let a bad path entry abort startup.
        pass


def _install_engine_paths():
    # ``sys._MEIPASS`` is defined only in a frozen (PyInstaller) process and
    # points at the extraction root of the one-dir/one-file bundle.
    meipass = getattr(sys, "_MEIPASS", None)
    if not meipass:
        # Not frozen: running from source -- the app inserts its own paths.
        return

    # Candidate roots that may hold the engine .py modules and their data
    # files (gems_extracted.json, ten_logo.png, ...).  Order matters only in
    # that more specific sub-folders come first; duplicates are ignored.
    candidates = [
        meipass,                                   # engine modules at bundle root
        os.path.join(meipass, "Hydraulics"),       # engine bundled under Hydraulics/
        os.path.join(meipass, "common"),           # units.py / style.py under common/
        os.path.join(meipass, "procalc_app"),      # app package, if nested
    ]
    for path in candidates:
        _add_path(path)


# Execute at import (which is when PyInstaller runs the hook).
_install_engine_paths()
