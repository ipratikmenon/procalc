#!/usr/bin/env bash
# Linux/macOS terminal launcher.
#   ./run_hydraulics.sh        → interactive text menu
#   ./run_hydraulics.sh gui    → small Tkinter GUI
cd "$(dirname "$0")" || exit 1

PY="$(command -v python3 || command -v python)"
if [ -z "$PY" ]; then
  echo "Python 3 not found on PATH."
  exit 1
fi

if [ "$1" = "gui" ]; then
  exec "$PY" run_hydraulics.py --gui
else
  exec "$PY" run_hydraulics.py
fi
