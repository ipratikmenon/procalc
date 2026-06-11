#!/usr/bin/env bash
# macOS: double-click this file in Finder to open the hydraulics launcher.
# (It runs the interactive text menu. For the GUI, run: ./run_hydraulics.command gui)
cd "$(dirname "$0")" || exit 1

PY="$(command -v python3 || command -v python)"
if [ -z "$PY" ]; then
  echo "Python 3 was not found. Install it from https://www.python.org/downloads/"
  read -r -p "Press Enter to close..."
  exit 1
fi

if [ "$1" = "gui" ]; then
  "$PY" run_hydraulics.py --gui
else
  "$PY" run_hydraulics.py
fi

echo
read -r -p "Press Enter to close..."
