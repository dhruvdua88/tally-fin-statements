#!/bin/bash
# ─────────────────────────────────────────────────────────────────────────────
#  Tally Financial Statements Generator — Mac Launcher
#  Double-click this file to install dependencies and open the app.
#
#  First time only: right-click → Open → click Open (macOS security dialog)
#  After that you can double-click directly.
# ─────────────────────────────────────────────────────────────────────────────

cd "$(dirname "$0")"

clear
echo "============================================================"
echo "  Tally Financial Statements Generator"
echo "  Schedule III Balance Sheet · P&L · 3-Year Projections"
echo "============================================================"
echo

# ── 1. Find Python 3.11+ ─────────────────────────────────────────────────────
PY=""
for candidate in python3 python3.13 python3.12 python3.11; do
    if command -v "$candidate" &>/dev/null; then
        ver=$("$candidate" -c "import sys; print(sys.version_info.major * 100 + sys.version_info.minor)")
        if [ "$ver" -ge 311 ]; then
            PY="$candidate"
            break
        fi
    fi
done

if [ -z "$PY" ]; then
    echo "X  Python 3.11 or later is not installed."
    echo
    echo "   Download it free from:"
    echo "   https://www.python.org/downloads/macos/"
    echo
    open "https://www.python.org/downloads/macos/"
    read -r -p "Press Return to close..."
    exit 1
fi

echo "OK  $("$PY" --version 2>&1) found."
echo

# ── 2. Create / reuse virtual environment ─────────────────────────────────────
VENV=".venv"
if [ ! -d "$VENV" ]; then
    echo "..  Creating virtual environment (first-time, ~10 seconds)..."
    "$PY" -m venv "$VENV"
    if [ $? -ne 0 ]; then
        echo "X  Could not create virtual environment."
        read -r -p "Press Return to close..."; exit 1
    fi
    echo "OK  Virtual environment created."
else
    echo "OK  Virtual environment ready."
fi
echo

# ── 3. Install / upgrade dependencies ────────────────────────────────────────
echo "..  Installing packages (openpyxl)..."
"$VENV/bin/pip" install --upgrade pip --quiet
"$VENV/bin/pip" install --upgrade -r requirements.txt --quiet
if [ $? -ne 0 ]; then
    echo "X  Package install failed. Check your internet connection."
    read -r -p "Press Return to close..."; exit 1
fi
echo "OK  Packages ready."
echo

# ── 4. Launch ────────────────────────────────────────────────────────────────
echo ">>  Starting app — this window can be closed once it opens."
echo
"$VENV/bin/python" financial_statements.py

if [ $? -ne 0 ]; then
    echo
    echo "X  App exited with an error (see above)."
    read -r -p "Press Return to close..."
fi
