#!/usr/bin/env bash
# Build out/C30_Building_Dashboard.xlsx with cached values: generate -> recalculate a copy in LibreOffice -> copy values back.
set -euo pipefail
cd "$(dirname "$0")"
RECALC=${RECALC:-/mnt/skills/public/xlsx/scripts/recalc.py}
TMP=$(mktemp -d)
python3 excel.py raw "$TMP/gen.xlsx"
cp "$TMP/gen.xlsx" "$TMP/calc.xlsx"
python3 "$RECALC" "$TMP/calc.xlsx" 300 | tee "$TMP/recalc.json"
grep -q '"total_errors": 0' "$TMP/recalc.json" || { echo "formula errors found"; exit 1; }
python3 inject_cache.py "$TMP/gen.xlsx" "$TMP/calc.xlsx" out/C30_Building_Dashboard.xlsx
cp "$TMP/calc.xlsx" out/.xlsx_recalc_check.xlsx
