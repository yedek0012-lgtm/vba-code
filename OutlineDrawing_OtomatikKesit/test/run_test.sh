#!/usr/bin/env bash
# Mono ile derleme + testler (AutoCAD gerekmez).
#   sudo apt-get install mono-mcs mono-runtime libmono-system-windows-forms4.0-cil
#
#   bash test/run_test.sh                 -> derleme kontrolu + sentetik kule testi
#   bash test/run_test.sh kule.tow        -> + gercek .tow uzerinde onerilen kesitler
#   bash test/run_test.sh kule.tow planes -> egik duzlem aramasi da acik
set -euo pipefail
cd "$(dirname "$0")/.."
OUT="${TMPDIR:-/tmp}/autosection_test"
mkdir -p "$OUT"

# 1) Buton kodu derleniyor mu? (mainform / dataprocessing imzalari stub)
mcs -langversion:7.2 -target:library -r:System.Windows.Forms.dll -out:"$OUT/form_check.dll" \
    AutoSectionDetector.cs mainform.OtomatikKesit.cs test/Stubs.cs test/FormStubs.cs

# 2) Algoritma testi
mcs -langversion:7.2 -out:"$OUT/autosection_test.exe" \
    AutoSectionDetector.cs test/Stubs.cs test/SyntheticTowerTest.cs
mono "$OUT/autosection_test.exe"

# 3) Gercek .tow (istege bagli)
if [ $# -ge 1 ]; then
    python3 test/tow_to_csv.py "$1" > "$OUT/kule.csv"
    mcs -langversion:7.2 -out:"$OUT/tow_run.exe" AutoSectionDetector.cs test/Stubs.cs test/TowCsvRun.cs
    mono "$OUT/tow_run.exe" "$OUT/kule.csv" "${2:-}"
fi
