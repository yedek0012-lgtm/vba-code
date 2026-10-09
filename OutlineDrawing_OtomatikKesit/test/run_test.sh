#!/usr/bin/env bash
# Mono ile derleme + sentetik kule testi (AutoCAD gerekmez).
#   sudo apt-get install mono-mcs mono-runtime libmono-system-windows-forms4.0-cil
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
