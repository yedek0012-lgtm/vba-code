#!/usr/bin/env bash
# Ön / yan görünüşte PLS-TOWER'da görünen her çizginin çizimde de olduğunu AutoCAD olmadan kontrol eder.
#   bash test/gorunus_testi.sh eklenti/bin/tek/OutlineDrawing.dll kule.tow [hepsi]
set -euo pipefail
DLL="$(realpath "$1")"; TOW="$(realpath "$2")"; AYRINTI="${3:-}"
cd "$(dirname "$0")"
OUT="${TMPDIR:-/tmp}/gorunus_testi"; rm -rf "$OUT"; mkdir -p "$OUT"
mcs -target:library -out:"$OUT/Acdbmgd.dll" sahte_autocad/Acdbmgd.cs
mcs -target:library -out:"$OUT/accoremgd.dll" sahte_autocad/accoremgd.cs
mcs -target:library -r:"$OUT/accoremgd.dll" -out:"$OUT/Acmgd.dll" sahte_autocad/Acmgd.cs
cp "$DLL" "$OUT/OutlineDrawing.dll"
mcs -r:"$OUT/OutlineDrawing.dll" -out:"$OUT/GK.exe" GorunusKontrol.cs 2>&1 | grep -v "custom attr" || true
cd "$OUT" && mono GK.exe "$TOW" $AYRINTI 2>&1 | grep -v "custom attr"
