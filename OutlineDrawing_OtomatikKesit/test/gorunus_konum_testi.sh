#!/usr/bin/env bash
# Görünüş yeri bulma testi (AutoCAD'siz).  bash test/gorunus_konum_testi.sh eklenti/bin/tek/OutlineDrawing.dll kule.tow
set -euo pipefail
DLL="$(realpath "$1")"; TOW="$(realpath "$2")"
cd "$(dirname "$0")"
OUT="${TMPDIR:-/tmp}/gorunus_konum_testi"; rm -rf "$OUT"; mkdir -p "$OUT"
mcs -target:library -out:"$OUT/Acdbmgd.dll" sahte_autocad/Acdbmgd.cs
mcs -target:library -out:"$OUT/accoremgd.dll" sahte_autocad/accoremgd.cs
mcs -target:library -r:"$OUT/accoremgd.dll" -out:"$OUT/Acmgd.dll" sahte_autocad/Acmgd.cs
cp "$DLL" "$OUT/OutlineDrawing.dll"
mcs -r:"$OUT/OutlineDrawing.dll" -out:"$OUT/GKT.exe" GorunusKonumTest.cs 2>&1 | grep -v "custom attr" || true
cd "$OUT" && mono GKT.exe "$TOW" 2>&1 | grep -v "custom attr"
