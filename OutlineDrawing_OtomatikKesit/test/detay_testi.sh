#!/usr/bin/env bash
# Uçtan uca detay testi (AutoCAD'siz): okuma, otomatik kesit + gerçek SectionDetection, görünüşler; atlanan eleman var mı?
#   bash test/detay_testi.sh eklenti/bin/tek/OutlineDrawing.dll kule.tow [ayrinti]
set -euo pipefail
DLL="$(realpath "$1")"; TOW="$(realpath "$2")"; AYRINTI="${3:-}"
cd "$(dirname "$0")"
OUT="${TMPDIR:-/tmp}/detay_testi"; mkdir -p "$OUT"
if [ ! -f "$OUT/Acmgd.dll" ]; then
  mcs -target:library -out:"$OUT/Acdbmgd.dll" sahte_autocad/Acdbmgd.cs
  mcs -target:library -out:"$OUT/accoremgd.dll" sahte_autocad/accoremgd.cs
  mcs -target:library -r:"$OUT/accoremgd.dll" -out:"$OUT/Acmgd.dll" sahte_autocad/Acmgd.cs
fi
cp "$DLL" "$OUT/OutlineDrawing.dll"
mcs -r:"$OUT/OutlineDrawing.dll" -out:"$OUT/DT.exe" DetayTesti.cs 2>&1 | grep -v "custom attr" | grep -v "^Compilation succeeded" || true
cd "$OUT" && mono DT.exe "$TOW" $AYRINTI 2>&1 | grep -v "custom attr"
