#!/usr/bin/env bash
# OutlineDrawing.dll'in (tek DLL) gerçek SectionDetection'ını AutoCAD olmadan çalıştırır.
#   bash test/kesit_testi.sh eklenti/bin/tek/OutlineDrawing.dll kule.tow
# Sahte AutoCAD kütüphaneleri (test/sahte_autocad) yalnız SectionDetection'ın kullandığı geometri ve
# Application tiplerini içerir; gerçek AutoCAD ile aynı matematik.
set -euo pipefail
DLL="$(realpath "$1")"; TOW="$(realpath "$2")"
cd "$(dirname "$0")"
OUT="${TMPDIR:-/tmp}/kesit_testi"; rm -rf "$OUT"; mkdir -p "$OUT"
mcs -target:library -out:"$OUT/Acdbmgd.dll" sahte_autocad/Acdbmgd.cs
mcs -target:library -out:"$OUT/accoremgd.dll" sahte_autocad/accoremgd.cs
mcs -target:library -r:"$OUT/accoremgd.dll" -out:"$OUT/Acmgd.dll" sahte_autocad/Acmgd.cs
cp "$DLL" "$OUT/OutlineDrawing.dll"; cp "$TOW" "$OUT/kule.tow"
mcs -r:"$OUT/OutlineDrawing.dll" -out:"$OUT/SD.exe" SectionDetectionTest.cs 2>&1 | grep -v "custom attr" || true
cd "$OUT" && mono SD.exe kule.tow 2>&1 | grep -v "custom attr"
