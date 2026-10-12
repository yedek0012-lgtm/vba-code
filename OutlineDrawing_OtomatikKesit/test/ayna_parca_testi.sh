#!/usr/bin/env bash
# Sentetik ayna parçası testi (bkz. AynaParcaTesti.cs)
#   bash test/ayna_parca_testi.sh eklenti/bin/tek/OutlineDrawing.dll
set -euo pipefail
DLL="$(realpath "$1")"
cd "$(dirname "$0")"
OUT="${TMPDIR:-/tmp}/ayna_parca_testi"; mkdir -p "$OUT"
if [ ! -f "$OUT/Acmgd.dll" ]; then
  mcs -target:library -out:"$OUT/Acdbmgd.dll" sahte_autocad/Acdbmgd.cs
  mcs -target:library -out:"$OUT/accoremgd.dll" sahte_autocad/accoremgd.cs
  mcs -target:library -r:"$OUT/accoremgd.dll" -out:"$OUT/Acmgd.dll" sahte_autocad/Acmgd.cs
fi
cp "$DLL" "$OUT/OutlineDrawing.dll"
mcs -r:"$OUT/OutlineDrawing.dll" -out:"$OUT/APT.exe" AynaParcaTesti.cs 2>&1 | grep -v "custom attr" | grep -v "^Compilation succeeded" || true
cd "$OUT" && mono APT.exe 2>&1 | grep -v "custom attr"
