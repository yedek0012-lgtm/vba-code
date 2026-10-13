#!/usr/bin/env bash
# Sentetik ayna parçası testi (bkz. AynaParcaTesti.cs)
#   bash test/ayna_parca_testi.sh eklenti/bin/tek/OutlineDrawing.dll
set -euo pipefail
DLL="$(realpath "$1")"
cd "$(dirname "$0")"
source ./ortak.sh; hazirla ayna_parca_testi AynaParcaTesti.cs APT.exe
cd "$OUT" && mono APT.exe 2>&1 | grep -v "custom attr"
