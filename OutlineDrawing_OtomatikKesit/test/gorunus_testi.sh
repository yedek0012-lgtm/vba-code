#!/usr/bin/env bash
# Ön / yan görünüşte PLS-TOWER'da görünen her çizginin çizimde de olduğunu AutoCAD olmadan kontrol eder.
#   bash test/gorunus_testi.sh eklenti/bin/tek/OutlineDrawing.dll kule.tow [hepsi]
set -euo pipefail
DLL="$(realpath "$1")"; TOW="$(realpath "$2")"; AYRINTI="${3:-}"
cd "$(dirname "$0")"
source ./ortak.sh; hazirla gorunus_testi GorunusKontrol.cs GK.exe
cd "$OUT" && mono GK.exe "$TOW" $AYRINTI 2>&1 | grep -v "custom attr"
