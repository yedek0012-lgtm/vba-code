#!/usr/bin/env bash
# Görünüş yeri bulma testi (AutoCAD'siz).  bash test/gorunus_konum_testi.sh eklenti/bin/tek/OutlineDrawing.dll kule.tow
set -euo pipefail
DLL="$(realpath "$1")"; TOW="$(realpath "$2")"
cd "$(dirname "$0")"
source ./ortak.sh; hazirla gorunus_konum_testi GorunusKonumTest.cs GKT.exe
cd "$OUT" && mono GKT.exe "$TOW" 2>&1 | grep -v "custom attr"
