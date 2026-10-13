#!/usr/bin/env bash
# Uçtan uca detay testi (AutoCAD'siz): okuma, otomatik kesit + gerçek SectionDetection, görünüşler; atlanan eleman var mı?
#   bash test/detay_testi.sh eklenti/bin/tek/OutlineDrawing.dll kule.tow [ayrinti]
set -euo pipefail
DLL="$(realpath "$1")"; TOW="$(realpath "$2")"; AYRINTI="${3:-}"
cd "$(dirname "$0")"
source ./ortak.sh; hazirla detay_testi DetayTesti.cs DT.exe
cd "$OUT" && mono DT.exe "$TOW" $AYRINTI 2>&1 | grep -v "custom attr"
