#!/usr/bin/env bash
# Ölçü yazısı çakışma testi (bkz. OlcuCakismaTesti.cs)
#   bash test/olcu_testi.sh eklenti/bin/tek/OutlineDrawing.dll kule.tow [ayrinti]
set -euo pipefail
DLL="$(realpath "$1")"; TOW="$(realpath "$2")"; AYRINTI="${3:-}"
cd "$(dirname "$0")"
source ./ortak.sh; hazirla olcu_testi OlcuCakismaTesti.cs OCT.exe
cd "$OUT" && mono OCT.exe "$TOW" $AYRINTI 2>&1 | grep -v "custom attr"
