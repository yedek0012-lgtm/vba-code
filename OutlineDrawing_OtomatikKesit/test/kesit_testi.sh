#!/usr/bin/env bash
# OutlineDrawing.dll'in (tek DLL) gerçek SectionDetection'ını AutoCAD olmadan çalıştırır.
#   bash test/kesit_testi.sh eklenti/bin/tek/OutlineDrawing.dll kule.tow
# Sahte AutoCAD kütüphaneleri (test/sahte_autocad) yalnız SectionDetection'ın kullandığı geometri ve
# Application tiplerini içerir; gerçek AutoCAD ile aynı matematik.
set -euo pipefail
DLL="$(realpath "$1")"; TOW="$(realpath "$2")"
cd "$(dirname "$0")"
source ./ortak.sh; hazirla kesit_testi SectionDetectionTest.cs SD.exe
IS=$(mktemp -d "${TMPDIR:-/tmp}/kesit_is.XXXXXX"); trap 'rm -rf "$IS"' EXIT; cp "$TOW" "$IS/kule.tow"
cd "$OUT" && mono SD.exe "$IS/kule.tow" 2>&1 | grep -v "custom attr"
