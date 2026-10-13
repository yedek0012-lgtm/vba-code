#!/usr/bin/env bash
# hepsi.sh'nin bir dosya için çalıştırdığı testler; sonuç satırı (ve hata varsa ayrıntı) stdout'a.
#   bash test/tek_dosya.sh <dll> <kule.tow>
set -uo pipefail
DLL="$(realpath "$1")"; f="$(realpath "$2")"
cd "$(dirname "$0")"
d=$(bash detay_testi.sh "$DLL" "$f" 2>&1); k=$(bash kesit_testi.sh "$DLL" "$f" 2>&1)
g=$(bash gorunus_testi.sh "$DLL" "$f" 2>&1); c=$(bash gorunus_konum_testi.sh "$DLL" "$f" 2>&1)
o=$(bash olcu_testi.sh "$DLL" "$f" 2>&1)
sd=$(echo "$d" | grep -q "SONUC: atlanan eleman yok" && echo TAMAM || echo HATA)
sk=$(echo "$k" | grep -q "SONUC: butun kesitler tam" && echo "$k" | grep -q "eslesmeyen 0, yanlis stil 0" && echo TAMAM || echo HATA)
# görünüş: ön ve yan görünüşün ikisinde de her elemanın izdüşümü çizimde
sg=$( [ "$(echo "$g" | grep -cE "(on|yan): izdusum [0-9]+, cizimde var [0-9]+, yok 0")" = 2 ] && echo TAMAM || echo HATA)
sc=$(echo "$c" | grep -q "SONUC: hicbir gorunus yanlis yere konmadi" && echo TAMAM || echo HATA)
so=$(echo "$o" | grep -q "SONUC: olcu yazilari temiz" && echo TAMAM || echo HATA)
printf "%-58s %-8s %-8s %-8s %-8s %-8s\n" "$(basename "$f" | cut -c1-58)" "$sd" "$sk" "$sg" "$sc" "$so"
if [ "$sd$sk$sg$sc$so" != TAMAMTAMAMTAMAMTAMAMTAMAM ]; then echo "$d"; echo "$k" | tail -5; echo "$g"; echo "$c" | tail -3; echo "$o" | tail -8; fi
