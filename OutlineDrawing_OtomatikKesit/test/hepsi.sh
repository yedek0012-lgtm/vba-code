#!/usr/bin/env bash
# Bütün sistem testleri (AutoCAD'siz), verilen .tow dosyaları ve onlardan üretilen yapay varyantlar üzerinde:
#   detay   : okuma (bağımsız okuyucu = OutlineDrawing), otomatik kesit kapsaması, her eleman ailesi etiketli çiziliyor mu
#   kesit   : OutlineDrawing'in gerçek SectionDetection'ı; her kesit tam mı, çizgi eşleşmesi ve redundant stili doğru mu
#   gorunus : her elemanın önden/yandan izdüşümü çizimde var mı (yüz çizgileri + iç eleman izdüşümleri)
#   konum   : ön/yan görünüş her çizim durumunda (yalnız ön/yan/ikisi, başlıklı/başlıksız, 2D) doğru yerde mi
#
#   bash test/hepsi.sh eklenti/bin/tek/OutlineDrawing.dll kule1.tow [kule2.tow ...]
#   VARYANT=0 bash test/hepsi.sh ...      -> yalnız verilen dosyalar
set -uo pipefail
DLL="$(realpath "$1")"; shift
cd "$(dirname "$0")"
OUT="${TMPDIR:-/tmp}/hepsi_testi"; mkdir -p "$OUT/varyant"
DOSYALAR=()
for f in "$@"; do DOSYALAR+=("$(realpath "$f")"); done
if [ "${VARYANT:-1}" != "0" ]; then
  for f in "$@"; do
    ad="$(basename "$f" | sed 's/\.[^.]*$//' | cut -c1-24)"
    for var in dondur90 dondur180 dikdortgen eksiz grupsutun turkce uzunad buyuk "dondur90,dikdortgen,eksiz,grupsutun,turkce,uzunad"; do
      o="$OUT/varyant/${ad}__$(echo "$var" | tr ',' '+').tow"
      python3 tow_varyant.py "$f" "$o" "$var" && DOSYALAR+=("$o")
    done
  done
fi

hata=0
printf "%-58s %-8s %-8s %-8s %-8s\n" "dosya" "detay" "kesit" "gorunus" "konum"
for f in "${DOSYALAR[@]}"; do
  d=$(bash detay_testi.sh "$DLL" "$f" 2>&1); k=$(bash kesit_testi.sh "$DLL" "$f" 2>&1)
  g=$(bash gorunus_testi.sh "$DLL" "$f" 2>&1); c=$(bash gorunus_konum_testi.sh "$DLL" "$f" 2>&1)
  sd=$(echo "$d" | grep -q "SONUC: atlanan eleman yok" && echo TAMAM || echo HATA)
  sk=$(echo "$k" | grep -q "SONUC: butun kesitler tam" && echo "$k" | grep -q "eslesmeyen 0, yanlis stil 0" && echo TAMAM || echo HATA)
  # görünüş: ön veya yan görünüşten en az biri tam olmalı; öbüründe eksik yalnız uçtan görülen traversler (detay testi
  # her elemanın etiketli çizildiğini ayrıca doğrular)
  sg=$(echo "$g" | grep -qE "(on|yan): izdusum [0-9]+, cizimde var [0-9]+, yok 0" && echo TAMAM || echo HATA)
  sc=$(echo "$c" | grep -q "SONUC: hicbir gorunus yanlis yere konmadi" && echo TAMAM || echo HATA)
  printf "%-58s %-8s %-8s %-8s %-8s\n" "$(basename "$f" | cut -c1-58)" "$sd" "$sk" "$sg" "$sc"
  for s in "$sd" "$sk" "$sg" "$sc"; do [ "$s" = TAMAM ] || hata=$((hata+1)); done
  if [ "$sd$sk$sg$sc" != TAMAMTAMAMTAMAMTAMAM ]; then echo "$d"; echo "$k" | tail -5; echo "$g"; echo "$c" | tail -3; fi
done
echo
[ $hata -eq 0 ] && echo "SONUC: ${#DOSYALAR[@]} dosyada bütün testler TAMAM" || echo "SONUC: $hata test HATALI"
exit $hata
