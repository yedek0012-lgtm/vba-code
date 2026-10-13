#!/usr/bin/env bash
# Bütün sistem testleri (AutoCAD'siz), verilen .tow dosyaları ve onlardan üretilen yapay varyantlar üzerinde:
#   detay   : okuma (bağımsız okuyucu = OutlineDrawing), otomatik kesit kapsaması, her eleman ailesi etiketli çiziliyor mu
#   kesit   : OutlineDrawing'in gerçek SectionDetection'ı; her kesit tam mı, çizgi eşleşmesi ve redundant stili doğru mu
#   gorunus : her elemanın önden/yandan izdüşümü çizimde var mı (yüz çizgileri + iç eleman izdüşümleri)
#   (ayrıca bir kez: ayna_parca_testi.sh, sentetik kopuk ayna parçaları)
#   olcu    : kesit ve görünüş ölçü yazıları sığıyor / çakışmıyor mu
#   konum   : ön/yan görünüş her çizim durumunda (yalnız ön/yan/ikisi, başlıklı/başlıksız, 2D) doğru yerde mi
#
#   bash test/hepsi.sh eklenti/bin/tek/OutlineDrawing.dll kule1.tow [kule2.tow ...]
#   VARYANT=0 bash test/hepsi.sh ...      -> yalnız verilen dosyalar
#   PARALEL=1 bash test/hepsi.sh ...      -> dosyaları sırayla (varsayılan: çekirdek sayısı kadar paralel)
set -uo pipefail
DLL="$(realpath "$1")"; shift
cd "$(dirname "$0")"
OUT="${TMPDIR:-/tmp}/hepsi_testi"; mkdir -p "$OUT/varyant"
DOSYALAR=()
for f in "$@"; do DOSYALAR+=("$(realpath "$f")"); done
if [ "${VARYANT:-1}" != "0" ]; then
  for f in "$@"; do
    ad="$(basename "$f" | sed 's/\.[^.]*$//' | cut -c1-24)"
    for var in dondur90 dondur180 dikdortgen eksiz grupsutun turkce uzunad buyuk bltyolu kesitsutun turkceetiket "dondur90,dikdortgen,eksiz,grupsutun,turkce,uzunad,bltyolu,kesitsutun,turkceetiket"; do
      o="$OUT/varyant/${ad}__$(echo "$var" | tr ',' '+').tow"
      python3 tow_varyant.py "$f" "$o" "$var" && DOSYALAR+=("$o")
    done
  done
fi

hata=0
# sentetik ayna parçası testi (kuleden bağımsız, bir kez)
a=$(bash ayna_parca_testi.sh "$DLL" 2>&1)
echo "$a" | grep -q "SONUC: ayna parca testi TAMAM" && echo "ayna parca testi: TAMAM" || { echo "$a"; hata=$((hata+1)); }

# test programlarını bir kez derle (paralel çalışmada her iş beklemesin)
for t in "detay_testi DetayTesti.cs DT.exe" "kesit_testi SectionDetectionTest.cs SD.exe" "gorunus_testi GorunusKontrol.cs GK.exe" \
         "gorunus_konum_testi GorunusKonumTest.cs GKT.exe" "olcu_testi OlcuCakismaTesti.cs OCT.exe"; do
  ( source ./ortak.sh; hazirla $t )
done

# dosyalar paralel (çekirdek sayısı kadar); çıktılar sırayla yazılır
SONUC="$OUT/sonuc"; rm -rf "$SONUC"; mkdir -p "$SONUC"
for i in "${!DOSYALAR[@]}"; do printf "%s\0%s\0" "$i" "${DOSYALAR[$i]}"; done |
  xargs -0 -n 2 -P "${PARALEL:-$(nproc)}" sh -c 'bash tek_dosya.sh "$0" "$3" > "$1/$(printf %04d "$2").txt" 2>&1' "$DLL" "$SONUC" 2>/dev/null || true
printf "%-58s %-8s %-8s %-8s %-8s %-8s\n" "dosya" "detay" "kesit" "gorunus" "konum" "olcu"
for i in "${!DOSYALAR[@]}"; do
  r="$SONUC/$(printf %04d "$i").txt"
  if [ -f "$r" ]; then cat "$r"; else echo "$(basename "${DOSYALAR[$i]}"): SONUC YOK"; hata=$((hata+1)); continue; fi
  n=$(head -1 "$r" | grep -o "HATA" | wc -l); hata=$((hata+n))
done
echo
[ $hata -eq 0 ] && echo "SONUC: ${#DOSYALAR[@]} dosyada bütün testler TAMAM" || echo "SONUC: $hata test HATALI"
exit $hata
