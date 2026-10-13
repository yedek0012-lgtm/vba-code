# Test betiklerinin ortak hazırlığı (source edilir). Sahte AutoCAD kütüphaneleri ve test programı bir kez derlenir,
# DLL ya da kaynak değişmedikçe yeniden derlenmez; aynı anda çalışan testler (hepsi.sh paralel) kilitle bekleşir.
#   hazirla <klasör adı> <kaynak.cs> <program.exe>   -> $OUT hazır, $OUT/<program.exe> derli
hazirla() {
  local ad="$1" kaynak="$2" exe="$3"
  OUT="${TMPDIR:-/tmp}/$ad"; mkdir -p "$OUT"
  (
    flock 9
    if [ ! -f "$OUT/Acmgd.dll" ]; then
      mcs -target:library -out:"$OUT/Acdbmgd.dll" sahte_autocad/Acdbmgd.cs
      mcs -target:library -out:"$OUT/accoremgd.dll" sahte_autocad/accoremgd.cs
      mcs -target:library -r:"$OUT/accoremgd.dll" -out:"$OUT/Acmgd.dll" sahte_autocad/Acmgd.cs
    fi
    if ! cmp -s "$DLL" "$OUT/OutlineDrawing.dll"; then cp "$DLL" "$OUT/OutlineDrawing.dll"; rm -f "$OUT/$exe"; fi
    if [ ! -f "$OUT/$exe" ] || [ "$kaynak" -nt "$OUT/$exe" ]; then
      mcs -r:"$OUT/OutlineDrawing.dll" -out:"$OUT/$exe" "$kaynak" 2>&1 | grep -v "custom attr" | grep -v "^Compilation succeeded" || true
    fi
  ) 9>"$OUT/.kilit"
}
