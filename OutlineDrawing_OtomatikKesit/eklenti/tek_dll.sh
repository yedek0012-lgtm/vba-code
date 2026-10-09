#!/usr/bin/env bash
# OutlineDrawing.dll + Otomatik Kesit eklentisi -> TEK DLL (ILRepack ile birleştirme)
#   bash eklenti/tek_dll.sh /yol/OutlineDrawing.dll      -> eklenti/bin/tek/OutlineDrawing.dll
# Çıkan dosya OutlineDrawing.dll'in yerine geçer (assembly adı aynı: OutlineDrawing).
# OutlineDrawing kaynak kodunda değişiklik yapılıp yeniden derlenirse bu betik yeniden çalıştırılmalı.
set -euo pipefail
# /copyattrs: eklentinin [assembly: ExtensionApplication] özniteliği de taşınsın (yoksa buton eklenmez)
cd "$(dirname "$0")"
KAYNAK_DLL="$(realpath "$1")"
OUT="${TMPDIR:-/tmp}/otomatikkesit_build"
REF="$OUT/acad"
FW="$OUT/net48/pkg/build/.NETFramework/v4.8"
TOOL="$OUT/ilrepack"
mkdir -p bin/tek "$TOOL" "$OUT/birlesik"   # ara dosya adı OutlineDrawing.dll olmalı: ILRepack assembly adını dosya adından alır

bash build.sh >/dev/null    # referansları indirir (AutoCAD.NET 24.1, .NET Framework 4.8)

if [ ! -f "$TOOL/pkg/tools/ILRepack.exe" ]; then
    curl -sSL -o "$TOOL/i.nupkg" "https://api.nuget.org/v3-flatcontainer/ilrepack/2.0.18/ilrepack.2.0.18.nupkg"
    python3 -c "import zipfile,sys; zipfile.ZipFile(sys.argv[1]).extractall(sys.argv[2])" "$TOOL/i.nupkg" "$TOOL/pkg"
fi

# Linux dosya sistemi büyük/küçük harfe duyarlı: referanslar "accoremgd", "Acdbmgd" adıyla aranıyor
LIB="$OUT/acadlib"
mkdir -p "$LIB"
for d in "$REF"/*/lib/net47; do cp -f "$d"/*.dll "$LIB"/; done
ln -sf AcCoreMgd.dll "$LIB/accoremgd.dll"; ln -sf AcDbMgd.dll "$LIB/Acdbmgd.dll"; ln -sf AcMgd.dll "$LIB/Acmgd.dll"

sed 's/^namespace OutlineDrawing$/namespace OtomatikKesit/' ../AutoSectionDetector.cs > "$OUT/AutoSectionDetector.gen.cs"
mcs -langversion:7.2 -target:library -optimize+ -nowarn:1591 -nostdlib -noconfig -define:TEK_DLL \
    -out:"$OUT/OtomatikKesit_tek.dll" \
    -lib:"$FW" -lib:"$FW/Facades" -r:mscorlib.dll -r:System.dll -r:System.Core.dll \
    -lib:"$REF/autocad.net/lib/net47" -lib:"$REF/autocad.net.core/lib/net47" -lib:"$REF/autocad.net.model/lib/net47" \
    -r:AcMgd.dll -r:AcCoreMgd.dll -r:AcDbMgd.dll -r:System.Windows.Forms.dll -r:System.Drawing.dll \
    OtomatikKesitEklenti.cs "$OUT/AutoSectionDetector.gen.cs"

mono "$TOOL/pkg/tools/ILRepack.exe" /targetplatform:v4,"$FW" /copyattrs \
    /lib:"$LIB" /out:"$OUT/birlesik/OutlineDrawing.dll" "$KAYNAK_DLL" "$OUT/OtomatikKesit_tek.dll"

# OutlineDrawing yamaları: FinalizeMember (simetri kodu 12) + SectionDetection (kesitler arası eleman paylaşımı)
CECIL="$OUT/cecil"
if [ ! -f "$CECIL/pkg/lib/net40/Mono.Cecil.dll" ]; then
    mkdir -p "$CECIL"
    curl -sSL -o "$CECIL/c.nupkg" "https://api.nuget.org/v3-flatcontainer/mono.cecil/0.11.5/mono.cecil.0.11.5.nupkg"
    python3 -c "import zipfile,sys; zipfile.ZipFile(sys.argv[1]).extractall(sys.argv[2])" "$CECIL/c.nupkg" "$CECIL/pkg"
fi
cp -f "$CECIL/pkg/lib/net40/Mono.Cecil.dll" "$OUT/"
mcs -out:"$OUT/OutlineDrawingYama.exe" -r:"$OUT/Mono.Cecil.dll" yama/OutlineDrawingYama.cs
mono "$OUT/OutlineDrawingYama.exe" "$OUT/birlesik/OutlineDrawing.dll" bin/tek/OutlineDrawing.dll "$LIB" "$FW" "$FW/Facades"
echo "Tek DLL: $(pwd)/bin/tek/OutlineDrawing.dll"
