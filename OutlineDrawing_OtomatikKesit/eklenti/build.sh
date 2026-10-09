#!/usr/bin/env bash
# OtomatikKesit.dll derlemesi (Mono mcs). AutoCAD referansları: NuGet AutoCAD.NET 24.1.51000 (AutoCAD 2022).
# Mono'nun kendi kütüphaneleri yerine .NET Framework 4.8 referans derlemelerine karşı derlenir
# (Mono'da olup .NET Framework'te olmayan API'ler -ör. String.Trim(char)- derlemede hata versin).
#   bash eklenti/build.sh            -> eklenti/bin/OtomatikKesit.dll
set -euo pipefail
cd "$(dirname "$0")"
OUT="${TMPDIR:-/tmp}/otomatikkesit_build"
REF="$OUT/acad"
NET="$OUT/net48"
mkdir -p "$OUT" "$REF" bin

if [ ! -d "$NET/pkg" ]; then
    mkdir -p "$NET"
    curl -sSL -o "$NET/r.nupkg" "https://api.nuget.org/v3-flatcontainer/microsoft.netframework.referenceassemblies.net48/1.0.3/microsoft.netframework.referenceassemblies.net48.1.0.3.nupkg"
    python3 -c "import zipfile,sys; zipfile.ZipFile(sys.argv[1]).extractall(sys.argv[2])" "$NET/r.nupkg" "$NET/pkg"
fi
FW="$NET/pkg/build/.NETFramework/v4.8"

for p in autocad.net autocad.net.core autocad.net.model; do
    if [ ! -d "$REF/$p" ]; then
        curl -sSL -o "$REF/$p.nupkg" "https://api.nuget.org/v3-flatcontainer/$p/24.1.51000/$p.24.1.51000.nupkg"
        python3 -c "import zipfile,sys; zipfile.ZipFile(sys.argv[1]).extractall(sys.argv[2])" "$REF/$p.nupkg" "$REF/$p"
    fi
done

# Algoritma dosyası eklentinin kendi namespace'ine alınır (OutlineDrawing tipleriyle karışmasın)
sed 's/^namespace OutlineDrawing$/namespace OtomatikKesit/' ../AutoSectionDetector.cs > "$OUT/AutoSectionDetector.gen.cs"

mcs -langversion:7.2 -target:library -optimize+ -nowarn:1591 -nostdlib -noconfig \
    -out:bin/OtomatikKesit.dll \
    -lib:"$FW" -lib:"$FW/Facades" -r:mscorlib.dll -r:System.dll -r:System.Core.dll \
    -lib:"$REF/autocad.net/lib/net47" -lib:"$REF/autocad.net.core/lib/net47" -lib:"$REF/autocad.net.model/lib/net47" \
    -r:AcMgd.dll -r:AcCoreMgd.dll -r:AcDbMgd.dll -r:System.Windows.Forms.dll -r:System.Drawing.dll \
    OtomatikKesitEklenti.cs "$OUT/AutoSectionDetector.gen.cs"
echo "Derlendi: $(pwd)/bin/OtomatikKesit.dll"
