# OutlineDrawing + Otomatik Kesit — stabilite incelemesi

Bu belgede iki tür madde var. Eklentide (`OtomatikKesit`) düzeltilenler **sürüm 1.6** ile geldi.
OutlineDrawing kaynak kodunda yapılması gerekenler ise kodlarıyla birlikte aşağıda.

## A. OutlineDrawing kaynak kodunda düzeltilmesi gerekenler

### A1. Simetri kodu 12 olan elemanlar çizilmiyor (eksik eleman) — KRİTİK
`dataprocessing.FinalizeMember` yalnız 1 (X), 2 (Y) ve 3 (XY) kodlarını çoğaltıyor. Tangent Tower-0°-2°-R1.TOW
dosyasında 14 elemanın kodu 12: C3-2, C3-6 ve C3-19…C3-30, yani alt traversin üst düzlemi. Bu yüzden 33–35 m
arasında sağ kolda 101, sol kolda 87 eleman çiziliyor. 3D modelde, görünüşlerde ve kesitlerde 14 eleman eksik.

Aynı kulede C3-1 (kod 1, `C_4PF0.33S → C_4PF0.33Y`) ile C3-2 (kod 12, `0C_4PF0.33Y → 0C_4PF0.33S`) aynı tip enine
eleman; tek farkları tanım yönü. Üst traversin üst düzlemi tek yönlü zikzak çaprazlı. Kod 12 bu yüzden pratikte
"yalnız X aynası" (diğer travers) gibi davranıyor. Y aynası da alınırsa zikzak çaprazlar X çaprazlamaya döner, bu yanlış.

```csharp
// FinalizeMember içinde, SymmetricCode'a göre kopya üretilen yerde:
int kod = am.SymmetricCode;
bool xKopya = kod == 1 || kod == 3 || kod == 12;   // 12: X simetrisi (tanım yönü ters)
bool yKopya = kod == 2 || kod == 3;
bool xyKopya = kod == 3;
if (kod > 3 && kod != 12) errorList.Add("Bilinmeyen eleman simetri kodu " + kod + ": " + anahtar);
```

**Tek DLL (1.7+) bu düzeltmeyi içeriyor:** `FinalizeMember`'ın sonuna `SimetriDuzeltici.Uygula` çağrısı eklendi
(`eklenti/yama/OutlineDrawingYama.cs`, Mono.Cecil). Tangent Tower'da 14 eleman eklenir ve 33–35 m arasında iki
travers eşitlenir (101 / 101).

Ayrı eklenti DLL'i (OutlineDrawing.dll'e dokunmayan) bu durumu YÜKLE ve Oto Kesit sırasında **"DİKKAT: … simetri kodu 12; aynaları çizilmiyor"** diye bildirir.

### A1b. Bir eleman birden fazla kesitte çizilemiyor (eksik kesit) — KRİTİK
`SectionDetection` bütün kesitlerin elemanlarını tek sözlükte eleman adıyla tutuyor (`section[elemanAdı] = ...`).
Bir eleman iki kesite giriyorsa yalnız **sonra işlenen** kesitte çiziliyor. Tangent Tower'daki 12 otomatik kesitte
OutlineDrawing'in gerçek kodu çalıştırıldığında (test/kesit_testi.sh) şu kesitler eksik çıkıyordu:
F 10/12 (2 H5 → G), B 7/9, G 34/36, J 40/42, K 2/12.

Düzeltme: anahtar `kesitAdı + "§" + elemanAdı` olmalı. Üç yerde değişiyor: yükseklik kesiti kaydı,
nokta kesiti kaydı ve nokta kesitinin `hitKeys` listesi. `Section.name` (etiket) eleman adı kalır.

```csharp
string anahtar = row.SectionName + "§" + uyeAnahtari;
section[anahtar] = ...;          // yükseklik ve nokta kesiti
hitKeys.Add(anahtar);            // AlignSectionMembersToPlaneFront / RotateSectionMinus90AboutZ bu anahtarlarla çalışır
```

**Tek DLL (1.8) bu düzeltmeyi içeriyor** (`KesitAnahtari.Olustur`, `eklenti/yama/OutlineDrawingYama.cs`).
Yamadan sonra 6 .tow dosyasının hepsinde bütün kesitler tam.

### A2. YÜKLE ve ÇALIŞTIR model alanındaki her şeyi siliyor — veri kaybı riski
`loadtowerfile` ve `runtowerfile` model alanındaki **bütün** nesneleri siliyor. Antet, notlar ya da aynı DWG'deki başka
kuleler de gider. Önerilen düzeltme: yalnız programın kendi katmanlarındaki nesneleri silmek.

```csharp
string kat = ent.Layer;
bool bizim = kat == "0" || kat.StartsWith("SECTION", StringComparison.OrdinalIgnoreCase);
if (bizim) { ent.UpgradeOpen(); ent.Erase(); }
```

(Eklenti 1.6, bu iki butona basıldığında başka katmanlarda nesne varsa önce onay soruyor.)

### A3. Lisans kontrolü (api.ipify.org) arayüzü dondurabilir
`checkconnection` ağ isteğini AutoCAD'in arayüz iş parçacığında yapıyor. İnternet yavaşsa veya yoksa AutoCAD
zaman aşımına kadar donar. Öneri: `HttpClient.Timeout = TimeSpan.FromSeconds(5)` ve sonucun belirli bir süre
(zaten var olan `_lastCheckUtc` ile) önbelleğe alınması.

### A4. Debug derleme
DLL **Debug** yapılandırmasıyla derlenmiş (`DebuggableAttribute`, optimizasyon kapalı). Dağıtım için
**Release** ile derleyin; özellikle görünürlük hesabı (`DetectVisibleByGrid`) daha hızlı çalışır.

### A5. .tow birim kontrolü
Koordinatlar her zaman metre kabul edilip 1000 ile çarpılıyor. Dosya başlığındaki `UNITS=` değeri kontrol edilmeli.
Feet kullanılan bir dosyada çizim 3,28 kat yanlış ölçekte çıkar.

### A6. Sabit satır sayısına dayanan .tow okuma
Eklem 5 satır, eleman 8 satır kabul ediliyor. PLS-TOWER sürümü değişip bir satır eklenirse okuma sessizce kayar.
Her kaydın başında beklenen biçim (`'etiket'` satırı, `; group, section` yorumu) doğrulanmalı. Uymuyorsa
`errorList`'e açık bir mesaj yazılmalı.

### A7. AutoCAD 2025 ve sonrası
AutoCAD 2025'ten itibaren .NET 8 kullanılıyor. Hem OutlineDrawing hem eklenti .NET Framework 4.8 olduğu için o
sürümlerde yüklenmez. Geçiş gerekirse proje SDK stiline çevrilip `net8.0-windows` hedeflenmeli.

### A8. Grup tipi okunmuyor
`Parse_GroupLabel` grup satırından yalnız açıklama, boyut ve malzemeyi alıyor. PLS-TOWER'ın "group type" sütunu
(7. alan: 1 Leg, 2 Other, 3 Redundant) okunmuyor. Redundant ayrımı bu sütuna göre yapılmalı; açıklama serbest
metin olduğundan güvenilmez (YAA-R1'de 104 redundant grubun yalnız 44'ünün açıklamasında "Redundant" geçiyor).
Ayrıca satırda tam 9 alan yoksa okuma o satırda duruyor ve sonraki gruplar sessizce atlanıyor.

```csharp
// GroupLabel'e: public int groupType { get; set; }
grpLabel[t[0].Trim()] = new GroupLabel { description = t[1].Trim(), size = t[2].Trim(), material = t[3].Trim(),
                                         groupType = int.Parse(t[6], CultureInfo.InvariantCulture) };
```

**Eklenti (1.11+)** bu sütunu .tow'dan kendisi okuyor (`KuleOkuyucu.TowRedundantGruplari`).
**Tek DLL (1.19)** "tam 9 alan" denetimini "en az 9 alan" yapıyor (`Parse_GroupLabel` yaması): PLS'in yeni bir sürümü
tabloya sütun eklerse grup tablosu yine okunuyor (yamasız OutlineDrawing 0 grup okuyor, bütün etiketlerde profil boş).

### A9. 15 karakter ve üstü düğüm / eleman adı tablonun geri kalanını sessizce düşürüyor — KRİTİK
`Parse_JointsGeometry`, `Parse_SecondaryJoints` ve `Parse_AngleMemberConnectivity` kayıt sayısını kullanmıyor; "sonraki
satır bir ad satırı mı" diye `IsPoseLine` ile bakıyor. `IsPoseLine` adın **15 karakterden kısa** ve boşluksuz olmasını
istiyor. Uzun bir ad (PLS'in ürettiği kesir düğümü adları kolayca 15'i geçer, ör. `002205aPF0.333`) görüldüğü yerde
okuma durur; sonraki bütün düğümler / elemanlar ve onlara bağlı elemanlar çizimde olmaz, `errorList` boş kalır.
Deneme: her 7. düğümün adı uzatılmış YAA-R1 ve Tangent Tower'da yamasız OutlineDrawing **0** eleman okuyor.

```csharp
// IsPoseLine yerine: boş değil, ';' yok (PLS'te tablo sonundaki başlık satırlarının hepsinde var), tek ad
bool IsPoseLine(string s) {
    if (string.IsNullOrWhiteSpace(s) || s.Contains(";")) return false;
    s = s.Trim();
    if (s[0] == '\'' || s[0] == '"') return s.Length >= 2 && s.IndexOf(s[0], 1) == s.Length - 1;   // tırnaklı ad boşluk içerebilir
    return s.IndexOfAny(new[] { ' ', '\t' }) < 0;
}
// Daha iyisi: başlıktaki kayıt sayısını (ör. "128 ; Joints Geometry") kullanmak.
```

**Tek DLL (1.19) bu düzeltmeyi içeriyor** (`TowYama.PozSatiri`). Ayrıca eklenti her YÜKLE / ÇALIŞTIR / Oto Kesit'te
.tow'daki açı elemanlarını kendisi sayıp OutlineDrawing'in okuduklarıyla karşılaştırıyor; okunmayan eleman varsa
adlarıyla uyarıyor (`TowKontrol.OkumaKontrol`). Ayrı eklenti DLL'inde yama yok, uyarı var.

### A10. .tow kod sayfası
`File.ReadAllLines(path)` UTF-8 varsayıyor; PLS dosyayı Windows-1254 (Türkçe ANSI) yazıyor. Türkçe karakterler
(ç, ğ, ı, ş ...) `�` olarak okunuyor: açıklama ve grup adlarında görünür bozulma. Okuma `Encoding.GetEncoding(1254)` ile
yapılmalı. (Eklenti grup tiplerini OutlineDrawing ile aynı kod çözmeyle okuyor ki anahtarlar eşleşsin.)

### A12. Sabit görünüş ve kesit aralığı
`runtowerfile` görünüşleri 50000, kesitleri 30000 mm arayla diziyor. 30 m'den geniş bir kesit (YAA-R1'in kol kotu
kesiti 28,7 m; 35–40 m traversli kulelerde daha geniş) komşusunun üstüne biniyor; 45 m'den geniş kulede ön ve yan
görünüş de çakışıyor. Aralık kule genişliğinden hesaplanmalı.

**Tek DLL (1.19)**: `TowYama.GorunusAraligi` = en az 50000, gerekirse genişlik + 10000; `TowYama.KesitAraligi` = en az
30000, gerekirse genişlik + 6000 (1000'e yuvarlanır). Tangent Tower 30000 / 50000 (değişmedi), YAA-R1 36000 / 50000.
Eklenti kesit çizgilerini toplarken pencereyi çizimdeki gerçek kesit aralığından ölçer.

Not (hata değil, davranış): `PreparePolygon` kesit çokgenini merkezine göre %5 büyütüyor; nokta kesitine çokgenin
biraz dışındaki elemanlar da girer. Eklentinin kopyası (`KesitSecici`) 1.19'dan beri bunu da birebir uyguluyor.

### A11. Açı elemanı dışındaki elemanlar
OutlineDrawing yalnız "Angle Member Connectivity" tablosunu okuyor; kablo, gergi (guy), brace, davit kolu, X-arm
(tüplü olanlar dahil) çizilmiyor. Eklenti .tow'da bu türden eleman varsa uyarıyor.

## B. Eklentide düzeltilenler (1.6)

| Konu | Önceki risk | Düzeltme |
|---|---|---|
| Hata yakalama | Buton/komut içindeki bir hata AutoCAD'e sızabiliyordu | Bütün olay ve komutlar `Guvenli()` içinde; hata mesaj kutusu ve komut satırında |
| Ölçü stili | Çizimin ölçü stilinde DIMSCALE/DIMLFAC ≠ 1 ise yazı dev çıkar veya ölçü yanlış olurdu | Her ölçüde DIMSCALE = DIMLFAC = DIMTFAC = 1, ondalık birim |
| Kilitli katman | `OTO_KESIT_OLCU` kilitliyse ölçü yenileme tamamen başarısız olurdu | Katman kilidi açılır, silme nesne bazında korunur |
| Tekrarlı okuma | .tow her işlemde (Oto Kesit, YÜKLE, ÇALIŞTIR) yeniden okunuyordu | Yol + değişme zamanı + boyut aynıysa önbellekten |
| Silme | A2 | Başka katmanda nesne varsa onay (oturum başına bir kez) |
| Eksik eleman | A1 fark edilmiyordu | YÜKLE / Oto Kesit sırasında uyarı |
| Uyarı satırı | `PrintErrorLabel` hatası olay zincirini kırabilirdi | Korumalı, olmazsa komut satırı |

## C. Test durumu

Tek komut: `bash test/hepsi.sh eklenti/bin/tek/OutlineDrawing.dll kule1.tow kule2.tow ...` (AutoCAD'siz; Mono,
OutlineDrawing'in gerçek okuyucusu ve gerçek `SectionDetection`'ı, sahte geometri kütüphaneleriyle). Verilen her .tow
ve ondan üretilen yapay varyantlar (`test/tow_varyant.py`: 90° / 180° döndürülmüş, dikdörtgen gövde, eksi kot, yeni
grup sütunu, Türkçe grup adı, 15+ karakter ad, 3 kat büyük kule ve bunların birleşimi) üzerinde dört test:

| Test | Ne doğrulanıyor |
|---|---|
| `detay_testi.sh` | Bağımsız .tow okuyucusu ile OutlineDrawing aynı elemanları (geometri olarak) okuyor mu; otomatik kesitler görünmeyen her elemanı kapsıyor mu; **her eleman ailesi (aynalarıyla) görünüşte veya bir kesitte etiketiyle çiziliyor mu** |
| `kesit_testi.sh` | Gerçek SectionDetection her kesitte beklenen bütün elemanları çiziyor mu; çizilen her çizgi tanınıyor mu; redundant stili doğru mu |
| `gorunus_testi.sh` | Her elemanın önden / yandan izdüşümü çizimde var mı (yüz çizgileri + iç eleman izdüşümleri) |
| `gorunus_konum_testi.sh` | Ön / yan görünüş her durumda (yalnız ön / yalnız yan / ikisi, başlıklı / başlıksız, 2D Aktar, araya kesitler) doğru yerde bulunuyor mu |

Sonuç (1.19): 7 gerçek kule + 63 varyant, bütün testler TAMAM (ayrıntı BENIOKU.md 1.19).

AutoCAD içinde test edilemeyenler: forma buton / kutucuk ekleme, ÇALIŞTIR'a bağlanma, silme onayı, ölçülerin ve
kesik çizginin ekranda görünüşü, yakınlaştırınca REGEN. Bunlar kullanımda bir kez doğrulanmalı.

## D. Sonraki adımlar (isteğe bağlı)

1. Kaynak kodu projeye almak (A1–A6 düzeltmeleri ile birlikte). Tek DLL birleştirmeye gerek kalmaz.
2. Ayar penceresi: renkler, çizgi tipi, ölçü yazı boyu, en küçük düzlem boyu.
3. Tipik kesit birleştirme (aynı düzendeki kotlar tek kesit + "TİPİK").
4. Köşe/gergi ve çift devre kulelerle deneme.
