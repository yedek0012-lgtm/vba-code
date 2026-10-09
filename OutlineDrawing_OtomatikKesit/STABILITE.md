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

**Tek DLL (1.7) bu düzeltmeyi içeriyor:** `FinalizeMember`'ın sonuna `SimetriDuzeltici.Uygula` çağrısı eklendi
(`eklenti/yama/FinalizeMemberYama.cs`, Mono.Cecil). Tangent Tower'da 14 eleman eklenir ve 33–35 m arasında iki
travers eşitlenir (101 / 101).

Ayrı eklenti DLL'i (OutlineDrawing.dll'e dokunmayan) bu durumu YÜKLE ve Oto Kesit sırasında **"DİKKAT: … simetri kodu 12; aynaları çizilmiyor"** diye bildirir.

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

AutoCAD dışında (Mono, OutlineDrawing'in gerçek okuyucusu ve AutoCAD 2022 referanslarıyla), 6 .tow dosyasında:

- Görünmeyen elemanlar: ana kule 207/219 (kalan 12 eleman, gövde yan yüz çaprazlarının traversin arkasında kalan
  kısmı; yan düşey görünüşte zaten çiziliyor), uzatmalar 42/42, 28/28, 24/24.
- Ön/yan görünüş çizgisi tanıma 100 %, 2D Aktar sonrası da 100 %.
- Kesit çizgisi eşleşmesi 100 %, yanlış redundant stili 0.
- Ölçü yerleşimi elle çizilen kesitlerle aynı (A = 4800 | 2400 | 4800 × 2400, B = 5700 | 2400, D = 5700 | 2400 | 5700).

AutoCAD içinde test edilemeyenler: forma buton ekleme, ÇALIŞTIR'a bağlanma ve silme onayı (WinForms olay listesi).
Bunlar kullanımda doğrulanmalı.

## D. Sonraki adımlar (isteğe bağlı)

1. Kaynak kodu projeye almak (A1–A6 düzeltmeleri ile birlikte). Tek DLL birleştirmeye gerek kalmaz.
2. Ayar penceresi: renkler, çizgi tipi, ölçü yazı boyu, en küçük düzlem boyu.
3. Tipik kesit birleştirme (aynı düzendeki kotlar tek kesit + "TİPİK").
4. Köşe/gergi ve çift devre kulelerle deneme.
