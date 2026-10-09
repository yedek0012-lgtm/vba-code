# OutlineDrawing – Otomatik Kesit

`OutlineDrawing v1.3.dll`'e eklenecek "Otomatik Kesit" özelliği. Kesitlerin nereden alınacağını
`.tow` geometrisinden kendisi bulur ve mevcut kesit tablosuna (grid) yazar. Kesitleri çizen kod
(SectionDetection → DrawSections) **olduğu gibi** kullanılır.

## Hazır DLL ile kurulum (önerilen, kaynak koda gerek yok)

`eklenti/bin/OtomatikKesit.dll` ayrı bir eklentidir; OutlineDrawing.dll'e dokunmaz.

1. `OtomatikKesit.dll`'i OutlineDrawing.dll ile aynı klasöre koyun.
   İnternetten indirildiyse: dosyaya sağ tık → Özellikler → **Engellemeyi kaldır** (Unblock) → Tamam.
   Bu yapılmazsa NETLOAD "Could not load file or assembly" hatası verebilir.
2. AutoCAD'de önce her zamanki gibi `OutlineDrawing.dll`'i, sonra `OtomatikKesit.dll`'i **NETLOAD** edin
   (güvenlik sorusu çıkarsa "Her zaman yükle").
3. OutlineDrawing formunu açın. "Kesit Alma" kutusunda Liste butonunun altında **Oto Kesit** butonu, **Ölçü** ve **Redundant** kutucukları belirir.
4. .tow seç → **Oto Kesit** → tabloyu kontrol et → **ÇALIŞTIR**. Komut satırından `OTOKESIT` de aynı işi yapar.

Her açılışta otomatik yüklenmesi için iki DLL'i de AutoCAD'in `APPLOAD` → "Startup Suite"ine ekleyebilirsiniz.

Eklenti, .tow dosyasını OutlineDrawing'in kendi okuyucusuyla (`dataprocessing.Run`) okur, yani koordinatlar
3D modelinizle birebir aynıdır. Yeniden derlemek için: `bash eklenti/build.sh` (Mono + NuGet AutoCAD.NET 24.1).

### Eklentinin yaptıkları (v2)

- **Oto Kesit:** OutlineDrawing'in kendi ön görünüş (`frontFace`) ve yan görünüş (`sideFace`) hesabına göre
  **görünmeyen** elemanları bulur (görünen bir elemanın ±X/±Y aynası da görülmüş sayılır). Ardından bu
  elemanların hepsi bir kesite girene kadar yatay ve eğik kesit seçer (`AutoSectionDetector.DetectHidden`).
  Her adımda en çok görünmeyen elemanı kapsayan kesit alınır; bir kesit aynalarını da kapsar.
  Kapsanamayan eleman kalırsa komut satırında `KAPSANMAYAN:` satırlarıyla listelenir.
- **Ölçü** kutucuğu (varsayılan işaretli): ÇALIŞTIR'dan sonra her `SECTION x` çizimine ölçü koyar.
  Üstte zincir ölçü (gövde köşeleri ve kol uçları), sağda toplam derinlik. Ölçüler `OTO_KESIT_OLCU`
  katmanına konur ve her çalıştırmada yenilenir. Elle: `OTOOLCU`. "2D Aktar" ölçüleri de birlikte döndürür.

- **Redundant** kutucuğu (varsayılan işaretli): ÇALIŞTIR'dan sonra ön görünüş (TRANSVERSE FACE), yan görünüş
  (LONGITUDINAL FACE) ve kesitlerdeki ana elemanlar düz çizgi ve
  cyan (4), redundant elemanlar (grup açıklaması "Redundant" olanlar) kesikli (DASHED) ve mavi (5) olur.
  Çizilen her çizgi, grid'deki kesit tanımıyla OutlineDrawing'in seçim kuralına göre bulunan elemanlarla
  rijit dönüşümle (döndürme, öteleme, aynalama) eşleştirilir. Ön/yan görünüş çizgileri ise OutlineDrawing'in
  frontFace/sideFace koordinatlarından birebir tanınır (2D Aktar sonrası da). Çizginin katmanı değişmez. Elle: `OTOSTIL`.
  Kesikli çizgi tipi `OTO_KESIK` eklenti tarafından oluşturulur (150 çizgi / 75 boşluk mm, .lin dosyası gerekmez;
  global LTSCALE'e göre ayarlanır). Sonuç ve olası hatalar formdaki UYARI satırında görünür.
- Eğik düzlem kesiti en az 6 eleman içermeli (`MinHiddenPlaneMembers`). Daha küçük düzlemler tek bir gizli
  elemanı göstermek için açılan anlamsız üçgenlerdi; o elemanlar artık `KAPSANMAYAN` olarak raporlanıyor.
  Ana kulede bunlar travers bölgesindeki gövde yan yüz çaprazları: CB-D1-T, CB-D3-T, CB-D9-T. Yan görünüşte
  traversin arkasında kaldıkları için görünmüyorlar.

Tangent Tower-0°-2°-R1 ile sonuç (OutlineDrawing'in görünürlük hesabıyla):

| Dosya | Görünmeyen eleman | Kapsanan | Kesit |
|---|---|---|---|
| Ana kule | 219 | 207 | 12 (8 yatay, 4 eğik) |
| 2BE / 5BE / 8BE | 42 | 42 | 2 |
| 0 LE / -2 LE | 28 / 24 | hepsi | 1 (bacak hip düzlemi) |

Ölçüler elle çizilen kesitlerle aynı çıkıyor: A = 4800 \| 2400 \| 4800 × 2400, B = 5700 \| 2400 × 2400,
D = 5700 \| 2400 \| 5700 × 2400.

## Mevcut durum (DLL'den okunan)

| Adım | Şu an |
|---|---|
| Yatay kesit | Kullanıcı 3D modelde nokta tıklar, Z alınır (`SelectHeightList_UntilEnter`). `SectionDetection`, iki ucu da Z ± 15 mm içinde kalan elemanları alır. |
| Eğik kesit | Kullanıcı 3–4 nokta tıklar (`SelectSectionPointSets_UntilEnter`). Orta noktası düzleme 100 mm'den yakın olan ve poligonun içine düşen elemanlar alınır. |
| Sorun | Her kule için tüm kesit kotlarını/düzlemlerini tek tek bulup tıklamak gerekiyor. Unutulan bir diyafram ya da yanlış tıklanan bir nokta çizimde fark edilmiyor. |

## Algoritma (`AutoSectionDetector.cs`)

**1. Yatay kesitler**
1. Yatay elemanlar (|ΔZ| ≤ 15 mm) kotlara göre gruplanır.
2. Her kotta, SectionDetection'ın seçeceği elemanların tamamı alınır (aynı kural).
3. Birbirine bağlı her eleman grubunun dış konturu (convex hull) hesaplanır.
   - Sadece konturda eleman var (dört yüzün kuşakları) → ön/yan görünüşte zaten görünüyor → **kesit gerekmez**.
   - Kontur içinde eleman var (plan çaprazı, diyafram, travers alt düzlemi) → **kesit önerilir**.

**2. Eğik düzlem kesitleri** (travers üst başlık düzlemi vb.). **Varsayılan kapalı** (`DetectPlanes = false`).
Gerçek kule çiziminde bu düzlemler ayrıca kesit olarak çizilmiyordu. Gerekirse açılabilir.
1. Ortak düğümü olan eleman çiftlerinden aday düzlemler üretilir.
2. Yataydan 3°–40° eğimli olanlar tutulur. Daha az eğimliler yatay kesite, daha dikler ön/yan görünüşe düşer.
3. Düzlemdeki elemanlar bağlı gruplara ayrılır. İçinde en az bir üçgen (gerçek kafes) olan grup kesit olur.
   Böylece karşılıklı yüz çaprazlarının tesadüfen oluşturduğu sahte düzlemler elenir.
4. ±X / ±Y ayna simetriği olan travers kesitlerinden yalnız biri (+X tarafı) tutulur.
5. Grubu saran en küçük dikdörtgen, 50 mm dışa büyütülerek 4 nokta olarak verilir (`First..Fourth Point`).

## Kaynak koda ekleyerek kurulum (alternatif)

1. `AutoSectionDetector.cs` ve `mainform.OtomatikKesit.cs` dosyalarını OutlineDrawing projesine ekleyin.
2. mainform tasarımcısında bir buton ekleyin (`Name = btnAutoSection`, `Text = "Otomatik Kesit"`) ve
   Click olayını `btnAutoSection_Click`'e bağlayın.
3. Kullanım: **Browse → Otomatik Kesit → (tabloyu kontrol et, istemediğini sil, eksik varsa elle ekle) → Run**.
   Satırdaki kesit harfinin üzerine gelince kesitin neden önerildiği görünür (ör. "Yatay kesit Z=24000 mm: 28 yatay eleman, 12 plan elemanı").

`test/` klasöründeki `Stubs.cs` ve `FormStubs.cs` dosyalarını projeye **eklemeyin**. Bu dosyalar yalnızca AutoCAD olmadan test etmek içindir.

## Ayarlar (`AutoSectionOptions`)

| Ayar | Varsayılan | Anlamı |
|---|---|---|
| `HeightTolerance` | 15 mm | `runtowerfile` → `SectionDetection(tol=15)` ile aynı olmalı |
| `MinPlanMembers` | 1 | Kotta en az bu kadar iç eleman varsa kesit. **0 = yatay elemanı olan her kot** |
| `MaxPlaneTiltDeg` | 40° | Daha dik düzlemler eğik kesit sayılmaz. Üçgen kesitli traversin yan yüzleri de istenirse artırın |
| `MinPlaneMembers` | 4 | Eğik düzlemde en az eleman sayısı |
| `MergeSymmetric` | true | Sağ/sol travers kesitlerinden birini tut |
| `MergeIdenticalLevels` | false | Aynı grup listesine sahip kotlardan yalnız ilkini tut ("tipik kesit") |
| `DetectPlanes` | false | true = eğik düzlem kesitlerini de ara |
| `TopDown` | true | A kesiti en üstte |

## Test

```
bash test/run_test.sh                  # derleme kontrolü + sentetik kule
bash test/run_test.sh kule.tow         # + gerçek .tow üzerinde önerilen kesitler
bash test/run_test.sh kule.tow planes  # eğik düzlem araması da açık
```

`test/tow_to_csv.py`, .tow dosyasını OutlineDrawing'in okuma kuralıyla (P/S son ekleri, X/Y/XY simetri,
-90° döndürme, mm, Z < 0 ise sıfıra kaydırma) eleman listesine çevirir. Böylece AutoCAD açmadan deneme yapılabilir.

### Gerçek kule: Tangent Tower-0°-2°-R1.TOW

439 düğüm, 1157 eleman, ~30 ms. Önerilen 7 yatay kesit ve elle çizilmiş kesitlerle karşılaştırması:

| Öneri | Kot | İçerik | Elle çizilen |
|---|---|---|---|
| A | 47800 | Tepe (toprak teli) traversi + gövde | SECTION A ✔ |
| B | 46000 | Gövde baklava plan çaprazı (H2) | SECTION C (büyük olasılıkla) |
| C | 42000 | Üst (tek taraflı) travers alt düzlemi + gövde | SECTION B ✔ |
| D | 34600 | Gövde baklava plan çaprazı (H5) | çizilmemiş |
| E | 33000 | Alt traversler + gövde | SECTION D ✔ |
| F | 15500 | Gövde uzatma üstü, baklava + köşe üçgenleri | SECTION E ✔ |
| G | 9000 | Gövde altı plan çaprazı (H8, R13–R17) | çizilmemiş |

Uzatma dosyaları (her biri ayrı .tow): 2BE, 5BE ve 8BE gövde uzatmalarında 1'er plan kesiti
(Z=9000, uzatma üstündeki plan çaprazı) öneriliyor. 0 LE ve -2 LE bacak uzatmalarında plan çaprazı olmadığı için
kesit önerilmiyor (bunlar düşey görünüş olarak çiziliyor).

Elle çizilen 5 kesitin hepsi bulunuyor. Fazladan 2 öneri gerçekten plan çaprazı olan kotlar
(çizilip çizilmeyeceği mühendis kararı). Eğik düzlem araması açılırsa travers üst başlık
düzlemleri için 3 öneri daha gelir.

Sentetik kule: 36 m gövde, 3 m paneller, Z=12000 / 21000 / 30000'de farklı plan çaprazları,
Z=24000 ve 30000'de ±X traversler (alt başlık yatay, üst başlık 21° eğik). Yüz çaprazları
karşılıklı ayna olacak şekilde kuruldu, yani sahte eğik düzlem tuzağı da modelde var.

Sonuç: 4 yatay kesit (30000, 24000, 21000, 12000) + 2 eğik travers üst düzlemi (simetri birleşik).
Sahte kesit çıkmıyor. 1 m panelli yoğun kulede (736 eleman) de aynı 6 kesit bulunuyor ve süre 50 ms'nin altında.

## Bilinen sınırlar / sonraki adımlar

- Tek bir gerçek kule ile denendi (yukarıda). Farklı kule tipleriyle (köşe/gergi, bacak uzatmalı) de denenmeli.
- Eleman simetri kodu 12 olan elemanlar OutlineDrawing'deki gibi çoğaltılmadan tek eleman alınıyor.
- Eğik kesit dörtgeninin köşe sırası, `AlignSectionMembersToPlaneFront`'un çizimi hangi yöne
  döndürdüğünü etkileyebilir. Çizim ters dönerse `BoundingQuad`'daki köşe sırası değiştirilebilir.
- Fikir: Önerilen kesitler 3D modelde geçici bir katmanda (düzlem/kot çizgisi olarak) gösterilebilir.
  Böylece Run'dan önce gözle kontrol edilir.
