# OutlineDrawing – Otomatik Kesit

`OutlineDrawing v1.3.dll`'e eklenecek "Otomatik Kesit" özelliği. Kesitlerin nereden alınacağını
`.tow` geometrisinden kendisi bulur ve mevcut kesit tablosuna (grid) yazar. Kesitleri çizen kod
(SectionDetection → DrawSections) **olduğu gibi** kullanılır.

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

## Kurulum

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
-90° döndürme, mm, Zmin = 0) eleman listesine çevirir. Böylece AutoCAD açmadan deneme yapılabilir.

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
