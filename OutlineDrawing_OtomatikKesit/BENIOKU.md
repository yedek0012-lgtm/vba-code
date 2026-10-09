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

**2. Eğik düzlem kesitleri** (travers üst başlık düzlemi vb.)
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
| `DetectPlanes` | true | false = yalnız yatay kesitler |
| `TopDown` | true | A kesiti en üstte |

## Test

```
bash test/run_test.sh
```

Sentetik kule: 36 m gövde, 3 m paneller, Z=12000 / 21000 / 30000'de farklı plan çaprazları,
Z=24000 ve 30000'de ±X traversler (alt başlık yatay, üst başlık 21° eğik). Yüz çaprazları
karşılıklı ayna olacak şekilde kuruldu, yani sahte eğik düzlem tuzağı da modelde var.

Sonuç: 4 yatay kesit (30000, 24000, 21000, 12000) + 2 eğik travers üst düzlemi (simetri birleşik).
Sahte kesit çıkmıyor. 1 m panelli yoğun kulede (736 eleman) de aynı 6 kesit bulunuyor ve süre 50 ms'nin altında.

## Bilinen sınırlar / sonraki adımlar

- Gerçek bir `.tow` ile henüz denenmedi. İlk denemede önerilen kesitleri elle aldıklarınızla karşılaştırın,
  toleransları buna göre ayarlayalım.
- Eğik kesit dörtgeninin köşe sırası, `AlignSectionMembersToPlaneFront`'un çizimi hangi yöne
  döndürdüğünü etkileyebilir. Çizim ters dönerse `BoundingQuad`'daki köşe sırası değiştirilebilir.
- Fikir: Önerilen kesitler 3D modelde geçici bir katmanda (düzlem/kot çizgisi olarak) gösterilebilir.
  Böylece Run'dan önce gözle kontrol edilir.
