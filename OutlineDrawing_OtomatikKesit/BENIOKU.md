# OutlineDrawing – Otomatik Kesit

`OutlineDrawing v1.3.dll`'e eklenecek "Otomatik Kesit" özelliği. Kesitlerin nereden alınacağını
`.tow` geometrisinden kendisi bulur ve mevcut kesit tablosuna (grid) yazar. Kesitleri çizen kod
(SectionDetection → DrawSections) **olduğu gibi** kullanılır.

## Tek DLL (önerilen)

`eklenti/bin/tek/OutlineDrawing.dll` = OutlineDrawing v1.3 + Otomatik Kesit, tek dosya (ILRepack ile birleştirildi).
Eski `OutlineDrawing.dll`'in **yerine** konur; ayrıca `OtomatikKesit.dll` yüklenmez (ikisi birlikte yüklenirse çakışır).
Komutlar: DRAWOUTLINE, CONVERTO2D (OutlineDrawing'in kendi komutları) + OTOKESIT, OTOOLCU, OTOSTIL, OTOIC, OTODIZ, OTOPAFTA, OTOISARET.

**4.0: Tek düğme (ÇİZ).** Oto Kesit, Hepsini Çiz, Pafta ve OutlineDrawing'in ÇALIŞTIR düğmeleri tek düğmede: **ÇİZ** otomatik
kesitleri bulur, çizer (görünüşler, kesitler, iç eleman, ölçü, redundant) ve paftayı kurar. Shift+tık: yalnız kesitleri tabloya
yazar (kontrol / satır silme için); sonraki ÇİZ tabloyu olduğu gibi çizer. Ctrl+tık: önce pafta şablonu seçilir. Tablo son
otomatik kesitlerden farklıysa (elle düzenlenmiş) sorulur: olduğu gibi çiz / yeniden bul. Eski düğmeler gizli (OTOPAFTA vb.
komutlar duruyor); YÜKLE (3D model) yerinde.

**3.9.1: Pafta hatasız kurulur (eInvalidDxfCode).** Pafta "eInvalidDxfCode" ile duruyordu (AutoCAD'in DXF okuma hata kodu;
büyük olasılıkla şablon DXF'i okunurken, yığın izi olmadığı için kesin değil). Şablon artık sırayla denenir: seçilen şablon → gömülü ŞA-RA A1 (DXF R2018) → aynı şablonun DXF R12
hali (ezdxf r12export, `sablon/SARA_A1_pafta_r12.dxf`; MTEXT yazıya, tarama SOLID'e açılmış) → eklentinin çizdiği sade A1
çerçeve. Okunamayan şablonun sebebi uyarı etiketinde, ayrıntısı (yığın) komut satırında. Pafta içinde de aşamalar korumalı:
taşınamayan tek bir nesne, ölçek yazısı ya da malzeme listesi paftayı durdurmaz (uyarıda not edilir); yine de bir hata olursa
mesaj hangi aşamada olduğunu söyler ("Pafta (taşıma): ..."). Bütün eklenti hataları artık yığın iziyle komut satırına yazılır.

**3.9: Tek tıkta pafta, daha az tarama.**
- *Hepsini Çiz* artık Oto Kesit + ÇALIŞTIR + Pafta: tek tıkla hazır pafta (son seçilen şablonla; şablon seçmek için Pafta'ya
  Ctrl+tık). Pafta kesitleri zaten yeniden yerleştirdiği için bu akışta ÇALIŞTIR sonrası kesit dizme ve ara REGEN atlanır
  (pafta kurulamazsa kesitler yine dizilir). Pafta ve REGEN süresi komut satırına yazılır.
- Kesit gruplama (bir ÇALIŞTIR + paftada 6 kez çağrılıyor: işaretleme, tam travers, kopuk parça, kesit yönü, dizme, pafta)
  model alanını üç yerine bir kez tarar; çizgilerin sınırı GeometricExtents yerine uç noktalarından alınır (çizimin çoğu
  çizgi). Aile uzatma çizimlerinde eski çizimleri silme ve çizim sınırı tek taramada; redundant stilinde görünüş çizgileri ve
  eleman isimleri tek taramada. Zaten aynı kesit adıyla işaretli nesnenin işareti yeniden yazılmaz (taşınmıyorsa nesne yazmaya
  açılmaz). Hesap tarafı (kesit bulma, pafta tahmini, iç eleman, işaret) zaten 1,5 s'nin altında (YAA, TAO, Tangent).

**3.8: Yerel kesit tam panel.** *Tüm elemanlar* işaretliyken görünüşlerde ve diğer kesitlerde olmayan tek bir eleman için
açılan yerel kesit (Tangent: kafesin sol yüzünde alt traversin üstündeki CB-D9 çaprazı, SECTION F) elemanın doğrultusunda
dönük dar bir dörtgendi: OutlineDrawing kuşakların yarısını, üst panelin tek çaprazını ve traversten bir çubuğu çiziyor,
kesit parça gibi görünüyordu. Dörtgen artık düzlemin eksenlerinde (yatay doğrultu / düzlemde yukarı) ve elemanın panelini
tam sarıyor: panelin kuşakları ve çaprazları bütün çizilir (`yerel_kesit_Tangent_F.png`). Sade modda (varsayılan) yerel
kesit zaten alınmaz; bu eleman görünüşte iç eleman çizgisi olarak durur.

**3.7: Paftayı doldur (ek kesitler).** Yeni kutu *Paftayı doldur* (varsayılan işaretli): Oto Kesit, paftada boş kalacak
yer varsa kapsama için gerekmeyen ama çizimcinin koyduğu türden ek kesitleri de tabloya yazar — bacak başı köşe planı (TAO
K1 gibi), kafes / gövde plan çaprazları. Paftanın yerleşimi ÇALIŞTIR'dan önce tahmin edilir (`PaftaTahmini`: görünüşler,
kesit işaretleri, kesitler, ölçüler, aile uzatma çizimleri, malzeme listesi); ek kesit yalnız pafta ölçeği büyümeden
sığıyorsa eklenir (kutular %4 büyük de denenir, denenen kesit ayrıca %10 büyük). Ek kesit olmayan yerler: mevcut yatay
kesite 300 mm'den yakın kot, yalnız görünüşteki kuşaklardan ya da yalnız dış kuşaklardan oluşan plan, iki kol ucu gibi
uzak parçalı uzun plan, büyük çerçevede küçük köşe parçaları, mevcut ya da önceki bir planla aynı desen (bacak içindeki
ardışık kotlar, aynı kafes planı). Ek kesitler tabloda "Ek kesit (pafta boşluğu)" ipucuyla görünür; istenmeyen satır
silinebilir. Tangent (yalnız ön görünüş, listesiz): C (kafes planı) ve K (bacak başı) eklenir, ölçek 1:90 aynı kalır
(`pafta_doldur_Tangent.png`). YAA ve TAO paftası zaten dolu: ek kesit sığmıyor, eklenmez.

**3.6: Bacak iç düzlemi işareti kasıktan ayağa.** Tangent kulesinde bacağın iç düzlem kesiti (SECTION J, kasıktaki iki iç
düğüm ile ayak) alınıyordu, ama ön görünüşte işareti bacağın yanına yatay ok olarak konuyordu: OutlineDrawing'in kesit
dörtgeni düzlemde dönük geldiği için alt köşesi yerden 1.2 m yukarıda kalıyor, düzlem "bacak" sayılmıyordu. Kesit hattı
artık düzlemde yatan ön yüz çizgilerinden (bacak iç çizgisi) bulunur: üst ucu kasık düğümü, alt ucu ayak. TAO'da da
bacak iç düzlemi işareti 1.8 m yukarıdan değil kasık düğümünden başlıyor.

**3.5: Genel kontrol düzeltmeleri.** Kod incelemesi (AutoCAD tarafı, çalıştırılamayan kısımlar) ve üç kulenin pafta
önizlemesiyle (`test/onizleme_ciz.py`, `OTO_ONIZLEME=dosya bash test/olcu_testi.sh ...`):
- Aynı işlemde iki katman / uygulama kaydı eklenince tablo ikinci kez yazmaya açılıyordu (AutoCAD eWasOpenForWrite):
  ilk Pafta (pafta + liste katmanı) düşebiliyordu. Bütün tablo yükseltmeleri korumalı.
- Pafta / dizme sonrası tek başına OTOOLCU / OTOSTIL / OTODIZ / OTOPAFTA: işaretli nesnesi olan kesite yakınlıkla nesne
  eklenmez (liste ve görünüş çizgileri kesite karışmıyordu), pafta / liste çizgileri kesit toplamasına girmez.
- Kesit sırası başlıkta tutulur (OTO_KESIT_SIRA): dizilmiş kesitler tekrar dizilince ya da paftaya konunca A-B-C sırası
  bozulmuyor (önce yatay konuma göre sıralanıp satırlar karışıyordu).
- Bacak boyu çizgileri kendi katmanında (OTO_BACAK_BOYU): OTOIC artık onları silmiyor.
- OTOSTIL tam travers kesitini (iki kol + gövde) tanır; öbür kolun redundantları kesikli kalır.
- Eksi uzatma adları ("-1.5 LEG EXTENSION"), kilitli pafta katmanında yeniden pafta, pafta nesnelerinin "yabancı nesne"
  uyarısı, katman adına uygun olmayan kesim etiketi, "Redundant" kapalıyken eklenti çizimlerinde kesikli redundant.

**3.4: Yatay kesitte kopuk uç parçaları.** OutlineDrawing yatay kesitte o kottaki bütün elemanları çizer; YAA-R1'de
Z=32500 (SECTION I) kol ucundaki 2 × 152 mm'lik "F" parçasını da alıyor, parça ana çizimden 7 m uzakta tek başına
duruyordu (yanındaki kesite aitmiş gibi görünüyordu). Ana parçanın %10'undan kısa, ana parçaya max(2 m, genişliğin %25'i)'nden
uzak ve elemanları başka bir kesitte gösterilen parçalar yatay kesit çiziminden kaldırılır (`KopukParca`,
`KopukParcaTemizle`; ölçülerden önce). Diğer test kulelerinde böyle parça yok.

**3.3: Bacak uzatması çizimlerinde üst çizgi.** Bacak uzatması görünüşlerinde ve L kesitlerinde üst yatay (bacak yüzünde
H1, L kesitinde iki kasık düğümünü bağlayan SH1) çizilmiyordu: bu elemanlar bacak bloğunda değil, aynı konfigürasyonun
gövde uzatması bloğunda tanımlı. Bacağın üst kotunda ve bacak alanının içindeki gövde uzatması yatayları bacağa eklenir.

**3.2: Dizmede kopan parçalar düzeltildi.** Aile dosyasında (TAO) "Kesitleri diz" sonrası geniş kesitlerin (tam travers)
uçları, bazı ölçüler ve K / L planlarının yarıları eski yerinde kalıyordu. Sebep: işaretsiz nesneler en yakın kesit
başlığına tek bir pencereyle bağlanıyordu ve pencere en sık iki başlığın aralığından hesaplanıyordu; en sağa sık dizilen
uzatma çizimleri pencereyi daraltınca uçtaki nesneler hiçbir kesite bağlanmadı (benzetimde 273 nesne).
- OutlineDrawing kesitleri çizer çizmez (düzen henüz eşit aralıklı) bütün kesit nesneleri kesit adıyla işaretlenir.
- Ölçüler kesit adıyla işaretlenerek oluşturulur.
- Gruplama penceresi başlık başına (komşu başlıklara uzaklığın büyüğünün yarısı).
- Test: ölçü / pafta testi ÇALIŞTIR sonundaki yerleşimde her çizgi ve ölçünün kendi başlığına bağlandığını denetler.

**3.1: Aile paftasının eksikleri tamamlandı, kararlılık.**
- *Tam travers kesiti*: OutlineDrawing eğik travers düzleminde yalnız bir kolu çiziyordu. Gövdeye bağlı simetrik kolun
  eğik düzleminde (Y doğrultusunu içeren, 5°-45°, aynası elemanlı, kesit tek tarafta) kesit, çizimcinin B, D, F, H
  kesitleri gibi iki kol gerçek boyunda + kök kotundaki gövde planı olarak yeniden çizilir (`TraversKesiti`,
  `TraversCizimi`); ön görünüşte işareti iki kolun ucuna konur. Kapsama hesapları bunu bilir: kökteki gövde planını
  ayrıca gösteren kesit (TAO: Z=59310) gereksiz sayılıp atılır. Önizleme: `tam_travers_TAO.png`.
- *Sade*: dik olmayan eğik düzlemde 8'den az elemanlı parça kesit alınmaz (TAO: bacak bölgesindeki 6 elemanlık köşegen).
- *Uzatma görünüşleri*: ölçülendirilir (kesitler gibi), yüzde olmayan elemanların izdüşümü kesikli çizilir.
- *Bacak boyları*: ana görünüşte sol bacakta diğer bacak uzatmalarının dış / iç bacak çizgileri kesikli, ayak kotlarında
  soldan "+4 LEG ... +0 LEG" yazıları.
- Bacak uzatması hangi köşede tanımlıysa oradan okunur (döndürülmüş kulelerde de); ayrı eklenti DLL'inde aile dosyası
  açılırsa uyarı; çizgi tipi yüklenemezse çizim sürer.
- Test: detay testi aile dosyasında uzatma çizimlerini (görünüş, K, L sayısı ve boyutları) denetler; ölçü / pafta testi
  tam travers ve uzatma çizimleriyle yapılır (TAO: 31 parça 1:155, binen yok). 91 + 13 (TAO) dosya TAMAM.

**3.0: PLS-TOWER aile (family) dosyaları.** 220kV TAO gibi aile dosyalarında gövdeden sonra her gövde uzatması ×
bacak uzatması birleşimi ayrı konfigürasyondur (TAO: 5 × 5 = 25). OutlineDrawing hepsini birden okuyordu: blok
düğümleri açılmış adlarla ('100P') yazıldığından ve OutlineDrawing her düğüm adına "P" eklediğinden düğümler yanlış
eşleşip uzatmalar üst üste biniyor, görünüşün alt kısmı karmakarışık çıkıyordu.
- OutlineDrawing'e yalnız dosyanın baş kısmı verilir: PLS onu varsayılan tam model olarak yazar (TAO: gövde + 12 m gövde
  uzatması + 4 m bacak, çizimcinin ana görünüşü). Ana görünüş ve otomatik kesitler buradan (`TowAilesi`).
- Diğer uzatmalar eklenti tarafında okunur (`AileParcalari`; bacak bloklarındaki PLS kaydırması dx, dy, dz uygulanır) ve
  ÇALIŞTIR sonunda çizimin sağına çizilir (`UzatmaCizimleri`, katman "SECTION n", grup başına etiket, OTO_AILE işaretli):
  her gövde uzatması için ön yüz görünüşü ("+9 BODY EXTENSION") ve alt planı (SECTION K2..), her bacak uzatması için ön
  yüz görünüşü ("+3 LEG EXTENSION") ve iç düzlemi (SECTION L2.., ayak ile iki kasık düğümünden geçen ters üçgen).
  Ölçüler TAO paftasıyla aynı (K2-K5: 11384, 10540, 9694, 8850; L2-L5: 7396, 6482, 5604, 4784 yükseklik).
  K / L kesitlerine ölçü konur; hepsi kesitlerle birlikte dizilir ve paftaya yerleşir.
- Bacak başlarındaki dört köşe parçası artık tek köşe yerine tam plan (TAO K1 gibi).
Önizleme: `aile_ana_gorunus_TAO.png`, `aile_uzatmalari_TAO.png`.

**2.9: "Tüm elemanlar" kutusu ve kesit yönü.**
- *Tüm elemanlar* kutusu (varsayılan işaretsiz): işaretsizken görünüş seçiminden bağımsız olarak sade kesit alınır,
  yani her gizli elemanı bir kesite sokmaya çalışılmaz (YAA-R1'de SECTION K gibi 5.7°'lik parça kesit çıkmaz). Kesit
  almayan elemanlar görünüşte iç eleman çizgisi olarak durur. İşaretlenirse her gizli eleman bir kesitte gösterilir.
- *Kesit yönü*: OutlineDrawing eğik düzlem kesitini düzlemin eğim yönü yukarı gelecek şekilde çiziyordu; travers üst
  düzlemleri traversi dikey (ucu aşağıda) veriyordu. ÇALIŞTIR sonunda plana yakın (45°'den yatık) eğik kesitler, düzlemde
  kulenin X eksenine en yakın doğrultu yatay ve +X sağda olacak şekilde 90° döndürülür (YAA-R1 G: saat yönü tersi; TAO
  travers üst düzlemleri PDF'teki gibi yatay). Dik düzlemler (yüz şeridi, çerçeve) ve bacak iç düzlemleri olduğu gibi
  kalır. Yazılar okunur yöne çevrilir, başlık kesitin altına ortalanır; aynı kesit ikinci kez döndürülmez.

**2.8: Kule yüzü kesit olmaz; bacak kesiti çizimci gibi işaretlenir.** Yalnız TRANSVERSE FACE seçiliyken yan yüz
(LONGITUDINAL) görünüş yerine yüz büyüklüğünde bir kesit olarak veriliyordu (YAA-R1: SECTION U, 150 eleman). Artık ön ve
yan yüz elemanları görünüş seçiminden bağımsız olarak görünüş sayılır; yüz düzlemi hiçbir zaman kesit olmaz (YAA-R1, yalnız
ön görünüş: 26 → 24 kesit). Bacağın iç düzlemi (TAO'daki SECTION L gibi; YAA-R1'de SH-C / K-C_35 / K-C-R_35 kesiti)
ön görünüşte kasıktaki iç düğümden ayağa uzanan kesit hattıyla, iki ucunda içeriden dışa bakan okla işaretlenir.

**2.7: Düzenli kesit işaretleri.** OutlineDrawing her kesit için ön görünüşe L oklu bloklar koyuyordu; eğik ve parça
kesitlerde bunlar gövdenin içine, eleman yazılarının ve birbirinin üstüne düşüyordu. ÇALIŞTIR sonunda (komut OTOISARET)
bu oklar kaldırılıp çizimci gibi işaretler konur (katman SECTION_MARK): yatay kesitte o kotta kulenin iki yanında kesit
çizgisi + aşağı ok + harf; travers üst düzlemi gibi ön görünüşte çizgi olan düzlemde izin dış ucunda bakış yönünde ok;
köşegen bacak düzlemlerinde o kotta kulenin dışından gösteren ok. Birbirine binen işaret dışa kaydırılır.
Önizleme: `kesit_isaretleri_YAA.png`.

**2.6: Sade (çizimci gibi) kesit, varsayılan.** Çizimcinin YAA_OUTLINE_rev0'da elle bıraktığı kesitlerle
karşılaştırıldı: 16 kesitin 15'ini Oto Kesit de aynı yerden alıyor; fazladan verdiği parçalar çizimcinin sildiği türden.
Ön + yan görünüş birlikte çizilirken artık şunlar kesit olmaz: tek panellik yerel kesitler (aynı düzlemde başkalarıyla
birleşenler şerit kesit olarak kalır), neredeyse yatay (15°'den az) eğik düzlemler, 4'ten az elemanlı parçalar.
Bu elemanlar görünüşte iç eleman çizgisi olarak durur (test: hiç çizilmeyen eleman 0); çizim sonrası denetim onları eksik
saymaz. Tek görünüş seçiliyken kapsama tam kalır. YAA-R1: 24 → 21 kesit, Tangent R1: 13 → 12. Kapatmak:
`AutoSectionOptions.Sade = false`.

**2.5: Aynı düzlemdeki parça kesitler tek kesit.** Çizimci, traversin arkasında kalan gövde yüzünü tek bir şerit
kesitte gösterir; Oto Kesit ise her gizli panel için ayrı küçük "yerel kesit" veriyordu. Artık aynı düzlemde olup o
düzlemdeki elemanlarla birbirine bağlanan düzlem kesitleri tek kesitte birleşir (220kV TAO: gövde yüzündeki 4 parça →
1 şerit, toplam 18 → 15 kesit; kapsama değişmedi, 91 test dosyası TAMAM). Kapatmak: `AyniDuzlemiBirlestir = false`.
Not: TAO gibi PLS *family* dosyalarında OutlineDrawing bütün gövde/bacak uzatmalarını üst üste okur; uzatmaya özel
kesitler (+3/+6 gövde uzatması yüzleri, K1-K5 gibi) bu yüzden ayrı ayrı çıkarılamaz, her uzatma için ayrı .tow gerekir.

**2.4: Pafta ve malzeme listesi.**
- *Pafta* butonu (komut OTOPAFTA, ÇALIŞTIR'dan sonra): ŞA-RA A1 antetli çerçeve (YAA_OUTLINE_rev0'dan çıkarıldı, DLL'e
  gömülü: `eklenti/sablon/SARA_A1_pafta.dxf`) ön görünüşün etrafına kurulur; yan görünüş yanına (yalnız yatay kayar),
  kesitler kalan alana (görünüşlerin sağı ve altı, malzeme listesinin altı) A-B-C sırasıyla yerleşir; ölçek otomatik,
  5'in katı (YAA-R1: 1:130) ve antetteki SCALE hücresine yazılır. Ctrl+Pafta: başka şablon (DWG/DXF; sol alt köşe 0,0).
  Yeniden basılırsa önceki pafta silinip yeniden kurulur. Önizleme: `pafta_onizleme_YAA.png`.
- *Malzeme listesi*: .tow grup tablosundan, şirket paftasındaki Excel tablosu düzeninde (Group Label | Group Desc. |
  Angle Size | Strength(Mpa), S355 → 355.0), sağ üstte 3 sütun (YAA-R1 222, Tangent 284 grup). Katman OTO_LISTE.
- *Form*: eklenti butonları ve kutucukları ÇALIŞTIR'ın sağına, kutunun içine dizilir (önceden "İç eleman" ve
  "Kesitleri diz" kutunun sağ kenarının dışında kalıp görünmüyordu).
- Test: `olcu_testi.sh` her kulede paftanın sığdığını, hiçbir parçanın birbirine, listeye ya da antete binmediğini
  ve hepsinin çerçeve içinde kaldığını denetler.

**2.3: Kesitler ızgaraya dizilir.** OutlineDrawing kesitleri en geniş kesit kadar aralıkla tek uzun sıraya diziyordu
(YAA-R1: 24 kesit, 829 m × 20 m). Artık ÇALIŞTIR sonunda ("Kesitleri diz" kutusu, komut OTODIZ) kesitler kendi
boyutlarına göre, A-B-C sırasıyla satırlar halinde sıkı bir bloğa yerleşir (YAA-R1: 4 satır, 71 m × 52 m; Tangent
Tower R1: 14 kesit, 403 m → 51 m × 36 m). Her kesitin çizgileri, eleman etiketleri, başlığı ve ölçüleri birlikte
taşınır ve kesit adıyla işaretlenir (XData OTO_KESIT); OTOOLCU / OTOSTIL dizilmiş kesitleri bu işaretle tanır. Tek
işlemde yapılır: hata olursa hiçbir şey taşınmaz. Örnek: `kesit_dizilim_YAA.png`. Test: `olcu_testi.sh` dizimde
hiçbir kesitin (ölçü ve başlığıyla) bir başkasına binmediğini denetler.

**2.2: Gereksiz kesit alınmaz.** Otomatik kesitler bulunduktan sonra, OutlineDrawing'in her kesitte gerçekte çizeceği
elemanlara (KesitSecici: %5 büyütülen poligon, 100 mm düzlem toleransı) bakılır; gösterdiği bütün görünmeyen elemanlar
başka kesitlerde de bulunan kesit tabloya yazılmaz (Tangent Tower R1: 15 → 14). 7 kule + 84 varyantta her kesit en az
bir elemanı tek başına gösteriyor; hiçbir eleman açıkta kalmıyor.

**2.1: Yan / ön görünüşte iç eleman çizgilerinin görünmemesi düzeltildi.**
- *Kök neden:* OTO_IC_ELEMAN / OTO_IC_ELEMAN_UC katmanı kapalı ya da dondurulmuşsa (ÇALIŞTIR nesneleri siler, katman
  durumunu korur) çizgiler çiziliyor ama görünmüyordu; uyarı oturumda bir kez, UC katmanı için hiç çıkmıyordu.
  OutlineDrawing kendi katmanlarını her çizimde açıyor; artık eklenti de iç eleman ve ölçü katmanlarını her çizimde açar
  (açtığını özette yazar). Özet ön / yan çizgi sayısını ayrı gösterir.
- *İç eleman kutusu* kapalı haliyle hatırlanmaz (her açılışta işaretli); kapalıysa uyarı kutusunda da yazar.
- *Ölçü yazısı stili:* ölçü stilinin yazı stili sabit yükseklikliyse (şablonda 2,5 gibi) AutoCAD yazı yüksekliğini yok
  sayıyordu, yazılar görünmeyecek kadar küçük çıkıyordu; o durumda aynı yazı tipinde yüksekliği 0 olan OTO_OLCU stili
  kullanılır.
- *Geçerli katman:* eklenti katmanıyken "0"a alma artık belge kilidiyle yapılır (önceden sessizce olmuyordu).

**2.0: Ölçülmüş, daha güvenli ve daha pratik sürüm.** (Önceki sürüm `eklenti/bin/yedek_1.26/OutlineDrawing.dll`'de;
bir sorun olursa onunla değiştirilebilir.)
- *UCS güvenliği:* 1.26'dan beri ÇALIŞTIR sonrası UCS kesit düzlemine (XZ) alınıyor. OutlineDrawing'in Nokta / Yükseklik
  seçimi `GetPoint` sonucunu UCS'den dünyaya çevirmeden kullandığı için, sonradan 3D modelden seçilen noktalar yanlış
  koordinat verirdi. Artık YÜKLE, ÇALIŞTIR, Nokta ve Yükseklik'ten hemen önce UCS dünya yapılır; ÇALIŞTIR bitince
  yeniden kesit düzlemine alınır.
- *ÇALIŞTIR özeti:* uyarı satırında iç eleman / redundant / ölçü sonuçları ve süre (OutlineDrawing'in kendi çizimi +
  eklentinin adımları); komut satırında adım adım süreler. İç eleman kapalıysa "İç eleman: kapalı" yazar.
- *Sorunlar tek kutuda:* çizilmeyen kesit, kesitte olmayan eleman, iç eleman çizgisi eklenemedi / görünüş bulunamadı,
  bir adımın hatası ve **kapalı / dondurulmuş eklenti katmanı** (OTO_IC_ELEMAN, OTO_IC_ELEMAN_UC, OTO_KESIT_OLCU: çizgiler
  çizilir ama görünmez) artık 8 saniyede kaybolan etiket yerine bir uyarı kutusunda gösterilir.
- *İç eleman çizgileri tek tek korunur:* bir çizgideki hata (ör. çizgi tipi) öbürlerini düşürmez; kaç tanesinin
  eklenemediği bildirilir.
- *Hepsini Çiz:* tablo bu .tow ve görünüş seçimi için Oto Kesit'in yazdığı haliyle duruyorsa yeniden hesaplanıp
  "silinsin mi?" diye sorulmaz, doğrudan çizilir; tabloyu elle değiştirince yine sorar.
- *Önbellekler:* Oto Kesit sonucu (aynı .tow + görünüş seçimi) ve .tow kontrolü dosya değişmedikçe yeniden hesaplanmaz;
  model alanı taramalarında yalnız çizgi / yazı nesneleri açılır.
- *Test:* ölçü yazısı çakışma testi (`olcu_testi.sh`) eklendi; test programları bir kez derlenir ve dosyalar paralel
  çalışır (tam test 30 dakikadan birkaç dakikaya indi).

Ölçümler (AutoCAD dışında, en büyük kule YAA-R1, 1386 eleman): .tow okuma 0,16 s, Oto Kesit 0,57 s, iç eleman
hesabı 0,03 s, kesit denetimi 0,06 s, .tow kontrolü 0,08 s. Eklentinin hesapları 1 saniyenin altında; ÇALIŞTIR'da
hissedilen süre büyük ölçüde AutoCAD'in çizimi ve REGEN'dir (artık özette ayrı ayrı görünür).

**1.22: Son denetim — okuma, kesit ve görünüş güvenceleri.**
- *Tablo sonu kayıt sayısından:* düğüm / ikincil düğüm / eleman tabloları başlıktaki kayıt sayısıyla ("128 ; Joints
  Geometry") okunur. 1.19–1.21'de Secondary Joints'ten sonraki .blt yolu boşluksuzsa (ör. `D:\PROJE\bolt.blt`)
  düğüm adı sanılıp okuma çöküyor, hiçbir şey çizilmiyordu.
- *Okunamayan .tow:* OutlineDrawing dosyayı okuyamazsa YÜKLE / ÇALIŞTIR model alanını silmeden durur ve sebebi gösterir.
- *Kesit etiketi tablosu:* "tam 15 alan" yerine kayıt sayısı + en az 15 alan (yeni PLS sürümünde katmanlar oluşmuyor,
  ön / yan görünüş hiç çizilmiyordu).
- *tow_section_dump.txt yazılamazsa* (salt okunur klasör, Excel'de açık) bütün kesitlerin atlanması düzeltildi.
- *Kod sayfası:* .tow UTF-8 değilse Windows-1254 okunur; yalnız Türkçe harfte farklı iki etiket artık tek anahtara inip
  birini kaybettirmiyor.
- *Ayna parçası:* tek parça kesiti için parçaların ayna olduğu uç noktalarıyla denetlenir (yalnız orta noktaya bakınca
  ters köşegen ayna sanılıp hiçbir kesitte çizilmiyordu).
- *Ön / Yan Yüz kutucukları:* Oto Kesit yalnız işaretli görünüşleri "görünür" sayar; ikisi de işaretsizse ikisini
  işaretler (işaretsiz görünüşün elemanları hiçbir yerde çizilmiyordu).
- *ÇALIŞTIR çizmezse* (tabloda hatalı satır, silme uyarısında Hayır) eklenti eski çizime dokunmaz ve OutlineDrawing'in
  hata mesajını ezmez; "ÇİZİM YENİLENMEDİ" kutusu çıkar.
- *ÇALIŞTIR sonrası kesit denetimi:* tablodaki her kesit çizildi mi (elemanı olmayan satır sessizce atlanıyordu) ve
  çizilen görünüşlerde görünmeyen her eleman (aynasıyla) bir kesitte mi; değilse uyarı kutusu + komut satırında liste.
- *2D Aktar sonrası* görünüşler gerçek dönüşümle, (x, z, 0), bulunur; görünüş bulunamazsa iç eleman çizgileri ve
  ölçüler silinmez. Bu katmanlarda yalnız çizgi / ölçü silinir; geçerli katman eklentinin katmanıysa ÇALIŞTIR öncesi
  "0" yapılır (başlıklar o katmana düşüp silinmesin).

**1.20: Kesit harfleri kendiliğinden sıralanır.** Tablodan bir kesit satırı silinince (satırı seçip Delete) veya bir
satırın içi boşaltılınca harfler yeniden dağıtılır: verisi olan satırlar sırayla A, B, C … (Z'den sonra AA, AB …), içi
boş satırın harfi silinir. OutlineDrawing harfleri yalnız Nokta / Yükseklik ile satır eklenince yeniliyordu.

**1.19: Sistem denetimi — hiçbir eleman atlanmasın.**
- *Yerel panel kesiti:* açgözlü kesit seçiminden sonra hâlâ hiçbir kesitte olmayan görünmeyen eleman için yalnız onun
  panelini gösteren küçük bir düzlem kesiti eklenir (Tangent Tower CB-D1/D3/D9-T, YAA-R1 CA-9). Görünmeyen eleman
  kapsaması 7 kulede %100; her eleman ailesi görünüşte veya bir kesitte etiketiyle çiziliyor.
- *OutlineDrawing okuyucusu (tek DLL yaması):* 15 karakter ve üstü düğüm / eleman adı tablonun geri kalanını sessizce
  düşürüyordu (uzun adlı denemede 0 eleman okunuyordu) -> düzeltildi (STABILITE A9). Grup tablosu fazladan sütunla da
  okunuyor (A8).
- *Okuma doğrulaması:* YÜKLE / ÇALIŞTIR / Oto Kesit'te eklenti .tow'daki açı elemanlarını kendisi sayıp OutlineDrawing'in
  okuduklarıyla karşılaştırır; okunmayan eleman, düğümü tanımsız eleman, açı elemanı olmayan eleman (kablo, davit,
  X-arm) ve okunamayan grup satırı varsa adlarıyla uyarır.
- Grup tipleri OutlineDrawing ile aynı kod çözmeyle okunur: Türkçe karakterli grup adlarında redundant ayrımı doğru.
- *Görünüş / kesit aralığı (tek DLL yaması):* sabit 50000 / 30000 yerine kule genişliğine göre (en az aynı değerler);
  geniş traversli kulelerde kesitler üst üste binmez (STABILITE A12). Kesit çizgileri toplanırken pencere çizimdeki
  gerçek kesit aralığından ölçülür.
- Kesit seçiminin kopyası (stil ve ölçü için) OutlineDrawing'in düzlem eksenleri ve %5 büyütülen çokgeniyle birebir.
- Form: üç kutucuk "Kesit Alma" kutusuna sığmazsa yan yana dizilir.
- Ölçülerden biri AutoCAD tarafından reddedilirse öbürleri yine konur; aynı bilgi komut satırına bir kez yazılır.
- Test: `bash test/hepsi.sh eklenti/bin/tek/OutlineDrawing.dll *.tow` — 7 gerçek kule ve onlardan üretilen 63 yapay
  varyant (döndürülmüş, dikdörtgen gövde, eksi kot, yeni grup sütunu, Türkçe ad, uzun ad, büyük kule): okuma,
  kapsama, kesit çizimi ve stil, görünüş izdüşümü, görünüş yeri testlerinin hepsi TAMAM.

**1.18: Görünüşün yeri başlıktan bulunur (yalnız ön görünüş çizildiğinde yan görünüş üstüne biniyordu).**
OutlineDrawing ön görünüşü ofset 0'a, yan görünüşü ön de çizildiyse 50000'e, çizilmediyse 0'a koyuyor. Kare gövdeli
kulelerde yan yüz çizgilerinin çoğu ön yüzünkilerle birebir çakışıyor (YAA-R1 195'te 136, Tangent 198'de 186);
eklenti yalnız TRANSVERSE FACE çizildiğinde onu yan görünüş de sanıyor, yan görünüşün ölçülerini ve iç elemanlarını
(tepe, girder, bel X'leri) ön görünüşün üstüne çiziyordu. Artık görünüşün yeri ve hali "TRANSVERSE FACE" /
"LONGITUDINAL FACE" başlığından bulunup çizgilerle doğrulanıyor; başlık yoksa yalnız o yüze özgü çizgilere bakılıyor,
yüzler özdeşse görünüş kullanılmıyor (`GorunusKonumBulucu`). Redundant stili de her görünüşü kendi yerinde ve kendi
çizgileriyle tanıyor. Komut satırına bulunan görünüşler yazılır. Test: `test/gorunus_konum_testi.sh` (yalnız ön /
yalnız yan / ikisi; başlıklı / başlıksız; 2D Aktar; araya kesit çizgileri): 7 kulede hiçbir görünüş yanlış yere konmuyor.

**1.17: Yakınlaştırınca kesikler kendiliğinden görünür.** AutoCAD kesik deseni REGEN anındaki yakınlaştırmaya göre
üretiyor: ÇALIŞTIR'dan sonra bütün çizim ekrandayken 150/75 mm'lik desen birkaç piksel kalıyor, çizgi düz çiziliyor ve
yakınlaştırınca REGEN olmadan düz kalıyor (LTSCALE değiştirmek REGEN yaptığı için kesikler o zaman görünüyordu).
Eklenti artık redundant stili uygulanmış çizimde görünüm son REGEN'dekinden 1,5 kat yakınlaşıp durunca kendisi REGEN
yapar (`KesikIzleyici`, VIEWSIZE izlenir; komut çalışırken bekler). Elle: `RE`.

**1.16: Kesikli çizgiler ÇALIŞTIR'dan hemen sonra görünür.** AutoCAD yeni atanan çizgi tipini ancak REGEN'de
gösteriyor; renk hemen değişse de kesikler LTSCALE değiştirilene (o da REGEN yapar) kadar düz görünüyordu. 1.13'teki
`Editor.Regen()` modeless formun butonundan çalışmıyordu. Artık ÇALIŞTIR / YÜKLE sonrası ve OTOSTIL, OTOIC, OTOOLCU
sonunda `REGEN` AutoCAD'in komut kuyruğuna gönderiliyor (SendStringToExecute).

**1.15: Görünüş tamlık kontrolü.** Bütün elemanlar önden ve yandan izdüşürülüp (PLS-TOWER görünüşü) her birinin
çizimde karşılığı olduğu doğrulandı (`test/gorunus_testi.sh`). Ön görünüş 7 kulede %100 (YAA-R1 1210/1210, Tangent
1068/1068). Yan görünüşte OutlineDrawing'in atladığı Tangent Tower en üst paneli (z 46000–47800: L13, H1, CB-D1;
önünde uçtan görülen travers var) artık iç eleman olarak çiziliyor; yan görünüşte eksik kalan yalnız gövde dışına
taşan, uçtan görülen travers elemanları (bilerek, okunmaz yığın olmasın diye).

**1.14: Görünüşlere iç elemanlar (PLS-TOWER görünüşündeki gibi).** "İç eleman" kutucuğu (varsayılan işaretli) veya
`OTOIC`: ön/yan görünüşe, yüz düzleminde olmayan elemanların izdüşümü eklenir (kalça çaprazları, iç diyaframlar; ör.
YAA-R1 20BE uzatmasındaki P-E1-5_20). Yalnız görünüşte yeni bir çizgi oluşturanlar eklenir: izdüşümü yüz
çizgilerinin üstüne binenler (arka yüz, yan yüzler, yatay düzlemler), ayna kopyaları, bakış yönüne yakın olanlar ve
gövdenin dışında kalanlar (uçtan görülen travers) eklenmez. Katman `OTO_IC_ELEMAN` (2.1: her ÇALIŞTIR'da açılır), stil ana/redundant
kuralıyla aynı. YAA-R1: ön 12, yan 80 çizgi; Tangent Tower: ön 4, yan 40.

**1.13: Kesik çizgi her ölçekte görünür.** Redundant çizgilerin çizgi tipi ölçeği LTSCALE'e ek olarak anotasyon
ölçeğini de (MSLTSCALE açıksa CANNOSCALE, ör. 1:100 → ×100) hesaba katar; önceden bu durumda 225 mm'lik desen
100 kat büyüyüp çizgiler düz görünüyordu. Stil sonrası REGEN yapılır; kullanılan ölçek komut satırına yazılır.

**1.12: Görünüş ölçüleri.** "Ölçü" işaretliyse (veya `OTOOLCU`) TRANSVERSE ve LONGITUDINAL FACE'e de ölçü konur:
sağda PLS-TOWER kısımlarının (Leg Ext., Body Ext., Basic Body, Crossarms, Pikes …) sınır kotları zincir ölçü ve toplam
yükseklik; görünüş başlığının altında taban genişliği; kol ve tepe uçlarında eksenden uca yatay ölçü (en üst uçta
üstte, diğerlerinde ucun altında). Görünüşün yeri (ofset 0 / 50000) ve hali (çizildiği gibi / 2D Aktar sonrası)
çizgilerinden bulunur; ölçüler başlıkla aynı düzlemde, `OTO_KESIT_OLCU` katmanında.

**1.11: Redundant tanımı PLS-TOWER grup tipine göre.** Grup tablosundaki "group type" sütunu 3 (Redundant,
PLS-TOWER çiziminde turuncu) olan gruplar da kesikli/mavi çizilir; önceden yalnız açıklamasında "Redundant" geçen
gruplar sayılıyordu. YAA-R1'de açıklaması "Crossarm Brace", "Basic Body Inner Horizontal", "... Rdnt" olan
60 redundant grup ana eleman çiziliyordu: redundant eleman 246 → 538 (TRANSVERSE FACE'te 53 → 101 çizgi).
OutlineDrawing bu sütunu okumadığı için eklenti .tow'dan kendisi okur. Tangent Tower ve uzatmalarında değişiklik yok.

**1.10:**
- Bir kottaki elemanlar birbirinin aynası olan kopuk parçalardan oluşuyorsa (iki toprak teli tepesinin kapağı,
  dört bacak köşesi) yükseklik kesiti yerine yalnız bir parçayı (+X / -Y tarafı) çevreleyen yatay düzlem kesiti
  (4 nokta) önerilir. YAA-R1'de A (tepe kapağı 829 × 610) ve V (bacak köşesi) böyle; görünmeyen kapsaması aynı.
- Ölçü aralığı ve yazısı kesit büyüklüğüne göre (büyük kesitte 500 / 200 mm, küçükte en az 150 / 100 mm).

**1.9 (YAA-R1 kulesiyle):**
- Bacakların iç yüzleri ve kiriş iç çerçeveleri gibi dik/düşey düzlemler (75°'den dik) kesit alır, ama yalnız
  elemanlarının en az yarısı ön/yan görünüşte görünmüyorsa. Dış yüzler (zaten görünüşte olanlar) kesit almaz.
- Kesişim noktasında düğüm olmayan X çaprazlı yüzler (yatay elemansız) artık kafes sayılıyor.
  YAA-R1: görünmeyen 422 elemanın 418'i kapsanıyor (önce 372).
- Ölçü: uzatma çizgileri sabit boy (DIMFXLON, 400 mm); kesitteki küçük kopuk parçalar (ör. kol ucundaki tek kısa
  eleman) zincir ölçüye katılmıyor.
- Uyarı satırı alt satıra kayıyor, tam metin ipucu balonunda; mesajlar kısaldı, ayrıntı komut satırında.

**1.8: Kesitler arası eleman paylaşımı düzeltildi.** `SectionDetection` bir elemanı yalnız son kesitte çiziyordu
(F kesitinin iki H5 elemanı G'ye gidiyordu). Anahtar artık "kesit§eleman"; ayrıntı `STABILITE.md` A1b.
Doğrulama: `bash test/kesit_testi.sh eklenti/bin/tek/OutlineDrawing.dll kule.tow`.

**1.7: OutlineDrawing düzeltmesi dahil.** `FinalizeMember` simetri kodu 12 olan elemanların aynalarını çizmiyordu
(Tangent Tower'da sol alt traverste 14 eleman eksikti). Tek DLL'de `FinalizeMember`'ın sonuna
`SimetriDuzeltici.Uygula` çağrısı eklendi (`eklenti/yama/OutlineDrawingYama.cs`, Mono.Cecil). Kod 12 için yalnız
X aynası üretilir, geometrisi zaten var olan atlanır. Ayrıntı: `STABILITE.md` A1.

Yeniden üretmek için (OutlineDrawing kaynak kodu değişip yeniden derlenirse gerekir):
`bash eklenti/tek_dll.sh /yol/OutlineDrawing.dll` → `eklenti/bin/tek/OutlineDrawing.dll`.
Tek DLL derlemesinde eklentinin `[assembly: CommandClass]` satırı kapatılır (`TEK_DLL`), yoksa AutoCAD yalnız
eklentinin komut sınıfını tarar ve DRAWOUTLINE kaybolur; `ExtensionApplication` özniteliği `/copyattrs` ile taşınır.

## Ayrı eklenti DLL ile kurulum (OutlineDrawing.dll'e dokunmadan)

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
  cyan (4), redundant elemanlar kesikli (DASHED) ve mavi (5) olur. Redundant: PLS-TOWER grup tipi Redundant
  (grup tablosunda 3, PLS'te turuncu) veya açıklamasında "Redundant" geçen gruplar.
  Çizilen her çizgi, grid'deki kesit tanımıyla OutlineDrawing'in seçim kuralına göre bulunan elemanlarla
  rijit dönüşümle (döndürme, öteleme, aynalama) eşleştirilir. Ön/yan görünüş çizgileri ise OutlineDrawing'in
  frontFace/sideFace koordinatlarından birebir tanınır (2D Aktar sonrası da). Çizginin katmanı değişmez. Elle: `OTOSTIL`.
  YÜKLE ile çizilen 3D modelde de uygulanır. Eleman isimleri (yükseklik < 250 olan, "SECTION x" / "0" katmanındaki
  yazılar) tek renk (7) yapılır. Eğik kesitte düzlemin önünde/arkasında kalan çizgiler de kesite dahildir;
  tanınamayan çizgi kalırsa ana eleman kabul edilir (tek renk) ve UYARI satırında sayısı yazılır.
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

Bütün sistem (önerilen, her değişiklikten sonra; AutoCAD gerekmez, Mono yeterli):

```
bash eklenti/tek_dll.sh /yol/OutlineDrawingv1.3.dll                     # tek DLL'i üret
bash test/hepsi.sh eklenti/bin/tek/OutlineDrawing.dll kule1.tow kule2.tow # her .tow + 12 yapay varyantı (paralel)
```

Her dosyada beş test (ayrıntı STABILITE.md C): `detay_testi.sh` (okuma = bağımsız okuyucu, kapsama, her eleman
etiketli çiziliyor mu), `kesit_testi.sh` (gerçek SectionDetection, eşleşme, stil), `gorunus_testi.sh` (izdüşüm),
`gorunus_konum_testi.sh` (görünüş yeri), `olcu_testi.sh` (ölçü yazıları sığıyor / çakışmıyor mu). Yapay varyantlar `test/tow_varyant.py` ile üretilir (döndürme, dikdörtgen
gövde, eksi kot, yeni grup sütunu, Türkçe ad, uzun ad, büyük kule, boşluksuz .blt yolu, yeni kesit etiketi sütunu,
yalnız Türkçe harfte farklı eleman adları; ayrıca bilerek bozuk dosyalar: `eksikdugum`, `xarm`, `grupeksik`).
Ayrıca bir kez `ayna_parca_testi.sh` (sentetik, ayna olmayan kopuk parça).

Eski / tek tek testler:

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
