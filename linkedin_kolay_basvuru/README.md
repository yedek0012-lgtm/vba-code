# Kolay Başvuru Doldurucu (LinkedIn Easy Apply)

LinkedIn'deki **Kolay Başvuru** pencerelerini, bir kez girdiğiniz profil bilgileriyle
otomatik dolduran bir Chrome / Edge / Brave eklentisi.

- Ad, soyad, e-posta, telefon, şehir, LinkedIn/GitHub/portföy bağlantıları
- "X ile kaç yıl deneyiminiz var?" soruları (beceri bazında yıl tablosu)
- Çalışma izni, vize sponsorluğu, taşınma, hibrit/ofis, uzaktan, diploma, ehliyet, İngilizce seviyesi
- Maaş beklentisi, ihbar süresi, ön yazı, özet
- Açılır listeler, Evet/Hayır seçimleri, "Kabul ediyorum" kutuları, şehir önerileri
- Hem İngilizce hem Türkçe LinkedIn arayüzü

## Güvenlik ilkesi: göndermeyi siz yaparsınız

Eklenti **"Başvuruyu gönder" butonuna asla basmaz.** Formu doldurur, isterseniz adım adım
ilerler ve son (inceleme) adımında durup Gönder butonunu mavi çerçeveyle işaretler.
Kontrol edip göndermek size kalır. LinkedIn otomatik toplu başvuru botlarını kullanım
koşullarında yasaklar ve hesap kısıtlayabilir; bu eklenti bu yüzden yalnızca form doldurma
yardımcısı olarak tasarlandı. Kendi sorumluluğunuzda, makul hızda kullanın.

## Kurulum

1. Bu klasörü (`linkedin_kolay_basvuru`) bilgisayarınıza indirin.
2. Chrome'da `chrome://extensions` adresini açın (Edge: `edge://extensions`).
3. Sağ üstten **Geliştirici modu**nu açın.
4. **Paketlenmemiş öğe yükle** → `linkedin_kolay_basvuru` klasörünü seçin.
5. Eklenti simgesine tıklayın → **Profil ve ayarlar** → bilgilerinizi girip **Kaydet**.

## Kullanım

1. LinkedIn'de bir ilanı açıp **Kolay Başvuru** butonuna basın.
2. Sağ altta mavi panel çıkar:
   - **⚡ Doldur** – yalnızca açık adımı doldurur.
   - **⏩ Otomatik** – doldurur, *İleri / İncele* butonlarına basarak ilerler, son adımda durur.
   - **■** – durdurur.
3. Aynı işlemler eklenti simgesindeki açılır pencereden de yapılabilir.

Renkler: **yeşil** = eklenti doldurdu, **turuncu** = cevap bulunamadı (siz doldurun).

### Öğrenme

Turuncu işaretli bir soruyu elle cevaplayıp *İleri*ye bastığınızda cevap **Özel cevaplar**
listesine kaydedilir; aynı soru bir daha geldiğinde otomatik doldurulur. Cevaplanamayan
sorular ayarlar sayfasındaki **Cevaplanamayan sorular** bölümünde de birikir; oradan toplu
cevap verebilirsiniz.

### Özel cevaplar

Ayarlar sayfasında *desen → cevap* çiftleri ekleyebilirsiniz. Desen, soru metninde geçen bir
parçadır (Türkçe karakter ve büyük/küçük harf farkı önemsizdir) ya da `/regex/` biçiminde
düzenli ifadedir. Özel cevaplar yerleşik kurallardan önce gelir.

Örnek:

| Desen | Cevap |
|---|---|
| `askerlik` | `Tamamlandı` |
| `/hafta sonu.*calis/` | `Yes` |
| `referans` | `Yok` |

### Yedekleme

**Dışa aktar (JSON)** ile profilinizi dosyaya kaydedip başka bir bilgisayarda **İçe aktar**
ile yükleyebilirsiniz. Veriler yalnızca tarayıcınızda (`chrome.storage.local`) tutulur.

## Bilinen sınırlamalar

- CV dosyası yüklenmez; LinkedIn'in kayıtlı CV seçimi olduğu gibi kullanılır.
- LinkedIn arayüzünü değiştirirse bazı alanlar tanınmayabilir. Bu durumda alan turuncu
  işaretlenir; elle doldurduğunuzda cevap öğrenilir.
- Sadece LinkedIn'in kendi Kolay Başvuru penceresinde çalışır, şirket sitelerine yönlenen
  başvurularda çalışmaz.

## Dosyalar

| Dosya | Görev |
|---|---|
| `manifest.json` | Eklenti tanımı (Manifest V3) |
| `src/matcher.js` | Soru metnini profil bilgisine eşleyen kurallar ve seçenek seçimi |
| `src/content.js` | LinkedIn sayfasında alanları bulur, doldurur, ilerler, öğrenir |
| `src/panel.css` | Sayfadaki mavi panelin stili |
| `src/popup.*` | Eklenti simgesine tıklayınca açılan pencere |
| `src/options.*` | Profil ve ayarlar sayfası |
| `tests/` | Birim testi ve sahte Easy Apply penceresiyle uçtan uca test |

## Testler

```bash
node tests/matcher.test.js
```

```bash
node tests/e2e.test.js
```

`e2e.test.js` için Playwright ve Chromium gerekir (`npm i -D playwright && npx playwright install chromium`).
