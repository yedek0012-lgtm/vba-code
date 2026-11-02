// Calistir: node tests/matcher.test.js
const assert = require("node:assert/strict");
const M = require("../src/matcher.js");

const data = M.withDefaults({
  profile: {
    firstName: "Ayşe", lastName: "Yılmaz", email: "ayse@example.com", phone: "5551112233",
    city: "İstanbul, Türkiye", salary: "60.000 TL", defaultYears: "1",
    experience: { python: "3", "c#": "2", excel: "5", "power bi": "4" },
    englishLevel: "Professional", military: "Tamamlandı"
  },
  customAnswers: [
    { pattern: "hangi takimi tutuyorsunuz", answer: "Fark etmez" },
    { pattern: "/referans.*kod/", answer: "REF-42" }
  ]
});

let n = 0;
function t(name, fn) { fn(); n++; }
const val = l => (M.resolve(l, data) || {}).value;
const rule = l => (M.resolve(l, data) || {}).ruleId;

t("norm Turkce", () => {
  assert.equal(M.norm("İŞ DENEYİMİ (Yıl)*"), "is deneyimi yil");
  assert.equal(M.norm("Şehir / Çalışma"), "sehir calisma");
});

t("temel alanlar", () => {
  assert.equal(val("First name"), "Ayşe");
  assert.equal(val("Ad"), "Ayşe");
  assert.equal(val("Soyadı"), "Yılmaz");
  assert.equal(val("Adres"), undefined); // "ad" ile karismamali
  assert.equal(val("Email address"), "ayse@example.com");
  assert.equal(val("Mobile phone number"), "5551112233");
  assert.equal(rule("Phone country code"), "phoneCountry");
  assert.equal(val("Ad Soyad"), "Ayşe Yılmaz");
  assert.equal(val("Location (city)"), "İstanbul, Türkiye");
});

t("deneyim yili", () => {
  assert.equal(val("How many years of work experience do you have with Python?"), "3");
  assert.equal(val("How many years of experience do you have with C#?"), "2");
  assert.equal(val("Power BI ile kaç yıllık iş deneyiminiz var?"), "4");
  assert.equal(val("How many years of experience do you have with Kubernetes?"), "1");
  assert.equal(val("Do you have experience with Excel?"), "Yes");
  assert.equal(M.resolve("Do you have experience with Rust?", data), null);
});

t("evet/hayir kurallari", () => {
  assert.equal(val("Will you now or in the future require sponsorship for employment visa status?"), "No");
  assert.equal(val("Are you legally authorized to work in Turkey?"), "Yes");
  assert.equal(rule("Are you comfortable commuting to this job's location?"), "commute");
  assert.equal(rule("Have you completed the following level of education: Bachelor's Degree?"), "degree");
  assert.equal(val("What is your level of proficiency in English?"), "Professional");
  assert.equal(val("Askerlik durumunuz nedir?"), "Tamamlandı");
  assert.equal(val("Gender"), "__decline__");
});

t("ozel cevaplar once gelir", () => {
  assert.equal(val("Hangi takımı tutuyorsunuz?"), "Fark etmez");
  assert.equal(val("Referans kodunuz varsa yazınız"), "REF-42");
  assert.equal(M.resolve("Tamamen bilinmeyen bir soru", data), null);
});

t("secenek secimi", () => {
  const yn = [{ text: "Select an option" }, { text: "Yes" }, { text: "No" }];
  assert.equal(M.pickOption(yn, "Yes"), 1);
  assert.equal(M.pickOption(yn, "Evet"), 1);
  assert.equal(M.pickOption(yn, "Hayır"), 2);
  const tr = [{ text: "Seçiniz" }, { text: "Evet" }, { text: "Hayır" }];
  assert.equal(M.pickOption(tr, "Yes"), 1);
  assert.equal(M.pickOption(tr, "No"), 2);
  const lvl = [{ text: "Select an option" }, { text: "None" }, { text: "Conversational" }, { text: "Professional" }, { text: "Native or bilingual" }];
  assert.equal(M.pickOption(lvl, "Professional"), 3);
  assert.equal(M.pickOption(lvl, "native"), 4);
  const g = [{ text: "Male" }, { text: "Female" }, { text: "Prefer not to say" }];
  assert.equal(M.pickOption(g, "__decline__"), 2);
  const yrs = [{ text: "0-1 years" }, { text: "2-4 years" }, { text: "5+ years" }];
  assert.equal(M.pickOption(yrs, "3"), 1);
  assert.equal(M.pickOption(yrs, "7"), 2);
  assert.equal(M.pickOption(yn, "Belki"), -1);
});

t("yardimcilar", () => {
  assert.equal(M.toNumeric("60.000 TL"), "60000");
  assert.equal(M.toNumeric("2,5 yıl"), "2.5");
  assert.equal(M.toNumeric("abc"), "");
  assert.deepEqual(M.parseExperience("python = 3\nExcel: 5\nbozuk satir"), { python: "3", Excel: "5" });
  assert.equal(M.formatExperience({ a: "1", b: "2" }), "a = 1\nb = 2");
});

console.log("matcher: " + n + " test grubu gecti");
