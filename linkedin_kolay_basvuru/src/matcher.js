/*
 * KBMatcher - soru metnini profil bilgisine eslestiren saf mantik.
 * DOM'a dokunmaz; hem tarayicida (content script) hem Node testlerinde calisir.
 */
(function (root) {
  "use strict";

  // ---- Varsayilan veri ----------------------------------------------------
  const DEFAULTS = {
    profile: {
      firstName: "",
      lastName: "",
      email: "",
      phoneCountry: "Turkey (+90)",
      phone: "",
      city: "",
      linkedin: "",
      website: "",
      github: "",
      currentTitle: "",
      currentCompany: "",
      salary: "",
      noticePeriod: "",
      defaultYears: "1",
      experience: {}, // { "python": "3", "excel": "5" }
      workAuthorization: "Yes",
      sponsorship: "No",
      relocate: "Yes",
      commute: "Yes",
      remote: "Yes",
      degree: "Yes",
      englishLevel: "Professional",
      driversLicense: "Yes",
      military: "",
      gender: "",
      coverLetter: "",
      summary: ""
    },
    customAnswers: [], // [{ pattern: "metin veya /regex/", answer: "..." }]
    unanswered: [], // [{ label, kind, options, url, ts }]
    settings: {
      autoAdvance: false,
      overwrite: false,
      highlightUnknown: true,
      learnAnswers: true,
      unfollowCompany: true,
      delayMs: 700
    }
  };

  // ---- Metin normallestirme -----------------------------------------------
  function norm(s) {
    return String(s == null ? "" : s)
      .replace(/İ/g, "i")
      .replace(/I/g, "i")
      .toLowerCase()
      .replace(/ı/g, "i")
      .normalize("NFD")
      .replace(/[̀-ͯ]/g, "")
      .replace(/[^a-z0-9+#.]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  const YES = ["yes", "evet", "true", "1"];
  const NO = ["no", "hayir", "false", "0"];
  const DECLINE_RE =
    /(decline|prefer not|not to (say|answer|disclose)|don ?t wish|do not wish|belirtmek istemiyorum|cevaplamak istemiyorum|soylemek istemiyorum)/;
  const PLACEHOLDER_RE =
    /^(select an option|select|choose|seciniz|secin|bir secenek secin|secenek secin|lutfen secin|please select|-+)$/;

  function isYes(v) { return YES.includes(norm(v)); }
  function isNo(v) { return NO.includes(norm(v)); }

  // ---- Deneyim yili --------------------------------------------------------
  function escapeRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }

  /** Soru metninde gecen beceriyi deneyim tablosunda arar (en uzun eslesme kazanir). */
  function findSkill(label, experience) {
    const n = " " + norm(label) + " ";
    let best = null;
    for (const key of Object.keys(experience || {})) {
      const k = norm(key);
      if (!k) continue;
      const re = new RegExp("(^|[^a-z0-9])" + escapeRe(k) + "([^a-z0-9]|$)");
      if (re.test(n) && (!best || k.length > best.k.length)) best = { k, key };
    }
    return best ? { skill: best.key, years: String(experience[best.key]) } : null;
  }

  // ---- Kurallar -------------------------------------------------------------
  // Sira onemli: daha ozel kurallar once gelir.
  const RULES = [
    { id: "phoneCountry", re: /(country code|ulke kodu|phone country|alan kodu)/, val: p => p.phoneCountry },
    { id: "email", re: /(e ?mail|e posta|eposta)/, val: p => p.email },
    { id: "phone", re: /(phone|telefon|mobile|cep no|cep numara|gsm)/, val: p => p.phone },
    { id: "fullName", re: /(full name|ad soyad|adiniz soyadiniz|isim soyisim|ad ve soyad)/, val: p => [p.firstName, p.lastName].filter(Boolean).join(" ") },
    { id: "firstName", re: /(first name|given name|^(ad|adi|adiniz|isim|isminiz)$)/, val: p => p.firstName },
    { id: "lastName", re: /(last name|surname|family name|soyad|soyisim)/, val: p => p.lastName },
    { id: "linkedin", re: /linkedin/, val: p => p.linkedin },
    { id: "github", re: /github/, val: p => p.github },
    { id: "website", re: /(website|web site|portfolio|portfoy|personal site|kisisel site)/, val: p => p.website },
    {
      id: "years",
      re: /(how many years|years of|years experience|kac yil|yil deneyim|yillik deneyim|yil tecrube|kac yillik)/,
      val: (p, label) => {
        const hit = findSkill(label, p.experience);
        return hit ? hit.years : p.defaultYears;
      }
    },
    {
      id: "hasExperience",
      re: /(experience (with|in|using)|deneyiminiz var|deneyimin var|tecrubeniz var|deneyimi var|ile calistiniz)/,
      val: (p, label) => (findSkill(label, p.experience) ? "Yes" : null)
    },
    { id: "salary", re: /(salary|maas|ucret beklenti|ucret talebi|compensation|ctc|pay expectation)/, val: p => p.salary },
    { id: "notice", re: /(notice period|ihbar|baslayabilir|start date|when can you start|available to start|ise baslama)/, val: p => p.noticePeriod },
    { id: "sponsorship", re: /(sponsor|vize|visa)/, val: p => p.sponsorship },
    { id: "workAuth", re: /(authori[sz]ed|legally|calisma izni|work permit|right to work|eligible to work|yasal olarak)/, val: p => p.workAuthorization },
    { id: "relocate", re: /(relocat|tasin)/, val: p => p.relocate },
    { id: "remote", re: /(remote|uzaktan)/, val: p => p.remote },
    { id: "commute", re: /(commut|on site|onsite|hybrid|hibrit|ofisten|ofise|office|ulasim|yerinde calis)/, val: p => p.commute },
    { id: "english", re: /(english|ingilizce)/, val: p => p.englishLevel },
    { id: "degree", re: /(bachelor|lisans|universite|degree|diploma|mezun)/, val: p => p.degree },
    { id: "license", re: /(driver|ehliyet|surucu belge)/, val: p => p.driversLicense },
    { id: "military", re: /(military|askerlik)/, val: p => p.military },
    { id: "gender", re: /(gender|cinsiyet)/, val: p => p.gender || "__decline__" },
    { id: "eeo", re: /(race|ethnic|veteran|disab|engellilik|engel durum|etnik)/, val: () => "__decline__" },
    { id: "title", re: /(current title|job title|current position|unvan|mevcut pozisyon)/, val: p => p.currentTitle },
    { id: "company", re: /(current company|current employer|employer|mevcut sirket|calistiginiz sirket|isveren)/, val: p => p.currentCompany },
    { id: "cover", re: /(cover letter|on yazi|ust yazi|niyet mektubu|message to the hiring|why do you want|why are you interested|neden bu)/, val: p => p.coverLetter },
    { id: "summary", re: /(summary|headline|ozet|hakkinizda|about yourself|kendinizden)/, val: p => p.summary },
    { id: "city", re: /(city|sehir|location|konum|ikamet|yasadiginiz)/, val: p => p.city, typeahead: true }
  ];

  // ---- Ozel cevaplar ---------------------------------------------------------
  function customMatch(label, customAnswers) {
    const n = norm(label);
    let best = null;
    for (const ca of customAnswers || []) {
      if (!ca || !ca.pattern) continue;
      const pat = String(ca.pattern).trim();
      let ok = false;
      const m = pat.match(/^\/(.+)\/([a-z]*)$/);
      if (m) {
        try { ok = new RegExp(m[1], m[2] || "i").test(n) || new RegExp(m[1], m[2] || "i").test(label); }
        catch (e) { ok = false; }
      } else {
        const np = norm(pat);
        ok = np.length > 0 && n.includes(np);
      }
      // En uzun (en ozel) desen kazanir.
      if (ok && (!best || pat.length > best.pattern.length)) best = ca;
    }
    return best;
  }

  /**
   * Bir soru icin cevap bulur.
   * @returns {{value:string, source:string, ruleId?:string, typeahead?:boolean} | null}
   */
  function resolve(label, data) {
    const profile = Object.assign({}, DEFAULTS.profile, (data && data.profile) || {});
    const ca = customMatch(label, data && data.customAnswers);
    if (ca && String(ca.answer).length) return { value: String(ca.answer), source: "custom" };

    const n = norm(label);
    if (!n) return null;
    for (const r of RULES) {
      if (!r.re.test(n)) continue;
      const v = r.val(profile, label);
      if (v == null || String(v).trim() === "") return null; // kural tutuyor ama profilde bilgi yok
      return { value: String(v), source: "rule", ruleId: r.id, typeahead: !!r.typeahead };
    }
    return null;
  }

  // ---- Secenek secimi --------------------------------------------------------
  /**
   * @param {{text:string, value?:string}[]} options
   * @param {string} answer
   * @returns {number} secilecek indeks, bulunamazsa -1
   */
  function pickOption(options, answer) {
    const opts = options.map((o, i) => ({ i, n: norm(o.text), raw: String(o.text || "").toLowerCase() })).filter(o => o.n && !PLACEHOLDER_RE.test(o.n));
    if (!opts.length) return -1;
    const a = norm(answer);

    if (answer === "__decline__") {
      const d = opts.find(o => DECLINE_RE.test(o.n));
      return d ? d.i : -1;
    }
    if (isYes(answer)) {
      const y = opts.find(o => YES.includes(o.n) || /^(yes|evet)\b/.test(o.n));
      if (y) return y.i;
    }
    if (isNo(answer)) {
      const x = opts.find(o => NO.includes(o.n) || /^(no|hayir)\b/.test(o.n));
      if (x) return x.i;
    }
    let hit = opts.find(o => o.n === a);
    if (hit) return hit.i;
    hit = opts.find(o => a && o.n.startsWith(a));
    if (hit) return hit.i;
    hit = opts.find(o => a && o.n.includes(a));
    if (hit) return hit.i;
    hit = opts.find(o => o.n.length > 2 && a.includes(o.n));
    if (hit) return hit.i;

    // Sayisal cevap: "3-5 yil", "5+", "10 veya daha fazla" gibi araliklar
    const num = parseFloat(String(answer).replace(",", "."));
    if (!isNaN(num)) {
      for (const o of opts) {
        const t = norm(o.raw.replace(/[-–—]/g, " to ").replace(/\+/g, " plus "));
        const range = t.match(/(\d+(?:\.\d+)?)\s*(?:to|ile)\s*(\d+(?:\.\d+)?)/);
        if (range && num >= +range[1] && num <= +range[2]) return o.i;
        const plus = t.match(/(\d+(?:\.\d+)?)\s*(?:plus|or more|veya (?:daha )?fazla|ve uzeri|uzeri)/);
        if (plus && num >= +plus[1]) return o.i;
        const less = t.match(/(?:less than|under|en fazla|alti)\s*(\d+(?:\.\d+)?)|(\d+(?:\.\d+)?)\s*(?:dan az|den az|alti)/);
        if (less && num < +(less[1] || less[2])) return o.i;
      }
    }
    return -1;
  }

  // ---- Diger yardimcilar -----------------------------------------------------
  /** "python = 3" satirlarini nesneye cevirir. */
  function parseExperience(text) {
    const out = {};
    String(text || "").split(/\r?\n/).forEach(line => {
      const m = line.match(/^\s*(.+?)\s*[=:]\s*([\d.,]+)\s*$/);
      if (m) out[m[1].trim()] = m[2].replace(",", ".");
    });
    return out;
  }
  function formatExperience(obj) {
    return Object.keys(obj || {}).map(k => k + " = " + obj[k]).join("\n");
  }

  /** Ayarlari varsayilanlarla birlestirir. */
  function withDefaults(data) {
    data = data || {};
    return {
      profile: Object.assign({}, DEFAULTS.profile, data.profile || {}),
      customAnswers: Array.isArray(data.customAnswers) ? data.customAnswers : [],
      unanswered: Array.isArray(data.unanswered) ? data.unanswered : [],
      settings: Object.assign({}, DEFAULTS.settings, data.settings || {})
    };
  }

  /** Sayisal alan icin degeri temizler ("25.000 TL" -> "25000"). */
  function toNumeric(v) {
    const s = String(v).replace(/\s/g, "");
    const m = s.match(/\d+(?:[.,]\d+)*/);
    if (!m) return "";
    let t = m[0];
    // binlik ayiricilari temizle: 25.000 / 25,000 -> 25000 ; 2.5 -> 2.5
    if (/^\d{1,3}([.,]\d{3})+$/.test(t)) t = t.replace(/[.,]/g, "");
    else t = t.replace(",", ".");
    return t;
  }

  const api = {
    DEFAULTS, RULES, norm, resolve, pickOption, customMatch, findSkill,
    parseExperience, formatExperience, withDefaults, toNumeric, isYes, isNo, DECLINE_RE
  };
  root.KBMatcher = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
