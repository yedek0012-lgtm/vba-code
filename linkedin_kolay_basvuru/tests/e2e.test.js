// Calistir: node tests/e2e.test.js   (Playwright + Chromium gerekir)
// Eklentinin content script'ini sahte bir Easy Apply penceresinde calistirir.
const path = require("path");
const assert = require("node:assert/strict");
const { chromium } = require("playwright");

const src = p => path.join(__dirname, "..", "src", p);
const chromeStub = `
  window.__store = {
    profile: { firstName: "Ayşe", lastName: "Yılmaz", email: "ayse@example.com", phone: "5551112233",
               phoneCountry: "Turkey (+90)", city: "İstanbul", salary: "60.000 TL",
               experience: { python: "3" }, englishLevel: "Professional" },
    settings: { delayMs: 50 }
  };
  window.chrome = {
    storage: { local: {
      get: (keys, cb) => cb(JSON.parse(JSON.stringify(window.__store))),
      set: (obj, cb) => { Object.assign(window.__store, JSON.parse(JSON.stringify(obj))); cb && cb(); }
    } },
    runtime: { onMessage: { addListener() {} } }
  };`;

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage();
  await page.goto("file://" + path.join(__dirname, "mock_easy_apply.html"));
  await page.addScriptTag({ content: chromeStub });
  await page.addScriptTag({ path: src("matcher.js") });
  await page.addScriptTag({ path: src("content.js") });
  await page.addStyleTag({ path: src("panel.css") });
  await page.waitForTimeout(600);
  assert.equal(await page.isVisible("#kb-panel"), true, "panel gorunmeli");

  // 1) Otomatik mod: 1. adimi doldurur, ilerler, 2. adimdaki bilinmeyen zorunlu soruda durur
  let r = await page.evaluate(() => window.__kb.run({ autoAdvance: true }));
  console.log("1. calisma:", r.message);
  const v = await page.evaluate(() => ({
    fn: fn.value, ln: ln.value, em: em.value, cc: cc.value, ph: ph.value, city: city.value,
    py: py.value, sal: document.getElementById("sal-numeric").value,
    auth: document.querySelector("[name=auth]:checked")?.value, sp: document.querySelector("[name=sp]:checked")?.value,
    en: en.value, agree: agree.checked, odd: odd.getAttribute("data-kb-unknown"),
    events: window.events, unanswered: window.__store.unanswered.map(u => u.label)
  }));
  assert.deepEqual(
    [v.fn, v.ln, v.em, v.cc, v.ph, v.city],
    ["Ayşe", "Yılmaz", "a", "tr", "5551112233", "Istanbul, Türkiye (secildi)"]);
  assert.deepEqual([v.py, v.sal, v.auth, v.sp, v.en, v.agree], ["3", "60000", "Yes", "No", "Professional", true]);
  assert.equal(v.odd, "Favori programlama diliniz nedir?");
  assert.deepEqual(v.events, ["next"]);
  assert.deepEqual(v.unanswered, ["Favori programlama diliniz nedir?"]);

  // 2) Kullanici bilinmeyen soruyu elle doldurur, tekrar calistirir -> ogrenir, inceleme adiminda durur
  await page.fill("#odd", "VBA");
  r = await page.evaluate(() => window.__kb.run({ autoAdvance: true }));
  console.log("2. calisma:", r.message);
  const w = await page.evaluate(() => ({
    events: window.events, follow: document.getElementById("follow-company-checkbox").checked,
    custom: window.__store.customAnswers, unanswered: window.__store.unanswered
  }));
  assert.deepEqual(w.events, ["next", "review"], "Gonder'e ASLA basilmamali");
  assert.equal(w.follow, false, "sirket takibi kaldirilmali");
  assert.deepEqual(w.custom.map(c => [c.pattern, c.answer]), [["favori programlama diliniz nedir", "VBA"]]);
  assert.equal(w.unanswered.length, 0);
  await page.screenshot({ path: path.join(__dirname, "e2e_son_adim.png") });

  await browser.close();
  console.log("e2e: tum kontroller gecti");
})().catch(e => { console.error(e); process.exit(1); });
