(function () {
  "use strict";
  const $ = id => document.getElementById(id);

  function activeTab() {
    return new Promise(res => chrome.tabs.query({ active: true, currentWindow: true }, t => res(t && t[0])));
  }
  function send(tab, msg) {
    return new Promise(res => chrome.tabs.sendMessage(tab.id, msg, r => {
      if (chrome.runtime.lastError) res(null); else res(r);
    }));
  }

  async function init() {
    chrome.storage.local.get(["profile", "unanswered"], d => {
      const data = KBMatcher.withDefaults(d);
      const warn = [];
      if (!data.profile.firstName || !data.profile.email) warn.push("Profil eksik: önce ayarlardan bilgilerinizi girin.");
      if (data.unanswered.length) warn.push(data.unanswered.length + " cevaplanmamış soru ayarlarda sizi bekliyor.");
      $("warn").textContent = warn.join(" ");
    });

    const tab = await activeTab();
    if (!tab || !/^https:\/\/www\.linkedin\.com\//.test(tab.url || "")) {
      $("status").textContent = "LinkedIn iş ilanı sayfasında açın.";
      $("fill").disabled = $("auto").disabled = true;
      return;
    }
    const st = await send(tab, { type: "kb-status" });
    if (!st) $("status").textContent = "Sayfayı yenileyin (eklenti henüz yüklenmedi).";
    else $("status").textContent = st.dialog ? "Kolay Başvuru penceresi açık." : "Önce ilandaki 'Kolay Başvuru' butonuna basın.";

    async function go(autoAdvance) {
      $("status").textContent = "Dolduruluyor…";
      const r = await send(tab, { type: "kb-fill", autoAdvance });
      $("status").textContent = r ? r.message : "Sayfayla bağlantı kurulamadı; sayfayı yenileyin.";
    }
    $("fill").onclick = () => go(false);
    $("auto").onclick = () => go(true);
  }

  $("opts").onclick = e => { e.preventDefault(); chrome.runtime.openOptionsPage(); };
  init();
})();
