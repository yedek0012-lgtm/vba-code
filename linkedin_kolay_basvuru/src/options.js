(function () {
  "use strict";
  const M = KBMatcher;
  const $ = s => document.querySelector(s);
  const $$ = s => [...document.querySelectorAll(s)];
  const KEYS = ["profile", "customAnswers", "unanswered", "settings"];
  let data = M.withDefaults({});

  function get() { return new Promise(r => chrome.storage.local.get(KEYS, d => r(d || {}))); }
  function set(o) { return new Promise(r => chrome.storage.local.set(o, () => r())); }

  function flash(t) {
    $("#msg").textContent = t;
    setTimeout(() => { $("#msg").textContent = ""; }, 2500);
  }

  // ---- Ozel cevaplar tablosu ----
  function addCustomRow(pattern, answer, learned) {
    const tr = document.createElement("tr");
    tr.innerHTML =
      '<td><input class="pat"></td><td><input class="ans"></td>' +
      '<td><button type="button" class="icon" title="Sil">✕</button></td>';
    tr.querySelector(".pat").value = pattern || "";
    tr.querySelector(".ans").value = answer || "";
    if (learned) tr.querySelector(".pat").title = "🧠 Elle girdiğiniz cevaptan öğrenildi";
    if (learned) tr.dataset.learned = "1";
    tr.querySelector("button").onclick = () => tr.remove();
    $("#custom tbody").appendChild(tr);
  }

  function readCustom() {
    return $$("#custom tbody tr")
      .map(tr => ({
        pattern: tr.querySelector(".pat").value.trim(),
        answer: tr.querySelector(".ans").value.trim(),
        learned: tr.dataset.learned === "1" || undefined
      }))
      .filter(c => c.pattern);
  }

  // ---- Cevaplanamayan sorular ----
  function renderUnanswered() {
    const box = $("#unanswered");
    box.innerHTML = "";
    $("#unCount").textContent = data.unanswered.length || "";
    if (!data.unanswered.length) {
      box.innerHTML = '<p class="muted">Şu an bekleyen soru yok. 🎉</p>';
      return;
    }
    data.unanswered.slice().reverse().forEach(u => {
      const div = document.createElement("div");
      div.className = "q";
      const opts = (u.options || []).filter(o => o && !/^(select an option|seçenek|bir seçenek)/i.test(o));
      div.innerHTML = '<div class="label"></div><div class="opts"></div><div class="row"></div>';
      div.querySelector(".label").textContent = u.label;
      div.querySelector(".opts").textContent = opts.length ? "Seçenekler: " + opts.join(" · ") : (u.kind === "number" ? "Sayısal alan" : "");
      const row = div.querySelector(".row");
      let input;
      if (opts.length) {
        input = document.createElement("select");
        ["", ...opts].forEach(o => { const op = document.createElement("option"); op.textContent = o; input.appendChild(op); });
      } else {
        input = document.createElement("input");
        input.placeholder = "Cevabınız";
      }
      const save = document.createElement("button");
      save.type = "button"; save.className = "secondary"; save.textContent = "Kaydet";
      save.onclick = async () => {
        const v = input.value.trim();
        if (!v) return;
        const pattern = M.norm(u.label);
        data.customAnswers = readCustom().filter(c => M.norm(c.pattern) !== pattern);
        data.customAnswers.push({ pattern, answer: v });
        data.unanswered = data.unanswered.filter(x => x !== u);
        await set({ customAnswers: data.customAnswers, unanswered: data.unanswered });
        render();
        flash("Cevap kaydedildi");
      };
      const del = document.createElement("button");
      del.type = "button"; del.className = "icon"; del.title = "Yok say"; del.textContent = "✕";
      del.onclick = async () => {
        data.unanswered = data.unanswered.filter(x => x !== u);
        await set({ unanswered: data.unanswered });
        renderUnanswered();
      };
      row.append(input, save, del);
      box.appendChild(div);
    });
  }

  // ---- Form <-> veri ----
  function render() {
    $$("[data-p]").forEach(el => { el.value = data.profile[el.dataset.p] == null ? "" : data.profile[el.dataset.p]; });
    $("#experience").value = M.formatExperience(data.profile.experience);
    $$("[data-s]").forEach(el => {
      const v = data.settings[el.dataset.s];
      if (el.type === "checkbox") el.checked = !!v; else el.value = v;
    });
    $("#custom tbody").innerHTML = "";
    data.customAnswers.forEach(c => addCustomRow(c.pattern, c.answer, c.learned));
    renderUnanswered();
  }

  function collect() {
    $$("[data-p]").forEach(el => { data.profile[el.dataset.p] = el.value.trim(); });
    data.profile.experience = M.parseExperience($("#experience").value);
    $$("[data-s]").forEach(el => {
      data.settings[el.dataset.s] = el.type === "checkbox" ? el.checked : (el.type === "number" ? Number(el.value) || 0 : el.value);
    });
    data.customAnswers = readCustom();
  }

  $("#save").onclick = async () => {
    collect();
    await set({ profile: data.profile, customAnswers: data.customAnswers, settings: data.settings });
    flash("Kaydedildi ✓");
  };
  $("#addCustom").onclick = () => addCustomRow("", "");
  $("#clearUn").onclick = async () => {
    data.unanswered = [];
    await set({ unanswered: [] });
    renderUnanswered();
  };

  $("#export").onclick = () => {
    collect();
    const blob = new Blob([JSON.stringify({ profile: data.profile, customAnswers: data.customAnswers, settings: data.settings }, null, 2)], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "kolay-basvuru-profil.json";
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  };

  $("#import").onchange = async e => {
    const f = e.target.files[0];
    if (!f) return;
    try {
      const imported = JSON.parse(await f.text());
      data = M.withDefaults(Object.assign({}, data, imported));
      await set({ profile: data.profile, customAnswers: data.customAnswers, settings: data.settings });
      render();
      flash("İçe aktarıldı ✓");
    } catch (err) {
      alert("Dosya okunamadı: " + err.message);
    }
    e.target.value = "";
  };

  // Baska sekmede ogrenilen cevaplar gelirse listeyi tazele
  chrome.storage.onChanged.addListener((changes, area) => {
    if (area !== "local") return;
    if (changes.unanswered) { data.unanswered = changes.unanswered.newValue || []; renderUnanswered(); }
  });

  get().then(d => { data = M.withDefaults(d); render(); });
})();
