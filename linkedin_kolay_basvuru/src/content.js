/*
 * LinkedIn Kolay Basvuru (Easy Apply) penceresindeki alanlari doldurur.
 * Basvuruyu ASLA kendisi gondermez: son "Gonder" tiklamasi kullaniciya aittir.
 */
(function () {
  "use strict";
  if (window.__kbLoaded) return;
  window.__kbLoaded = true;

  const M = globalThis.KBMatcher;
  const STORE_KEYS = ["profile", "customAnswers", "unanswered", "settings"];
  const UNKNOWN_ATTR = "data-kb-unknown";
  const FILLED_ATTR = "data-kb-filled";

  let data = M.withDefaults({});
  let running = false;
  let stopRequested = false;

  // ---- Depolama ---------------------------------------------------------------
  function storageGet() {
    return new Promise(res => chrome.storage.local.get(STORE_KEYS, r => res(r || {})));
  }
  function storageSet(obj) {
    return new Promise(res => chrome.storage.local.set(obj, () => res()));
  }
  async function load() { data = M.withDefaults(await storageGet()); }

  // ---- Yardimcilar --------------------------------------------------------------
  const sleep = ms => new Promise(r => setTimeout(r, ms));

  function visible(el) {
    if (!el || !el.isConnected) return false;
    const st = getComputedStyle(el);
    if (st.display === "none" || st.visibility === "hidden") return false;
    return el.getClientRects().length > 0 || el.type === "radio" || el.type === "checkbox";
  }

  function cleanText(t) {
    return String(t || "")
      .replace(/\s+/g, " ")
      .replace(/\*/g, "")
      .replace(/\b(required|gerekli|zorunlu)\b/gi, "")
      .trim();
  }

  /** LinkedIn metni hem gorunur hem ekran okuyucu icin iki kez yazar; ilkini al. */
  function textOf(el) {
    if (!el) return "";
    const hidden = el.querySelector && el.querySelector('span[aria-hidden="true"]');
    return cleanText(hidden ? hidden.textContent : el.textContent);
  }

  function getDialog() {
    const modal = document.querySelector(".jobs-easy-apply-modal");
    if (modal && visible(modal)) return modal;
    const dialogs = [...document.querySelectorAll('[role="dialog"]')].filter(visible);
    return dialogs.find(d => d.querySelector("form") || /easy apply|kolay basvuru|kolay başvuru/i.test(d.textContent)) || null;
  }

  function labelFor(el) {
    if (el.type === "radio" || (el.type === "checkbox" && groupOf(el).length > 1)) {
      const fs = el.closest("fieldset");
      const lg = fs && fs.querySelector("legend");
      if (lg) return textOf(lg);
    }
    if (el.id) {
      const l = el.ownerDocument.querySelector('label[for="' + CSS.escape(el.id) + '"]');
      if (l && textOf(l)) return textOf(l);
    }
    const by = el.getAttribute("aria-labelledby");
    if (by) {
      const t = by.split(/\s+/).map(id => textOf(document.getElementById(id))).join(" ").trim();
      if (t) return t;
    }
    if (el.getAttribute("aria-label")) return cleanText(el.getAttribute("aria-label"));
    const wrapLabel = el.closest("label");
    if (wrapLabel && textOf(wrapLabel)) return textOf(wrapLabel);
    const grp = el.closest(".jobs-easy-apply-form-element, .fb-dash-form-element, .jobs-easy-apply-form-section__grouping");
    if (grp) {
      const l = grp.querySelector("label, legend, .fb-dash-form-element__label, span.t-14");
      if (l && textOf(l)) return textOf(l);
    }
    return cleanText(el.placeholder || el.name || "");
  }

  function groupOf(el) {
    if (!el.name) return [el];
    const scope = el.closest("fieldset") || el.form || document;
    return [...scope.querySelectorAll('input[type="' + el.type + '"][name="' + CSS.escape(el.name) + '"]')];
  }

  function optionLabel(input) {
    if (input.id) {
      const l = document.querySelector('label[for="' + CSS.escape(input.id) + '"]');
      if (l) return textOf(l);
    }
    const wrap = input.closest("label");
    return wrap ? textOf(wrap) : cleanText(input.value);
  }

  function isRequired(el) {
    if (el.required || el.getAttribute("aria-required") === "true") return true;
    const fs = el.closest("fieldset");
    if (fs && fs.querySelector('[aria-required="true"], .fb-dash-form-element__label--is-required')) return true;
    const grp = el.closest(".jobs-easy-apply-form-element, .fb-dash-form-element");
    return !!(grp && /\*/.test(grp.textContent));
  }

  // React'in degisikligi fark etmesi icin yerel setter + olaylar
  function setNativeValue(el, value) {
    const proto = el.tagName === "TEXTAREA" ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    const setter = Object.getOwnPropertyDescriptor(proto, "value").set;
    el.focus();
    setter.call(el, value);
    el.dispatchEvent(new Event("input", { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
  }

  function setSelect(sel, index) {
    const setter = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, "value").set;
    setter.call(sel, sel.options[index].value);
    sel.dispatchEvent(new Event("input", { bubbles: true }));
    sel.dispatchEvent(new Event("change", { bubbles: true }));
  }

  function isNumericField(el) {
    return el.type === "number" || /numeric/i.test(el.id || "") || el.inputMode === "numeric" || el.inputMode === "decimal";
  }

  function isTypeahead(el) {
    return el.getAttribute("role") === "combobox" || el.getAttribute("aria-autocomplete") === "list";
  }

  async function pickTypeahead(el) {
    for (let i = 0; i < 15; i++) {
      await sleep(150);
      const ownsId = el.getAttribute("aria-controls") || el.getAttribute("aria-owns");
      const box = (ownsId && document.getElementById(ownsId)) || document;
      const opt = box.querySelector('[role="option"], .basic-typeahead__selectable, .search-typeahead-v2__hit');
      if (opt && visible(opt)) {
        opt.click();
        return true;
      }
    }
    // Yedek: klavye ile ilk oneriyi sec
    el.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true }));
    el.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
    return false;
  }

  // ---- Isaretleme ---------------------------------------------------------------
  function markUnknown(target, label, kind, options) {
    target.setAttribute(UNKNOWN_ATTR, label);
    if (data.settings.highlightUnknown) {
      target.style.outline = "3px solid #f59e0b";
      target.style.outlineOffset = "2px";
    }
    const key = M.norm(label);
    if (!key) return;
    if (!data.unanswered.some(u => M.norm(u.label) === key)) {
      data.unanswered.push({ label, kind, options: options || [], url: location.href, ts: Date.now() });
      if (data.unanswered.length > 200) data.unanswered.shift();
    }
  }

  function markFilled(target) {
    target.setAttribute(FILLED_ATTR, "1");
    target.removeAttribute(UNKNOWN_ATTR);
    target.style.outline = "2px solid #22c55e";
    target.style.outlineOffset = "2px";
  }

  // ---- Alan doldurma ------------------------------------------------------------
  async function fillDialog(dialog) {
    const stats = { filled: 0, skipped: 0, unknown: 0, unknownRequired: 0 };
    const doneGroups = new Set();
    const controls = [...dialog.querySelectorAll("input, select, textarea")].filter(el => {
      const t = (el.type || "").toLowerCase();
      if (["hidden", "file", "submit", "button", "search"].includes(t)) return false;
      if (el.disabled || el.readOnly) return false;
      return visible(el);
    });

    for (const el of controls) {
      const type = (el.type || "").toLowerCase();

      if (type === "radio" || type === "checkbox") {
        const group = groupOf(el);
        const key = type + ":" + (el.name || el.id);
        if (doneGroups.has(key)) continue;
        doneGroups.add(key);
        await fillChoiceGroup(el, group, type, stats);
        continue;
      }

      const label = labelFor(el);
      if (el.tagName === "SELECT") {
        await fillSelect(el, label, stats);
        continue;
      }

      // metin / sayi / textarea
      if (el.value && el.value.trim() && !data.settings.overwrite) { stats.skipped++; continue; }
      const ans = M.resolve(label, data);
      if (!ans) {
        markUnknown(el, label, el.tagName === "TEXTAREA" ? "textarea" : "text");
        stats.unknown++;
        if (isRequired(el)) stats.unknownRequired++;
        continue;
      }
      let value = ans.value;
      if (isNumericField(el)) {
        value = M.toNumeric(value);
        if (!value) { markUnknown(el, label, "number"); stats.unknown++; if (isRequired(el)) stats.unknownRequired++; continue; }
      }
      setNativeValue(el, value);
      if (isTypeahead(el)) await pickTypeahead(el);
      el.dispatchEvent(new Event("blur", { bubbles: true }));
      markFilled(el);
      stats.filled++;
      await sleep(60);
    }
    return stats;
  }

  async function fillSelect(sel, label, stats) {
    const cur = sel.options[sel.selectedIndex];
    const curText = cur ? M.norm(cur.text) : "";
    const isPlaceholder = !cur || !cur.value || /^(select|choose|secin|seciniz|bir secenek)/.test(curText);
    if (!isPlaceholder && !data.settings.overwrite) { stats.skipped++; return; }

    const options = [...sel.options].map(o => ({ text: o.text, value: o.value }));
    const ans = M.resolve(label, data);
    const idx = ans ? M.pickOption(options, ans.value) : -1;
    if (idx < 0) {
      markUnknown(sel, label, "select", options.map(o => o.text));
      stats.unknown++;
      if (isRequired(sel)) stats.unknownRequired++;
      return;
    }
    setSelect(sel, idx);
    markFilled(sel);
    stats.filled++;
  }

  async function fillChoiceGroup(first, group, type, stats) {
    const fs = first.closest("fieldset") || first.parentElement;
    const label = labelFor(first);

    // Tek onay kutusu (ornegin "Sartlari kabul ediyorum")
    if (type === "checkbox" && group.length === 1) {
      if (first.checked) { stats.skipped++; return; }
      const custom = M.customMatch(label, data.customAnswers);
      const agree = /(agree|accept|consent|acknowledge|kabul|onayl|beyan|okudum)/i.test(M.norm(label));
      if ((custom && M.isYes(custom.answer)) || (!custom && agree)) {
        first.click();
        markFilled(fs);
        stats.filled++;
      } else if (isRequired(first)) {
        markUnknown(fs, label, "checkbox");
        stats.unknown++;
        stats.unknownRequired++;
      }
      return;
    }

    if (group.some(g => g.checked) && !data.settings.overwrite) { stats.skipped++; return; }
    const options = group.map(g => ({ text: optionLabel(g), value: g.value }));
    const ans = M.resolve(label, data);
    const idx = ans ? M.pickOption(options, ans.value) : -1;
    if (idx < 0) {
      markUnknown(fs, label, type, options.map(o => o.text));
      stats.unknown++;
      if (isRequired(first)) stats.unknownRequired++;
      return;
    }
    const input = group[idx];
    const lbl = input.id && document.querySelector('label[for="' + CSS.escape(input.id) + '"]');
    (lbl || input).click();
    if (!input.checked) input.click();
    markFilled(fs);
    stats.filled++;
  }

  // ---- Elle girilen cevaplari ogrenme ----------------------------------------
  function currentValueOf(target) {
    if (target.matches("select")) {
      const o = target.options[target.selectedIndex];
      return o && o.value ? o.text.trim() : "";
    }
    if (target.matches("input, textarea")) return target.value.trim();
    const checked = target.querySelector("input:checked");
    if (!checked) return "";
    if (checked.type === "checkbox" && target.querySelectorAll('input[type="checkbox"]').length === 1) return "Yes";
    return optionLabel(checked);
  }

  async function learnFrom(dialog) {
    if (!data.settings.learnAnswers) return;
    const learned = [];
    dialog.querySelectorAll("[" + UNKNOWN_ATTR + "]").forEach(t => {
      const label = t.getAttribute(UNKNOWN_ATTR);
      const val = currentValueOf(t);
      if (!label || !val) return;
      const pattern = M.norm(label);
      const existing = data.customAnswers.find(c => M.norm(c.pattern) === pattern);
      if (existing) existing.answer = val;
      else data.customAnswers.push({ pattern, answer: val, learned: true });
      data.unanswered = data.unanswered.filter(u => M.norm(u.label) !== pattern);
      t.removeAttribute(UNKNOWN_ATTR);
      learned.push(label);
    });
    if (learned.length) {
      await storageSet({ customAnswers: data.customAnswers, unanswered: data.unanswered });
      toast("Öğrenildi: " + learned.length + " yeni cevap kaydedildi");
    }
  }

  // ---- Buton bulma ----------------------------------------------------------------
  function findButton(dialog, kind) {
    const btns = [...dialog.querySelectorAll("button")].filter(b => visible(b) && !b.disabled);
    const res = {
      submit: /(submit application|submit|basvuruyu gonder|gonder)/,
      review: /(review your application|review|incele|gozden gecir)/,
      next: /(continue to next step|next|ileri|sonraki|devam)/
    }[kind];
    return btns.find(b => res.test(M.norm((b.getAttribute("aria-label") || "") + " " + b.textContent))) || null;
  }

  function hasErrors(dialog) {
    return [...dialog.querySelectorAll(".artdeco-inline-feedback--error, [role=alert]")].some(e => visible(e) && e.textContent.trim());
  }

  function handleReviewStep(dialog) {
    if (!data.settings.unfollowCompany) return;
    const cb = dialog.querySelector('#follow-company-checkbox, input[type="checkbox"][id*="follow"]');
    if (cb && cb.checked) cb.click();
  }

  // ---- Ana akis ---------------------------------------------------------------------
  async function run(opts) {
    if (running) return { ok: false, message: "Zaten çalışıyor" };
    running = true;
    stopRequested = false;
    setPanelBusy(true);
    try {
      await load();
      const autoAdvance = opts && typeof opts.autoAdvance === "boolean" ? opts.autoAdvance : data.settings.autoAdvance;
      let total = { filled: 0, unknown: 0, steps: 0 };

      for (let step = 0; step < 20 && !stopRequested; step++) {
        const dialog = getDialog();
        if (!dialog) {
          return finish(total.steps ? "Pencere kapandı." : "Kolay Başvuru penceresi bulunamadı. Önce 'Kolay Başvuru' butonuna basın.");
        }
        const stats = await fillDialog(dialog);
        total.filled += stats.filled;
        total.unknown += stats.unknown;
        total.steps++;
        await storageSet({ unanswered: data.unanswered });

        const submitBtn = findButton(dialog, "submit");
        if (submitBtn) {
          handleReviewStep(dialog);
          submitBtn.style.outline = "3px solid #0a66c2";
          return finish("Son adım: bilgileri kontrol edip 'Başvuruyu gönder'e kendiniz basın. (" + total.filled + " alan dolduruldu)");
        }
        if (!autoAdvance) {
          return finish(stats.filled + " alan dolduruldu" + (stats.unknown ? ", " + stats.unknown + " soru size kaldı (turuncu)" : "") + ".");
        }
        if (stats.unknownRequired > 0) {
          return finish(stats.unknownRequired + " zorunlu soruyu cevaplayamadım (turuncu). Doldurup tekrar çalıştırın.");
        }
        await sleep(data.settings.delayMs);
        if (hasErrors(dialog)) return finish("Formda hata var, lütfen kontrol edin.");

        const btn = findButton(dialog, "review") || findButton(dialog, "next");
        if (!btn) return finish("İleri butonu bulunamadı.");
        await learnFrom(dialog);
        btn.click();
        await sleep(data.settings.delayMs + 300);
        const after = getDialog();
        if (after && hasErrors(after)) return finish("LinkedIn bir alanı kabul etmedi; kırmızı uyarıları kontrol edin.");
      }
      return finish(stopRequested ? "Durduruldu." : "Adım sınırına ulaşıldı.");
    } catch (e) {
      console.error("[KolayBasvuru]", e);
      return finish("Hata: " + e.message);
    } finally {
      running = false;
      setPanelBusy(false);
    }
  }

  function finish(message) {
    toast(message);
    return { ok: true, message };
  }

  // ---- Sayfa ici panel ----------------------------------------------------------------
  let panel, statusEl, toastTimer;

  function ensurePanel() {
    if (panel && panel.isConnected) return panel;
    panel = document.createElement("div");
    panel.id = "kb-panel";
    panel.innerHTML =
      '<button type="button" data-act="fill" title="Bu adımı doldur">⚡ Doldur</button>' +
      '<button type="button" data-act="auto" title="Gönder adımına kadar doldur ve ilerle">⏩ Otomatik</button>' +
      '<button type="button" data-act="stop" title="Durdur">■</button>' +
      '<div class="kb-status"></div>';
    statusEl = panel.querySelector(".kb-status");
    panel.addEventListener("click", e => {
      const act = e.target && e.target.getAttribute("data-act");
      if (act === "fill") run({ autoAdvance: false });
      if (act === "auto") run({ autoAdvance: true });
      if (act === "stop") stopRequested = true;
    });
    document.documentElement.appendChild(panel);
    return panel;
  }

  function setPanelBusy(b) { if (panel) panel.classList.toggle("kb-busy", b); }

  function toast(msg) {
    ensurePanel();
    statusEl.textContent = msg;
    statusEl.style.display = "block";
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { statusEl.style.display = "none"; }, 9000);
  }

  function syncPanel() {
    const show = !!getDialog();
    if (show) ensurePanel().style.display = "flex";
    else if (panel) panel.style.display = "none";
  }

  // Kullanici "Ileri"ye kendisi bastiginda da elle girdigi cevaplari ogren
  document.addEventListener("click", e => {
    const btn = e.target && e.target.closest && e.target.closest("button");
    if (!btn) return;
    const dialog = getDialog();
    if (!dialog || !dialog.contains(btn)) return;
    if (/(next|review|submit|ileri|sonraki|devam|incele|gonder)/.test(M.norm((btn.getAttribute("aria-label") || "") + " " + btn.textContent))) {
      learnFrom(dialog);
    }
  }, true);

  let syncQueued = false;
  new MutationObserver(() => {
    if (syncQueued) return;
    syncQueued = true;
    setTimeout(() => { syncQueued = false; syncPanel(); }, 400);
  }).observe(document.documentElement, { childList: true, subtree: true });

  chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    if (msg && msg.type === "kb-fill") {
      run({ autoAdvance: !!msg.autoAdvance }).then(sendResponse);
      return true;
    }
    if (msg && msg.type === "kb-status") {
      sendResponse({ dialog: !!getDialog(), running });
    }
    return false;
  });

  load().then(syncPanel);
  // Testler icin disariya ac
  window.__kb = { run, fillDialog, getDialog, labelFor, learnFrom };
})();
