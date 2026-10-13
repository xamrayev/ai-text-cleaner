(function () {
  "use strict";
  var C = window.AICleaner;
  var S = window.AISettings;
  var L = window.I18N, t = L.t;
  var opts = S.load();
  var inWord = false;
  var $ = function (id) { return document.getElementById(id); };

  // --- Qoidalar ro'yxati ------------------------------------------------------
  function ruleRow(r, extraClass) {
    var lab = document.createElement("label");
    lab.className = "rule" + (extraClass ? " " + extraClass : "");
    var cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = !!opts[r.id];
    cb.dataset.id = r.id;
    cb.addEventListener("change", function () {
      opts[r.id] = cb.checked;
      S.save(opts);
      syncDependent();
    });
    var tx = L.rule(r);
    var b = document.createElement("b"); b.textContent = tx.label;
    var sm = document.createElement("small"); sm.textContent = tx.desc;
    lab.append(cb, b, sm);
    return lab;
  }

  function renderRules() {
    var list = $("rules-list"), extra = $("extra-list");
    list.textContent = ""; extra.textContent = "";
    var QUOTE_SUBS = ["uzApostrophe", "guillemets"];
    C.RULES.forEach(function (r) {
      list.appendChild(ruleRow(r));
      if (r.id === "quotes") C.EXTRA_OPTIONS.forEach(function (o) { if (QUOTE_SUBS.indexOf(o.id) >= 0) list.appendChild(ruleRow(o, "sub")); });
    });
    C.EXTRA_OPTIONS.forEach(function (o) { if (QUOTE_SUBS.indexOf(o.id) < 0) extra.appendChild(ruleRow(o)); });
    syncDependent();
  }

  function syncDependent() {
    ["uzApostrophe", "guillemets"].forEach(function (id) {
      var sub = document.querySelector('input[data-id="' + id + '"]');
      if (sub) { sub.disabled = !opts.quotes; sub.closest("label").style.opacity = opts.quotes ? "" : ".5"; }
    });
    var tr = document.querySelector('input[data-id="trackChanges"]');
    if (tr && !inWord) tr.closest("label").hidden = true;
  }

  // --- Natijalar jadvali ------------------------------------------------------
  // Til almashganda qayta chizish uchun oxirgi natijalar (sarlavha va izoh — kalit + parametrlar)
  var lastResults = {};
  function renderResults(el, title, counts, note) {
    lastResults[el.id] = { title: title, counts: counts, note: note };
    el.hidden = false;
    el.textContent = "";
    var h = document.createElement("h2"); h.textContent = msg(title); el.appendChild(h);
    var tbl = document.createElement("table");
    C.RULES.forEach(function (r) {
      if (!opts[r.id]) return;
      var n = Math.round(counts[r.id] || 0);
      var tr = tbl.insertRow();
      if (!n) tr.className = "zero";
      tr.insertCell().textContent = L.rule(r).label;
      var c = tr.insertCell(); c.className = "n"; c.textContent = L.num(n);
    });
    var tot = tbl.insertRow(); tot.className = "total";
    tot.insertCell().textContent = t("total");
    var tc = tot.insertCell(); tc.className = "n"; tc.textContent = L.num(Math.round(C.total(counts)));
    el.appendChild(tbl);
    if (note) { var p = document.createElement("p"); p.className = "hint"; p.style.marginTop = "6px"; p.textContent = msg(note); el.appendChild(p); }
  }

  /** [kalit, parametrlar] → joriy tildagi matn. */
  function msg(m) { return t(m[0], m[1]); }

  function setStatus(text, kind) {
    var s = $("status");
    s.textContent = text || "";
    s.className = "status" + (kind ? " " + kind : "");
  }

  function scope() { return document.querySelector('input[name="scope"]:checked').value; }

  function busy(on) {
    ["btn-show", "btn-clean", "btn-clear"].forEach(function (id) { $(id).disabled = on; });
    if (!on) {
      $("btn-clear").disabled = indicatorState === "off";
      if (inWord && !WordEngine.canMark()) $("btn-show").disabled = true;
    }
  }

  // --- Indikator ----------------------------------------------------------------
  var indicatorState = "off", indicatorN = 0, indicatorKey = null;
  function setIndicator(state, n, key) {
    indicatorState = state; indicatorN = n; indicatorKey = key || null;
    var el = $("indicator");
    el.className = "indicator " + state;
    $("indicator-text").textContent = t(key ||
      (state === "mark" ? "indMark" : state === "clean" ? "indClean" : state === "unknown" ? "indUnknown" : "indOff"),
      { n: L.num(n || 0) });
    $("btn-clear").disabled = state === "off";
  }

  // --- Topilgan joylar ro'yxati ------------------------------------------------
  function ruleLabel(id) {
    var r = C.RULES.filter(function (x) { return x.id === id; })[0];
    return r ? L.rule(r).label : id;
  }
  var VISIBLE = { "\t": "→", " ": "␣", "\u00A0": "⍽", "\u202F": "⍽", "\u2007": "⍽" };
  var PAGE = 100, shown = 0, currentSpots = [], currentDone = false;

  function snippetNode(spot) {
    var box = document.createElement("div");
    box.className = "snip";
    var marks = spot.marks.slice().sort(function (a, b) { return a.s - b.s; });
    var pos = 0, txt = spot.text;
    marks.forEach(function (m) {
      if (m.s < pos) return;
      if (m.s > pos) box.appendChild(document.createTextNode(txt.slice(pos, m.s)));
      var raw = txt.slice(m.s, m.e);
      var mk = document.createElement("mark");
      var vis = VISIBLE[raw] || (m.rule === "hidden" ? "¤" : null);
      mk.textContent = vis || raw;
      if (vis) mk.className = "inv";
      mk.title = ruleLabel(m.rule) + " (U+" + raw.codePointAt(0).toString(16).toUpperCase().padStart(4, "0") + ")";
      box.appendChild(mk);
      pos = m.e;
    });
    var rest = txt.slice(pos).replace(/ +$/, "");
    if (rest) box.appendChild(document.createTextNode(rest));
    if (!box.textContent.trim()) box.textContent = t("emptySpot");
    return box;
  }

  function renderSpots(spots, done) {
    currentSpots = spots; currentDone = done; shown = 0;
    $("spots-list").textContent = "";
    $("spots").hidden = !spots.length;
    $("spots-title").textContent = t(done ? "spotsCleaned" : "spotsFound") + " (" + spots.length + ")";
    moreSpots();
  }

  function moreSpots() {
    var list = $("spots-list");
    currentSpots.slice(shown, shown + PAGE).forEach(function (sp) {
      var li = document.createElement("li");
      li.className = "spot" + (currentDone ? " done" : "");
      li.tabIndex = 0;
      li.appendChild(snippetNode(sp));
      var tags = document.createElement("div"); tags.className = "tags";
      sp.rules.forEach(function (r) { var tg = document.createElement("span"); tg.className = "tag"; tg.textContent = ruleLabel(r); tags.appendChild(tg); });
      li.appendChild(tags);
      var go = function () { if (sp.name) WordEngine.selectMark(sp.name).catch(function () {}); };
      li.addEventListener("click", go);
      li.addEventListener("keydown", function (e) { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); go(); } });
      list.appendChild(li);
    });
    shown = Math.min(currentSpots.length, shown + PAGE);
    $("btn-more").hidden = shown >= currentSpots.length;
    $("btn-more").textContent = t("more", { n: currentSpots.length - shown });
  }

  // --- Hujjat rejimi ----------------------------------------------------------
  async function onShow() {
    busy(true); setStatus(t("stSearching"));
    try {
      var a = await WordEngine.analyzeInWord(opts, scope());
      var m = await WordEngine.markTraces(opts, scope(), "mark");
      renderResults($("doc-results"), [a.usedSelection ? "resFoundSel" : "resFoundDoc"], a.counts);
      renderSpots(m.spots, false);
      setIndicator(m.total ? "mark" : "off", m.total);
      setStatus(t(m.total ? "stTracesFound" : "stNoTraces"), m.total ? "" : "ok");
    } catch (e) { setStatus(t("error", { msg: e.message || e }), "err"); console.error(e); }
    finally { busy(false); }
  }

  async function onClean() {
    busy(true); setStatus(t("stCleaning"));
    try {
      var r = await WordEngine.cleanAndMark(opts, scope(), function (m) { setStatus(m); });
      var n = Math.round(C.total(r.counts));
      var note = r.missed ? ["noteMissed", { n: r.missed }] : [opts.trackChanges ? "noteTrack" : "noteUndo"];
      renderResults($("doc-results"), ["resCleaned"], r.counts, note);
      renderSpots(r.spots, true);
      setIndicator(r.spots.length ? "clean" : "off", r.spots.length);
      setStatus(n ? t("stChanges", { n: L.num(n) }) : t("stNothing"), r.missed ? "warn" : "ok");
    } catch (e) { setStatus(t("error", { msg: e.message || e }), "err"); console.error(e); }
    finally { busy(false); }
  }

  async function onClear() {
    busy(true);
    try {
      var n = await WordEngine.clearMarks();
      setIndicator("off");
      $("spots").hidden = true;
      setStatus(t(n ? "stColorsRemoved" : "stNoColors"), "ok");
    } catch (e) { setStatus(t("error", { msg: e.message || e }), "err"); console.error(e); }
    finally { busy(false); }
  }

  async function refreshIndicator() {
    try {
      var n = await WordEngine.countMarks();
      if (n === null) setIndicator("unknown");
      else setIndicator(n ? "mark" : "off", n);
      if (n) setIndicator("mark", n, "indCount");
    } catch (e) { /* e'tiborsiz */ }
  }

  // --- Matn rejimi ------------------------------------------------------------
  function onTextClean() {
    var r = C.clean($("txt-in").value, opts);
    $("txt-out").value = r.text;
    $("btn-copy").disabled = !r.text;
    $("btn-text-insert").disabled = !inWord || !r.text;
    renderResults($("text-results"), ["resResult"], r.counts);
  }

  async function onCopy() {
    var out = $("txt-out");
    try { await navigator.clipboard.writeText(out.value); }
    catch (e) { out.select(); document.execCommand("copy"); }
    $("btn-copy").textContent = t("btnCopied");
    setTimeout(function () { $("btn-copy").textContent = t("btnCopy"); }, 1500);
  }

  async function onInsert() {
    try {
      await Word.run(async function (ctx) {
        ctx.document.getSelection().insertText($("txt-out").value, "Replace");
        await ctx.sync();
      });
      $("btn-text-insert").textContent = t("btnInserted");
      setTimeout(function () { $("btn-text-insert").textContent = t("btnInsert"); }, 1500);
    } catch (e) { setTextStatus(e); }
  }

  // --- Tablar -----------------------------------------------------------------
  function switchTab(name) {
    document.querySelectorAll(".tab").forEach(function (t) {
      var on = t.dataset.tab === name;
      t.classList.toggle("active", on);
      t.setAttribute("aria-selected", on);
    });
    $("tab-doc").hidden = name !== "doc";
    $("tab-text").hidden = name !== "text";
  }

  function setTextStatus(e) {
    var el = $("text-results");
    el.hidden = false;
    el.textContent = "";
    var p = document.createElement("p"); p.className = "status err"; p.textContent = t("error", { msg: e.message || e });
    el.appendChild(p);
  }

  // --- Til ----------------------------------------------------------------------
  function applyLang() {
    L.apply(document);
    $("lang").value = L.lang();
    var q = "?lang=" + L.lang();
    [["lnk-privacy", "privacy.html", "lnkPrivacy"], ["lnk-terms", "terms.html", "lnkTerms"], ["lnk-support", "support.html", "lnkSupport"]]
      .forEach(function (l) { $(l[0]).href = l[1] + q; $(l[0]).textContent = t(l[2]); });
    renderRules();
    setIndicator(indicatorState, indicatorN, indicatorKey);
    Object.keys(lastResults).forEach(function (id) {
      var r = lastResults[id];
      if (!$(id).hidden) renderResults($(id), r.title, r.counts, r.note);
    });
    if (!$("spots").hidden) renderSpots(currentSpots, currentDone);
    if (inWord && !WordEngine.canMark()) $("btn-show").title = t("btnShowUnsupported");
    setStatus("");
  }

  function init() {
    L.detect();
    applyLang();
    $("lang").addEventListener("change", function () { L.setLang($("lang").value); applyLang(); });
    document.querySelectorAll(".tab").forEach(function (t) {
      t.addEventListener("click", function () { switchTab(t.dataset.tab); });
    });
    $("btn-show").addEventListener("click", onShow);
    $("btn-clean").addEventListener("click", onClean);
    $("btn-clear").addEventListener("click", onClear);
    $("btn-more").addEventListener("click", moreSpots);
    $("btn-text-clean").addEventListener("click", onTextClean);
    $("btn-copy").addEventListener("click", onCopy);
    $("btn-text-insert").addEventListener("click", onInsert);
    $("btn-reset").addEventListener("click", function () {
      opts = C.defaultOptions(); S.save(opts); renderRules();
    });
    if (inWord) {
      if (!WordEngine.canMark()) {
        $("btn-show").disabled = true;
        $("btn-show").title = t("btnShowUnsupported");
      }
      refreshIndicator();
    }
    if (!inWord) {
      // Word'dan tashqarida (oddiy brauzerda) faqat matn rejimi ishlaydi
      document.querySelector('.tab[data-tab="doc"]').disabled = true;
      switchTab("text");
    }
  }

  if (window.Office && Office.onReady) {
    Office.onReady(function (info) {
      inWord = info.host === Office.HostType.Word;
      init();
    });
  } else {
    document.addEventListener("DOMContentLoaded", init);
  }
})();
