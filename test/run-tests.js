/* Testlar: node test/run-tests.js
 * 1) cleaner.js — matn rejimi
 * 2) word-engine.js — soddalashtirilgan Word API imitatsiyasi ustida
 */
"use strict";
const assert = require("assert");
const path = require("path");
const C = require(path.join(__dirname, "../web/cleaner.js"));

let pass = 0;
function t(name, fn) {
  return Promise.resolve().then(fn).then(
    () => { pass++; console.log("  ✓ " + name); },
    (e) => { console.error("  ✗ " + name + "\n    " + e.message); process.exitCode = 1; });
}
const all = (o) => Object.assign(C.defaultOptions(), o);

// ---------------------------------------------------------------- Word mock
// Matn + har bir belgi uchun bo'yoq (highlight) + xatcho'plar.
function makeWord(docText, initialHl) {
  const doc = { text: docText, hl: Array.from({ length: docText.length }, (_, i) => (initialHl && initialHl[i]) || null), bm: new Map() };
  let queue = [];
  const live = new Set();
  const track = (r) => (live.add(r), r);
  function findAll(hay, needle, opt) {
    const res = []; let i = 0;
    while ((i = hay.indexOf(needle, i)) >= 0) {
      const before = hay[i - 1], after = hay[i + needle.length];
      const ok = !opt.matchWholeWord || (!/\p{L}/u.test(before || "") && !/\p{L}/u.test(after || ""));
      if (ok) res.push(i);
      i += needle.length;
    }
    return res;
  }
  const decode = (q) => q.replace(/\^u(\d+)/g, (_, d) => String.fromCharCode(+d))
    .replace(/\^t/g, "\t").replace(/\^s/g, " ").replace(/\^-/g, "­")
    .replace(/\^\+/g, "—").replace(/\^=/g, "–").replace(/\^\^/g, "^");
  const collection = (fill) => { const c = { items: [], load() { return c; } }; queue.push({ kind: "read", run: () => { c.items = fill(); } }); return c; };

  class Range {
    constructor(s, e) { this.s = s; this.e = e; this.isNullObject = false; const self = this;
      this.font = {
        get highlightColor() { const set = new Set(doc.hl.slice(self.s, self.e)); return set.size === 1 ? [...set][0] : (set.size ? "" : null); },
        set highlightColor(v) { queue.push({ kind: "hl", r: self, v }); }
      };
    }
    get text() { return doc.text.slice(this.s, this.e); }
    load() { return this; }
    search(q, opt) {
      return collection(() => { const n = decode(q); return findAll(this.text, n, opt || {}).map((i) => track(new Range(this.s + i, this.s + i + n.length))); });
    }
    get paragraphs() {
      return collection(() => { let pos = this.s; return this.text.split("\r").map((p) => { const r = track(new Range(pos, pos + p.length)); pos += p.length + 1; return r; }); });
    }
    getTextRanges(marks, trim) {
      return collection(() => { const out = []; const re = / *[^ ]+ */g; let m;   // so'z + undan keyingi bo'shliqlar
        while ((m = re.exec(this.text))) out.push(track(new Range(this.s + m.index, this.s + m.index + m[0].length)));
        if (!out.length && this.text.length) out.push(track(new Range(this.s, this.e)));
        return out; });
    }
    getRange() { return track(new Range(this.s, this.e)); }
    getBookmarks() { const res = { value: [] }; queue.push({ kind: "read", run: () => { res.value = [...doc.bm.keys()]; } }); return res; }
    insertBookmark(name) { queue.push({ kind: "bm", run: () => doc.bm.set(name, track(new Range(this.s, this.e))) }); }
    insertText(txt) { queue.push({ kind: "edit", r: this, txt }); return this; }
    delete() { queue.push({ kind: "edit", r: this, txt: "" }); }
    select() {}
  }
  const body = track(new Range(0, docText.length));
  const ctx = {
    document: {
      body,
      getSelection() { return track(new Range(0, 0)); },
      getBookmarkRangeOrNullObject(n) { const r = new Range(0, 0);
        queue.push({ kind: "read", run: () => { const b = doc.bm.get(n); if (b) { r.s = b.s; r.e = b.e; track(r); } else r.isNullObject = true; } }); return r; },
      deleteBookmark(n) { queue.push({ kind: "bm", run: () => doc.bm.delete(n) }); },
      set changeTrackingMode(v) {}
    },
    async sync() {
      const q = queue; queue = [];
      q.filter((x) => x.kind === "read").forEach((x) => x.run());
      q.filter((x) => x.kind === "hl").forEach((x) => { for (let i = x.r.s; i < x.r.e; i++) doc.hl[i] = x.v; });
      q.filter((x) => x.kind === "bm").forEach((x) => x.run());
      const edits = q.filter((x) => x.kind === "edit").sort((a, b) => b.r.s - a.r.s);
      for (const ed of edits) {
        const { s, e } = ed.r; const n = ed.txt.length; const delta = n - (e - s);
        const inherit = doc.hl[s] !== undefined && e > s ? doc.hl[s] : (doc.hl[s - 1] || null);
        doc.text = doc.text.slice(0, s) + ed.txt + doc.text.slice(e);
        doc.hl.splice(s, e - s, ...Array(n).fill(inherit));
        const map = (p, isEnd) => p <= s ? p : p >= e ? p + delta : (isEnd ? s + n : s);
        for (const r of live) { const ns = r === ed.r ? s : map(r.s, false); const ne = r === ed.r ? s + n : map(r.e, true); r.s = ns; r.e = Math.max(ns, ne); }
      }
    }
  };
  global.Word = { run: (fn) => fn(ctx), ChangeTrackingMode: { trackAll: "TrackAll" } };
  global.Office = { context: { requirements: { isSetSupported: () => true } } };
  // yordamchi: qaysi belgilarda qaysi rang
  doc.colored = (color) => doc.text.split("").filter((_, i) => doc.hl[i] === color).join("");
  return doc;
}

global.window = global; global.AICleaner = C;
require(path.join(__dirname, "../web/word-engine.js"));
const E = global.WordEngine;

(async () => {
  console.log("cleaner.js");
  await t("yashirin belgilar o'chadi", () => {
    const r = C.clean("Sa​lom﻿ dunyo­!", all());
    assert.strictEqual(r.text, "Salom dunyo!"); assert.strictEqual(r.counts.hidden, 3);
  });
  await t("NBSP, tire, qo'shtirnoq, ko'p nuqta", () => {
    const r = C.clean("A B — “test” va ‘x’ …", all({ uzApostrophe: false }));
    assert.strictEqual(r.text, "A B - \"test\" va 'x' ...");
  });
  await t("o'zbekcha o‘, g‘ va tutuq belgisi saqlanadi", () => {
    const r = C.clean("O‘zbekiston, g‘alaba, ma’no, ‘iqtibos’", all());
    assert.strictEqual(r.text, "O‘zbekiston, g‘alaba, ma’no, 'iqtibos'");
  });
  await t("«» qo'shtirnoqlarga tegilmaydi", () => {
    assert.strictEqual(C.clean("«Kitob»", all()).text, "«Kitob»");
  });
  await t("markdown sarlavha, yulduzcha, oxirgi bo'shliq", () => {
    const r = C.clean("## Kirish  \nBu **muhim** gap.\t\n#hashtag", all());
    assert.strictEqual(r.text, "Kirish\nBu muhim gap.\n#hashtag");
    assert.strictEqual(r.counts.headings, 1); assert.strictEqual(r.counts.asterisks, 4); assert.strictEqual(r.counts.trailing, 3);
  });
  await t("o'xshash harflar faqat aralash so'zlarda", () => {
    const r = C.clean("pаypal va Москва", all({ lookalike: true }));
    assert.strictEqual(r.text, "paypal va Москва"); assert.strictEqual(r.counts.lookalike, 1);
  });
  await t("NFKC normallashtirish", () => {
    assert.strictEqual(C.clean("Ｈｅｌｌｏ ﬁle 𝐇𝐢", all({ nfkc: true })).text, "Hello file Hi");
  });
  await t("o'chirilgan qoida ishlamaydi", () => {
    assert.strictEqual(C.clean("a — b", all({ dashes: false })).text, "a — b");
  });

  // Yashirin "teg" belgilari bilan yozilgan matn (ASCII smuggling)
  const tags = (str) => Array.from(str, (c) => String.fromCodePoint(0xE0000 + c.charCodeAt(0))).join("");
  const SCOTLAND = "\u{1F3F4}" + tags("gbsct") + "\u{E007F}";
  await t("emoji ichidagi ZWJ, VS16 va teri rangi saqlanadi", () => {
    const src = "Oila: 👨‍👩‍👧, 👩🏽‍💻, ❤️, 👁️‍🗨️, 1️⃣";
    const r = C.clean(src, all());
    assert.strictEqual(r.text, src); assert.strictEqual(r.counts.hidden, 0);
  });
  await t("bayroq teglari saqlanadi, matnga yashirilgan teglar o'chadi", () => {
    assert.strictEqual(C.clean("Bayroq " + SCOTLAND, all()).text, "Bayroq " + SCOTLAND);
    const r = C.clean("Salom" + tags("ignore all") + " dunyo", all());
    assert.strictEqual(r.text, "Salom dunyo"); assert.strictEqual(r.counts.hidden, 10);
  });
  await t("emojidan tashqaridagi ZWJ va variation selector o'chadi", () => {
    const r = C.clean("so‍z a️b c︁d e\u{E0101}f 😀‍x", all());
    assert.strictEqual(r.text, "soz ab cd ef 😀x"); assert.strictEqual(r.counts.hidden, 5);
  });
  await t("ieroglif variantlari va braille matni saqlanadi", () => {
    const src = "葛\u{E0100} ⠓⠀⠊";
    assert.strictEqual(C.clean(src, all()).text, src);
  });
  await t("bo'sh ko'rinadigan to'ldiruvchi belgilar o'chadi", () => {
    const r = C.clean("aㅤb⠀c͏dﾠe", all());
    assert.strictEqual(r.text, "abcde"); assert.strictEqual(r.counts.hidden, 4);
  });
  await t("«» faqat sozlama yoqilganda almashtiriladi", () => {
    assert.strictEqual(C.clean("«Kitob» ‹x›", all()).text, "«Kitob» ‹x›");
    const r = C.clean("«Kitob» ‹x›", all({ guillemets: true }));
    assert.strictEqual(r.text, "\"Kitob\" 'x'"); assert.strictEqual(r.counts.quotes, 4);
    assert.strictEqual(C.clean("«Kitob»", all({ guillemets: true, quotes: false })).text, "«Kitob»");
  });
  await t("izlar ro'yxati kontekstli yashirin belgilarni to'g'ri ko'rsatadi", () => {
    const tr = C.traces("👨‍👩 a‍b" + tags("x"), all());
    assert.deepStrictEqual(tr.map((x) => [x.i, x.len, x.rule]), [[7, 1, "hidden"], [9, 2, "hidden"]]);
  });

  console.log("word-engine.js (imitatsiya)");
  await t("hujjat tozalanadi, natija matn rejimi bilan bir xil", async () => {
    const src = "## Sarlavha  \rSa​lom dunyo — “test”…\rO‘zbek tili, ma’no, ‘x’ **qalin**\t";
    const doc = makeWord(src);
    const r = await E.cleanInWord(all({ lookalike: true }), "document");
    const expected = C.clean(src.replace(/\r/g, "\n"), all({ lookalike: true })).text.replace(/\n/g, "\r");
    assert.strictEqual(doc.text, expected);
    assert.strictEqual(r.missed, 0);
  });
  await t("Word: emoji saqlanadi, yashirin teg va ZWJ o'chadi", async () => {
    const src = "Oila 👨‍👩‍👧 va " + SCOTLAND + "\rMat‍n" + tags("hi") + " «tayyor»";
    const doc = makeWord(src);
    const r = await E.cleanInWord(all({ guillemets: true }), "document");
    const expected = C.clean(src.replace(/\r/g, "\n"), all({ guillemets: true })).text.replace(/\n/g, "\r");
    assert.strictEqual(doc.text, expected);
    assert.strictEqual(doc.text, "Oila 👨‍👩‍👧 va " + SCOTLAND + "\rMatn \"tayyor\"");
    assert.strictEqual(r.counts.hidden, 3); assert.strictEqual(r.missed, 0);
  });
  await t("o'xshash harfli so'z Word'da tuzatiladi", async () => {
    const doc = makeWord("Login: pаypal.com, Москва");
    const r = await E.cleanInWord(all({ lookalike: true }), "document");
    assert.strictEqual(doc.text, "Login: paypal.com, Москва"); assert.strictEqual(r.counts.lookalike, 1);
  });
  await t("tahlil hujjatni o'zgartirmaydi", async () => {
    const doc = makeWord("a​b — c");
    const r = await E.analyzeInWord(all(), "document");
    assert.strictEqual(doc.text, "a​b — c"); assert.strictEqual(C.total(r.counts), 2);
  });

  console.log("AI izlarini ranglash");
  const Y = "#FFFF00", G = "#00FF00";
  await t("ko'rsatish: izli so'zlar sariq, toza so'zlar rangsiz", async () => {
    const doc = makeWord("Toza gap. Sa\u200Blom dunyo — “test”\rOddiy qator");
    const r = await E.markTraces(all(), "document", "mark");
    assert.strictEqual(doc.colored(Y), "Sa\u200Blom — “test”");
    assert.strictEqual(r.total, 3);
    assert.ok(r.spots[0].rules.includes("hidden") && r.spots[0].marks[0].s === 2);
  });
  await t("tozalash: aynan o'sha joylar yashil, sariq qolmaydi", async () => {
    const doc = makeWord("Toza gap. Sa\u200Blom dunyo — “test”");
    await E.markTraces(all(), "document", "mark");
    await E.cleanAndMark(all(), "document");
    assert.strictEqual(doc.text, "Toza gap. Salom dunyo - \"test\"");
    assert.strictEqual(doc.colored(Y), "");
    assert.strictEqual(doc.colored(G), "Salom - \"test\"");
  });
  await t("yo'qolib ketadigan '## ' o'rniga qo'shni so'z yashil bo'ladi", async () => {
    const doc = makeWord("## Kirish qismi");
    await E.cleanAndMark(all(), "document");
    assert.strictEqual(doc.text, "Kirish qismi");
    assert.strictEqual(doc.colored(G), "Kirish ");
  });
  await t("ranglarni o'chirish foydalanuvchining o'z bo'yog'ini tiklaydi", async () => {
    const src = "Muhim — gap va oddiy — so'z";
    const own = src.split("").map((_, i) => (i < 6 ? "#00FFFF" : null));   // "Muhim " oldindan ko'k
    const doc = makeWord(src, own);
    await E.markTraces(all(), "document", "mark");
    assert.ok(doc.colored(Y).length > 0);
    const n = await E.clearMarks();
    assert.strictEqual(n, 2);
    assert.strictEqual(doc.colored(Y) + doc.colored(G), "");
    assert.strictEqual(doc.colored("#00FFFF"), "Muhim ");
    assert.strictEqual(await E.countMarks(), 0);
  });
  await t("qayta ko'rsatish eski belgilarni ikki marta qo'shmaydi", async () => {
    makeWord("a — b — c");
    await E.markTraces(all(), "document", "mark");
    await E.markTraces(all(), "document", "mark");
    assert.strictEqual(await E.countMarks(), 2);
  });

  console.log("i18n (tarjimalar)");
  const fs = require("fs");
  const read = (f) => fs.readFileSync(path.join(__dirname, "..", f), "utf8");
  require(path.join(__dirname, "../web/i18n.js"));
  const L = global.I18N;
  const params = (s) => (s.match(/\{\w+\}/g) || []).sort().join(",");
  await t("har bir tilda barcha kalitlar va bir xil o'rinbosarlar bor", () => {
    const keys = Object.keys(L.STR.en);
    L.LANGS.forEach((lang) => {
      assert.deepStrictEqual(Object.keys(L.STR[lang]).sort(), keys.slice().sort(), lang);
      keys.forEach((k) => assert.strictEqual(params(L.STR[lang][k]), params(L.STR.en[k]), lang + "." + k));
    });
  });
  await t("barcha qoidalar inglizcha va ruschaga tarjima qilingan", () => {
    C.RULES.concat(C.EXTRA_OPTIONS).forEach((r) => ["en", "ru"].forEach((lang) => {
      const tr = L.RULES[lang][r.id];
      assert.ok(tr && tr[0] && tr[1], lang + "." + r.id);
    }));
  });
  await t("til almashadi, o'rinbosarlar to'ldiriladi", () => {
    L.setLang("ru-RU", false);
    assert.strictEqual(L.lang(), "ru");
    assert.strictEqual(L.t("indClean", { n: 5 }), "Очищено мест: 5 (зелёный)");
    L.setLang("uz-Latn-UZ", false);
    assert.strictEqual(L.rule(C.RULES[0]).label, C.RULES[0].label);
    L.setLang("de-DE", false);
    assert.strictEqual(L.lang(), "en");
  });
  await t("HTML'dagi data-i18n kalitlari mavjud", () => {
    ["web/taskpane.html", "web/notice.html"].forEach((f) => {
      const used = [...read(f).matchAll(/data-i18n(?:-title|-placeholder)?="(\w+)"/g)].map((m) => m[1]);
      assert.ok(used.length, f);
      used.forEach((k) => assert.ok(k in L.STR.en, f + ": " + k));
    });
  });
  await t("JS fayllarda kalitlar mavjud, o'zbekcha satrlar qattiq yozilmagan", () => {
    ["web/taskpane.js", "web/commands.js", "web/word-engine.js"].forEach((f) => {
      const src = read(f);
      [...src.matchAll(/\bt(?:r)?\("(\w+)"/g)].forEach((m) => assert.ok(m[1] in L.STR.en, f + ": " + m[1]));
      const code = src.replace(/\/\*[\s\S]*?\*\/|\/\/.*$/gm, "");
      assert.ok(!/["'][^"'\n]*[‘’«»][^"'\n]*["']/.test(code), f + ": qattiq yozilgan UI satri");
    });
  });
  await t("manifestdagi har bir satr ru-RU va uz-Latn-UZ tarjimasiga ega", () => {
    const m = read("manifest.xml");
    const blocks = [...m.matchAll(/<bt:String id="([^"]+)"[^>]*>([\s\S]*?)<\/bt:String>/g)];
    assert.ok(blocks.length >= 11);
    blocks.forEach(([, id, body]) => {
      assert.ok(body.includes('Locale="ru-RU"') && body.includes('Locale="uz-Latn-UZ"'), id);
    });
    assert.ok(/<SupportUrl DefaultValue="[^"]+\/support\.html"/.test(m));
  });
  await t("huquqiy sahifalarda uchala til va aloqa manzili bor", () => {
    ["privacy", "terms", "support"].forEach((n) => {
      const h = read("web/" + n + ".html");
      L.LANGS.forEach((lang) => assert.ok(h.includes('<section lang="' + lang + '"'), n + ":" + lang));
      assert.ok(h.includes("{{SUPPORT_EMAIL}}"), n);
    });
    assert.ok(read("web/privacy.html").includes("AI Text Cleaner"));
    const sup = read("web/support.html");
    assert.strictEqual((sup.match(/href="manifest\.xml" download/g) || []).length, 3, "har tilda manifest havolasi");
  });

  console.log(`\n${pass} ta test o'tdi`);
})();
