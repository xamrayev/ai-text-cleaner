/*
 * AI Matn Tozalagich — tozalash yadrosi (Word API'ga bog'liq emas).
 * Brauzerda window.AICleaner, Node'da module.exports sifatida ishlaydi.
 */
(function (root) {
  "use strict";

  var RULES = [
    { id: "hidden", label: "Yashirin belgilar", desc: "Zero-width space, BOM, yumshoq defis, yo‘nalish belgilari, Unicode teglar (emoji saqlanadi)", def: true },
    { id: "nbsp", label: "Uzilmas bo‘shliqlar", desc: "NBSP va tor bo‘shliqlarni oddiy bo‘shliqqa almashtirish", def: true },
    { id: "dashes", label: "Tirelar", desc: "— va – belgilarini oddiy defisga (-) aylantirish", def: true },
    { id: "quotes", label: "Jingalak qo‘shtirnoqlar", desc: "“ ” ‘ ’ → \" va '", def: true },
    { id: "ellipsis", label: "Ko‘p nuqta", desc: "… belgisini uchta nuqtaga (...) aylantirish", def: true },
    { id: "trailing", label: "Qator oxiridagi bo‘shliqlar", desc: "Har paragraf oxiridagi ortiqcha bo‘shliq va tablar", def: true },
    { id: "asterisks", label: "Yulduzchalar (*)", desc: "Markdown qoldig‘i bo‘lgan barcha * belgilarini olib tashlash", def: true },
    { id: "headings", label: "Markdown sarlavhalari", desc: "Paragraf boshidagi # … ###### belgilarini olib tashlash", def: true },
    { id: "lookalike", label: "O‘xshash harflar", desc: "Lotin so‘z ichiga aralashgan kirill/yunon harflarini tuzatish", def: false },
    { id: "nfkc", label: "Unicode normallashtirish (NFKC)", desc: "Ｈｅｌｌｏ, ﬁ, 𝐇𝐞𝐥𝐥𝐨 → oddiy shakl (², ™ ham o‘zgaradi)", def: false }
  ];

  var EXTRA_OPTIONS = [
    { id: "uzApostrophe", label: "O‘zbekcha o‘, g‘ va tutuq belgisini saqlash", desc: "o‘, g‘, ma’no kabi joylardagi ‘ ’ belgilariga tegmaydi", def: true },
    { id: "guillemets", label: "«» qo‘shtirnoqlarni ham almashtirish", desc: "« » → \" va ‹ › → '", def: false },
    { id: "trackChanges", label: "O‘zgarishlarni kuzatish rejimida", desc: "Har bir tuzatish Word’da qayta ko‘rib chiqish uchun belgilanadi", def: false }
  ];

  function defaultOptions() {
    var o = {};
    RULES.concat(EXTRA_OPTIONS).forEach(function (r) { o[r.id] = r.def; });
    return o;
  }

  function emptyCounts() {
    var c = {};
    RULES.forEach(function (r) { c[r.id] = 0; });
    return c;
  }

  // --- Belgilar to'plamlari -------------------------------------------------
  var HIDDEN = {};
  [0x200B, 0x200C, 0x2060, 0x2061, 0x2062, 0x2063, 0x2064, 0xFEFF, 0x00AD,
   0x200E, 0x200F, 0x202A, 0x202B, 0x202C, 0x202D, 0x202E,
   0x2066, 0x2067, 0x2068, 0x2069, 0x180E, 0x061C,
   // bo'sh ko'rinadigan "to'ldiruvchi" belgilar
   0x3164, 0xFFA0, 0x115F, 0x1160, 0x034F, 0x17B4, 0x17B5, 0xFFF9, 0xFFFA, 0xFFFB].forEach(function (cp) { HIDDEN[cp] = true; });

  var GUILLEMETS = { 0x00AB: "\"", 0x00BB: "\"", 0x2039: "'", 0x203A: "'" };
  var NBSP = { 0x00A0: true, 0x202F: true, 0x2007: true };
  var DASHES = { 0x2010: true, 0x2011: true, 0x2012: true, 0x2013: true, 0x2014: true, 0x2015: true };
  var DQUOTES = { 0x201C: true, 0x201D: true, 0x201E: true, 0x201F: true, 0x2033: true };
  var SQUOTES = { 0x2018: true, 0x2019: true, 0x201A: true, 0x201B: true, 0x2032: true };
  // O'zbekcha kontekstda saqlanishi mumkin bo'lgan apostroflar
  var UZ_SENSITIVE = { 0x2018: true, 0x2019: true };

  var LOOKALIKE = {
    // Kirill
    "а": "a", "е": "e", "о": "o", "р": "p", "с": "c", "у": "y", "х": "x", "і": "i", "ј": "j", "ѕ": "s", "ԁ": "d", "һ": "h", "ӏ": "l", "ԛ": "q", "ԝ": "w",
    "А": "A", "В": "B", "Е": "E", "К": "K", "М": "M", "Н": "H", "О": "O", "Р": "P", "С": "C", "Т": "T", "Х": "X", "І": "I", "Ј": "J", "Ѕ": "S", "Ү": "Y",
    // Yunon
    "Α": "A", "Β": "B", "Ε": "E", "Ζ": "Z", "Η": "H", "Ι": "I", "Κ": "K", "Μ": "M", "Ν": "N", "Ο": "O", "Ρ": "P", "Τ": "T", "Υ": "Y", "Χ": "X",
    "ο": "o", "ν": "v", "ι": "i", "κ": "k", "ρ": "p"
  };

  var LATIN_RE = /[A-Za-z]/;
  var WORD_RE = /[\p{L}\p{M}]+/gu;
  var HEADING_RE = /^([ \t]{0,3}#{1,6}[ \t]+)(?=\S)/;
  var TRAILING_RE = /[ \t   　]+$/;

  function isLetter(ch) { return !!ch && /[\p{L}]/u.test(ch); }

  /** O'zbekcha o‘ / g‘ / tutuq belgisi kontekstidami? */
  function isUzProtected(prev, next) {
    if (prev && /[oOgG]/.test(prev)) return true;          // o‘, g‘
    if (isLetter(prev) && isLetter(next)) return true;      // ma’no, san’at
    return false;
  }

  /**
   * Bitta belgi uchun shartsiz almashtirish qoidasi.
   * Kontekstga bog'liq bo'lsa (uzApostrophe) null qaytaradi — u alohida ishlanadi.
   */
  function charRule(ch, opts) {
    var cp = ch.codePointAt(0);
    if (opts.hidden && HIDDEN[cp]) return { rule: "hidden", replace: "" };
    if (opts.nbsp && NBSP[cp]) return { rule: "nbsp", replace: " " };
    if (opts.dashes && DASHES[cp]) return { rule: "dashes", replace: "-" };
    if (opts.quotes && opts.guillemets && GUILLEMETS[cp]) return { rule: "quotes", replace: GUILLEMETS[cp] };
    if (opts.quotes && DQUOTES[cp]) return { rule: "quotes", replace: "\"" };
    if (opts.quotes && SQUOTES[cp]) {
      if (opts.uzApostrophe && UZ_SENSITIVE[cp]) return null; // kontekstli
      return { rule: "quotes", replace: "'" };
    }
    if (opts.ellipsis && cp === 0x2026) return { rule: "ellipsis", replace: "..." };
    if (opts.asterisks && ch === "*") return { rule: "asterisks", replace: "" };
    if (opts.nfkc && cp > 0x7F) {
      var n = ch.normalize("NFKC");
      if (n !== ch) return { rule: "nfkc", replace: n };
    }
    return null;
  }

  // --- Kontekstga bog'liq yashirin belgilar ----------------------------------
  // ZWJ, variation selector, teglar va bo'sh braille emoji, bayroq, ieroglif va braille
  // matnida qonuniy; boshqa joyda esa ular matnga yashirin ma'lumot qo'shish uchun ishlatiladi.
  var PICT_RE = /\p{Extended_Pictographic}/u;
  var HAN_RE = /\p{Script=Han}/u;
  var MATH_RE = /\p{Sm}/u;
  var KEYCAP_BASE_RE = /^[#*0-9]$/;

  function cpOf(c) { return c ? c.codePointAt(0) : -1; }
  function isVS(cp) { return cp >= 0xFE00 && cp <= 0xFE0F; }
  function isVSSup(cp) { return cp >= 0xE0100 && cp <= 0xE01EF; }
  function isTag(cp) { return cp >= 0xE0000 && cp <= 0xE007F; }
  function isContextHidden(cp) { return cp === 0x200D || isVS(cp) || isVSSup(cp) || isTag(cp) || cp === 0x2800; }

  /** cps[k] dan oldingi emoji asosi (VS16 va teri rangi modifikatorlari o'tkazib yuboriladi). */
  function emojiBase(cps, k) {
    for (var j = k - 1; j >= 0; j--) {
      var cp = cpOf(cps[j]);
      if (cp === 0xFE0F || (cp >= 0x1F3FB && cp <= 0x1F3FF)) continue;
      return cps[j];
    }
    return null;
  }

  /** k atrofidagi bir xil turdagi belgilar ketma-ketligi [a, b]. */
  function run(cps, k, test) {
    var a = k, b = k;
    while (a > 0 && test(cpOf(cps[a - 1]))) a--;
    while (b < cps.length - 1 && test(cpOf(cps[b + 1]))) b++;
    return [a, b];
  }

  /** 🏴 + teglar + U+E007F — Shotlandiya kabi bayroqlar. */
  function validTagRun(cps, k) {
    var r = run(cps, k, isTag), a = r[0], b = r[1];
    if (cpOf(cps[a - 1]) !== 0x1F3F4 || cpOf(cps[b]) !== 0xE007F || b === a) return false;
    for (var j = a; j < b; j++) { var cp = cpOf(cps[j]); if (cp < 0xE0020 || cp > 0xE007E) return false; }
    return true;
  }

  function isBraille(cp) { return cp >= 0x2801 && cp <= 0x28FF; }

  /** cps — code point'lar massivi. true: cps[k] o'chirilishi kerak. */
  function hiddenInContext(cps, k) {
    var cp = cpOf(cps[k]), prev = cps[k - 1], next = cps[k + 1];
    if (cp === 0x200D) {
      var base = emojiBase(cps, k);
      return !(base && PICT_RE.test(base) && next && PICT_RE.test(next));
    }
    if (cp === 0xFE0E || cp === 0xFE0F) {
      return !(prev && (PICT_RE.test(prev) || (KEYCAP_BASE_RE.test(prev) && cpOf(next) === 0x20E3)));
    }
    if (isVS(cp)) return !(prev && (HAN_RE.test(prev) || MATH_RE.test(prev)));
    if (isVSSup(cp)) return !(prev && HAN_RE.test(prev));
    if (isTag(cp)) return !validTagRun(cps, k);
    if (cp === 0x2800) {
      var r = run(cps, k, function (c) { return c === 0x2800; });
      return !(isBraille(cpOf(cps[r[0] - 1])) || isBraille(cpOf(cps[r[1] + 1])));
    }
    return false;
  }

  /** Kontekstli yashirin belgi: true — o'chiriladi, false — saqlanadi, null — bu qoida emas. */
  function contextHidden(cps, k, opts) {
    if (!opts.hidden || !isContextHidden(cpOf(cps[k]))) return null;
    return hiddenInContext(cps, k);
  }

  function contextQuoteReplace(text, i, ch, opts) {
    if (!(opts.quotes && opts.uzApostrophe && UZ_SENSITIVE[ch.codePointAt(0)])) return null;
    var prev = text[i - 1], next = text[i + 1];
    return isUzProtected(prev, next) ? null : "'";
  }

  function fixWord(word) {
    var out = "";
    for (var k = 0; k < word.length; k++) out += LOOKALIKE[word[k]] || word[k];
    return out;
  }

  /** Lotin harfi va unga o'xshash kirill/yunon harfi aralashgan so'zlar. */
  function lookalikeWords(text) {
    var map = {}, list = [];
    var m;
    WORD_RE.lastIndex = 0;
    while ((m = WORD_RE.exec(text))) {
      var w = m[0];
      if (!LATIN_RE.test(w)) continue;
      var fixed = fixWord(w);
      if (fixed === w) continue;
      if (!map[w]) { map[w] = { find: w, replace: fixed, count: 0, chars: 0 }; list.push(map[w]); }
      map[w].count++;
      for (var k = 0; k < w.length; k++) if (LOOKALIKE[w[k]]) map[w].chars++;
    }
    return list;
  }

  // --- Matn rejimi: to'liq tozalash -----------------------------------------
  function clean(text, opts) {
    opts = opts || defaultOptions();
    var counts = emptyCounts();
    var out = "";

    // 1) belgi darajasi (code point bo'yicha)
    var chars = Array.from(text);
    for (var i = 0; i < chars.length; i++) {
      var ch = chars[i];
      var r = charRule(ch, opts);
      if (r) { counts[r.rule]++; out += r.replace; continue; }
      var hc = contextHidden(chars, i, opts);
      if (hc !== null) { if (hc) counts.hidden++; else out += ch; continue; }
      var q = contextQuoteReplace(chars, i, ch, opts);
      if (q !== null) { counts.quotes++; out += q; continue; }
      out += ch;
    }

    // 2) aralash so'zlar
    if (opts.lookalike) {
      lookalikeWords(out).forEach(function (w) { counts.lookalike += w.chars; });
      out = out.replace(WORD_RE, function (w) { return LATIN_RE.test(w) ? fixWord(w) : w; });
    }

    // 3) qator darajasi
    if (opts.headings || opts.trailing) {
      out = out.split("\n").map(function (line) {
        var cr = line.endsWith("\r") ? "\r" : "";
        if (cr) line = line.slice(0, -1);
        if (opts.headings) {
          var h = HEADING_RE.exec(line);
          if (h) { counts.headings++; line = line.slice(h[1].length); }
        }
        if (opts.trailing) {
          var t = TRAILING_RE.exec(line);
          if (t) { counts.trailing += t[0].length; line = line.slice(0, t.index); }
        }
        return line + cr;
      }).join("\n");
    }
    return { text: out, counts: counts };
  }

  function analyze(text, opts) { return clean(text, opts).counts; }

  function total(counts) {
    return Object.keys(counts).reduce(function (s, k) { return s + counts[k]; }, 0);
  }

  // --- Word uchun reja ------------------------------------------------------
  /** Hujjatdagi shartsiz almashtiriladigan noyob belgilar ro'yxati. */
  function charPlan(text, opts) {
    var seen = {}, list = [];
    Array.from(text).forEach(function (ch) {
      var r = charRule(ch, opts);
      if (!r) return;
      if (!seen[ch]) { seen[ch] = { ch: ch, rule: r.rule, replace: r.replace, count: 0 }; list.push(seen[ch]); }
      seen[ch].count++;
    });
    return list;
  }

  function occurrenceIndices(text, ch, positions) {
    var idx = [], n = 0;
    for (var i = 0; i < text.length; i++) {
      if (text[i] === ch) { if (positions(i)) idx.push(n); n++; }
    }
    return idx;
  }

  /**
   * Bitta paragraf uchun kontekstli amallar:
   *  {kind:'prefix', find, rule}            — boshidagi markdown sarlavha
   *  {kind:'occ', find, indices, replace, rule} — belgining n-chi uchrashuvlari
   */
  function paragraphPlan(text, opts) {
    var ops = [];
    if (opts.headings) {
      var h = HEADING_RE.exec(text);
      if (h) ops.push({ kind: "prefix", find: h[1], rule: "headings" });
    }
    if (opts.trailing) {
      var t = TRAILING_RE.exec(text);
      if (t) {
        var start = t.index;
        var uniq = {};
        Array.from(t[0]).forEach(function (c) { uniq[c] = true; });
        Object.keys(uniq).forEach(function (c) {
          ops.push({ kind: "occ", find: c, replace: "", rule: "trailing",
            indices: occurrenceIndices(text, c, function (i) { return i >= start; }) });
        });
      }
    }
    if (opts.hidden) {
      var cps = Array.from(text), byChar = {}, order = [];
      cps.forEach(function (c, k) {
        if (!isContextHidden(cpOf(c))) return;
        if (!byChar[c]) { byChar[c] = { n: 0, idx: [] }; order.push(c); }
        if (hiddenInContext(cps, k)) byChar[c].idx.push(byChar[c].n);
        byChar[c].n++;
      });
      order.forEach(function (c) {
        if (byChar[c].idx.length) ops.push({ kind: "occ", find: c, replace: "", rule: "hidden", indices: byChar[c].idx });
      });
    }
    if (opts.quotes && opts.uzApostrophe) {
      ["‘", "’"].forEach(function (c) {
        if (text.indexOf(c) < 0) return;
        var idx = occurrenceIndices(text, c, function (i) { return !isUzProtected(text[i - 1], text[i + 1]); });
        if (idx.length) ops.push({ kind: "occ", find: c, replace: "'", rule: "quotes", indices: idx });
      });
    }
    return ops;
  }

  // --- AI izlarini ko'rsatish uchun: aniq pozitsiyalar ----------------------
  /**
   * Paragraf matnidagi har bir iz: {i, len, rule, del}
   * del=true — tozalashda belgi butunlay o'chadi, false — almashtiriladi.
   */
  function traces(text, opts) {
    var out = [];
    var headEnd = 0, tailStart = text.length;
    if (opts.headings) {
      var h = HEADING_RE.exec(text);
      if (h) { headEnd = h[1].length; for (var a = 0; a < headEnd; a++) out.push({ i: a, len: 1, rule: "headings", del: true }); }
    }
    if (opts.trailing) {
      var t = TRAILING_RE.exec(text);
      if (t && t.index >= headEnd) { tailStart = t.index; for (var b = t.index; b < text.length; b++) out.push({ i: b, len: 1, rule: "trailing", del: true }); }
    }
    var cps = Array.from(text);
    for (var i = 0, k = 0; k < cps.length; k++) {
      var ch = cps[k];
      var L = ch.length;
      if (i >= headEnd && i < tailStart) {
        var r = charRule(ch, opts);
        if (r) out.push({ i: i, len: L, rule: r.rule, del: r.replace === "" });
        else if (contextHidden(cps, k, opts)) out.push({ i: i, len: L, rule: "hidden", del: true });
        else if (contextQuoteReplace(text, i, ch, opts) !== null) out.push({ i: i, len: L, rule: "quotes", del: false });
      }
      i += L;
    }
    if (opts.lookalike) {
      var m;
      WORD_RE.lastIndex = 0;
      while ((m = WORD_RE.exec(text))) {
        if (!LATIN_RE.test(m[0])) continue;
        for (var k = 0; k < m[0].length; k++) {
          if (LOOKALIKE[m[0][k]]) out.push({ i: m.index + k, len: 1, rule: "lookalike", del: false });
        }
      }
    }
    return out.sort(function (x, y) { return x.i - y.i; });
  }

  /** Word bergan "so'z bo'laklari" matnlarini paragraf ichidagi pozitsiyalarga bog'lash. */
  function mapTokens(text, tokTexts) {
    var cur = 0;
    return tokTexts.map(function (tt) {
      if (!tt) return null;
      var at = text.indexOf(tt, cur);
      if (at < 0) return null;
      cur = at + tt.length;
      return { start: at, end: at + tt.length };
    });
  }

  /**
   * Izlarni bo'laklar bo'yicha guruhlash.
   * withNeighbors=true bo'lsa, tozalashda butunlay yo'qoladigan bo'lak (masalan, "## " yoki
   * oxirgi bo'shliqlar) uchun qo'shni bo'lak ham belgilanadi — yashil rang ko'rinib qolishi uchun.
   * Natija: [{tok, marks:[{s,e,rule}], rules:[...], listed:bool}]
   */
  function groupTraces(text, trs, spans, withNeighbors) {
    var groups = {}, order = [];
    function nearest(i) {
      var best = -1;
      for (var k = 0; k < spans.length; k++) {
        if (!spans[k]) continue;
        if (spans[k].start <= i) best = k; else { if (best < 0) best = k; break; }
      }
      return best;
    }
    function covering(i) {
      for (var k = 0; k < spans.length; k++) if (spans[k] && spans[k].start <= i && i < spans[k].end) return k;
      return -1;
    }
    trs.forEach(function (tr) {
      var k = covering(tr.i), inside = k >= 0;
      if (!inside) k = nearest(tr.i);
      if (k < 0) return;
      var g = groups[k];
      if (!g) { g = groups[k] = { tok: k, marks: [], rules: [], delChars: 0, listed: true }; order.push(k); }
      if (g.rules.indexOf(tr.rule) < 0) g.rules.push(tr.rule);
      if (inside) {
        g.marks.push({ s: tr.i - spans[k].start, e: tr.i - spans[k].start + tr.len, rule: tr.rule });
        if (tr.del) g.delChars += tr.len;
      }
    });
    if (withNeighbors) {
      var vanishes = function (k) {
        var g = groups[k], sp = spans[k];
        return !!(g && sp && g.delChars >= sp.end - sp.start);
      };
      order.slice().forEach(function (k) {
        if (!vanishes(k)) return;                                   // bo'lak saqlanib qoladi
        var dir = groups[k].rules.indexOf("headings") >= 0 ? 1 : -1; // sarlavha → keyingi, aks holda oldingi
        for (var n = k + dir; n >= 0 && n < spans.length; n += dir) {
          if (!spans[n] || vanishes(n)) continue;
          if (!groups[n]) { groups[n] = { tok: n, marks: [], rules: [], delChars: 0, listed: false }; order.push(n); }
          break;
        }
      });
    }
    return order.sort(function (x, y) { return x - y; }).map(function (k) { return groups[k]; });
  }

  // Word qidiruvi uchun maxsus kodlar
  var WORD_CODES = { "\t": "^t", " ": "^s", "­": "^-", "—": "^+", "–": "^=" };

  function escapeWordSearch(s) { return s.replace(/\^/g, "^^"); }

  /** Belgini Word'da topish uchun nomzod qidiruv satrlari (birinchisi asosiy). */
  function wordSearchCandidates(s) {
    var c = [];
    if (s.length <= 2 && WORD_CODES[s]) c.push(WORD_CODES[s]);
    c.push(escapeWordSearch(s));
    var cp = s.codePointAt(0);
    if (Array.from(s).length === 1 && cp <= 0xFFFF && cp > 0x7F) c.push("^u" + cp);
    return c.filter(function (v, i, a) { return a.indexOf(v) === i; });
  }

  /** Ko'p belgili satr (prefix/so'z) uchun Word qidiruv satri. */
  function wordSearchString(s) {
    return escapeWordSearch(s).replace(/\t/g, "^t");
  }

  var api = {
    RULES: RULES, EXTRA_OPTIONS: EXTRA_OPTIONS,
    defaultOptions: defaultOptions, emptyCounts: emptyCounts,
    clean: clean, analyze: analyze, total: total,
    charPlan: charPlan, paragraphPlan: paragraphPlan, lookalikeWords: lookalikeWords,
    traces: traces, mapTokens: mapTokens, groupTraces: groupTraces,
    wordSearchCandidates: wordSearchCandidates, wordSearchString: wordSearchString
  };

  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.AICleaner = api;
})(typeof window !== "undefined" ? window : this);
