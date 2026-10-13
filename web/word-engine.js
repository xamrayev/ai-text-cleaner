/*
 * Word hujjatida tozalash. Matnni butunlay qayta yozmaydi —
 * faqat kerakli belgilarni qidirib almashtiradi, shuning uchun
 * shrift, rang, havola, jadval va rasmlar saqlanib qoladi.
 */
(function (root) {
  "use strict";
  var C = root.AICleaner;
  var SEARCH = { matchCase: true, matchWildcards: false };
  function tr(key, params) { return root.I18N ? root.I18N.t(key, params) : key; }

  function apply(item, replace) {
    if (replace === "") item.delete();
    else item.insertText(replace, "Replace");
  }

  /** Doira: belgilangan qism (bo'sh bo'lsa — butun hujjat) yoki butun hujjat. */
  async function resolveScope(ctx, scope) {
    if (scope === "selection") {
      var sel = ctx.document.getSelection();
      sel.load("text");
      await ctx.sync();
      if (sel.text && sel.text.length > 0) return { range: sel, usedSelection: true };
    }
    return { range: ctx.document.body, usedSelection: false };
  }

  /** Faqat hisoblash (hujjat o'zgarmaydi). */
  async function analyzeInWord(opts, scope) {
    return Word.run(async function (ctx) {
      var s = await resolveScope(ctx, scope);
      s.range.load("text");
      await ctx.sync();
      // Word paragraflarni \r bilan ajratadi
      var text = s.range.text.replace(/\r/g, "\n");
      return { counts: C.analyze(text, opts), usedSelection: s.usedSelection };
    });
  }

  async function cleanInWord(opts, scope, onProgress) {
    var progress = onProgress || function () {};
    return Word.run(async function (ctx) {
      var counts = C.emptyCounts();
      var missed = 0;
      var s = await resolveScope(ctx, scope);
      var range = s.range;

      if (opts.trackChanges && Office.context.requirements.isSetSupported("WordApi", "1.4")) {
        ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
      }

      // 1) Paragraf darajasi: kontekstli yashirin belgilar, sarlavha prefikslari, oxirgi bo'shliqlar, o'zbekcha apostroflar
      if (opts.hidden || opts.headings || opts.trailing || (opts.quotes && opts.uzApostrophe)) {
        progress(tr("prParagraphs"));
        var paras = range.paragraphs;
        paras.load("items/text");
        await ctx.sync();

        var jobs = [];
        paras.items.forEach(function (p) {
          C.paragraphPlan(p.text, opts).forEach(function (op) {
            var q = op.kind === "prefix" ? C.wordSearchString(op.find) : C.wordSearchCandidates(op.find)[0];
            op.res = p.search(q, SEARCH);
            op.res.load("items");
            jobs.push(op);
          });
        });
        if (jobs.length) {
          await ctx.sync();
          jobs.forEach(function (op) {
            var items = op.res.items;
            if (op.kind === "prefix") {
              if (items[0]) { items[0].delete(); counts[op.rule]++; } else missed++;
            } else {
              op.indices.forEach(function (i) {
                if (items[i]) { apply(items[i], op.replace); counts[op.rule]++; } else missed++;
              });
            }
          });
          await ctx.sync();
        }
      }

      // 2) Belgi darajasi: har bir noyob belgi bitta qidiruv bilan
      range.load("text");
      await ctx.sync();
      var plan = C.charPlan(range.text, opts);
      var attempt = plan.map(function (op) { return { op: op, cands: C.wordSearchCandidates(op.ch), tried: 0 }; });

      while (attempt.length) {
        progress(tr("prReplacing", { n: attempt.length }));
        attempt.forEach(function (a) {
          a.res = range.search(a.cands[a.tried], SEARCH);
          a.res.load("items");
        });
        await ctx.sync();
        var next = [];
        attempt.forEach(function (a) {
          var items = a.res.items;
          if (items.length === 0) {
            a.tried++;
            if (a.tried < a.cands.length) next.push(a);   // boshqa qidiruv kodi bilan qayta urinish
            else missed += a.op.count;
            return;
          }
          items.forEach(function (it) { apply(it, a.op.replace); });
          counts[a.op.rule] += items.length;
        });
        await ctx.sync();
        attempt = next;
      }

      // 3) Aralash yozuvli so'zlar (lotin + kirill/yunon)
      if (opts.lookalike) {
        progress(tr("prLookalike"));
        range.load("text");
        await ctx.sync();
        var words = C.lookalikeWords(range.text).filter(function (w) { return w.find.length < 250; });
        words.forEach(function (w) {
          w.res = range.search(C.wordSearchString(w.find), { matchCase: true, matchWholeWord: true });
          w.res.load("items");
        });
        if (words.length) {
          await ctx.sync();
          words.forEach(function (w) {
            var per = w.chars / w.count;
            w.res.items.forEach(function (it) { it.insertText(w.replace, "Replace"); counts.lookalike += per; });
            if (w.res.items.length < w.count) missed += w.count - w.res.items.length;
          });
          await ctx.sync();
        }
      }

      return { counts: counts, missed: missed, usedSelection: s.usedSelection };
    });
  }

  // =========================================================================
  //  AI izlarini rang bilan ko'rsatish
  //  Har bir belgilangan bo'lak yashirin xatcho'p (bookmark) bilan eslab qolinadi:
  //  shu tufayli «Ranglarni o'chirish» faqat bizning ranglarimizni olib tashlaydi
  //  va foydalanuvchining o'z bo'yoqlarini (avvalgi rangini) qaytaradi.
  // =========================================================================
  var COLORS = { mark: "#FFFF00", clean: "#00FF00" };   // sariq — iz, yashil — tozalangan
  var PREFIX = "_AITC";

  function canMark() { return Office.context.requirements.isSetSupported("WordApi", "1.3"); }
  function canBookmark() { return Office.context.requirements.isSetSupported("WordApi", "1.4"); }

  function encodeColor(c) {
    if (!c || typeof c !== "string" || c.charAt(0) !== "#" || c.length !== 7) return "N";
    var up = c.toUpperCase();
    return up === COLORS.mark || up === COLORS.clean ? "N" : up.slice(1);
  }
  function decodeColor(code) { return code && code !== "N" ? "#" + code : null; }

  async function listOurBookmarks(ctx) {
    var res = ctx.document.body.getRange("Whole").getBookmarks(true, true);
    await ctx.sync();
    return res.value.filter(function (n) { return n.indexOf(PREFIX) === 0; });
  }

  async function clearInContext(ctx) {
    var removed = 0;
    if (canBookmark()) {
      var names = await listOurBookmarks(ctx);
      if (!names.length) return 0;
      var items = names.map(function (n) {
        var r = ctx.document.getBookmarkRangeOrNullObject(n);
        r.load("isNullObject");
        return { n: n, r: r };
      });
      await ctx.sync();
      items.forEach(function (it) {
        if (!it.r.isNullObject) { it.r.font.highlightColor = decodeColor(it.n.split("_").pop()); removed++; }
        ctx.document.deleteBookmark(it.n);
      });
      await ctx.sync();
      return removed;
    }
    // Eski Word: xatcho'psiz — faqat aynan bizning ikki rangimizni olib tashlaymiz
    var toks = ctx.document.body.getTextRanges([" "], false);
    toks.load("items/font/highlightColor");
    await ctx.sync();
    toks.items.forEach(function (t) {
      var c = (t.font.highlightColor || "").toUpperCase();
      if (c === COLORS.mark || c === COLORS.clean) { t.font.highlightColor = null; removed++; }
    });
    await ctx.sync();
    return removed;
  }

  /** Barcha ranglarni o'chirish. Qaytaradi: nechta joy tozalandi. */
  async function clearMarks() {
    if (!canMark()) return 0;
    return Word.run(clearInContext);
  }

  /** Hujjatda hozir nechta belgilangan joy bor (indikator uchun). */
  async function countMarks() {
    if (!canBookmark()) return null;
    return Word.run(async function (ctx) { return (await listOurBookmarks(ctx)).length; });
  }

  /**
   * AI izlarini bo'yash.
   * mode: "mark"  — sariq (qayerda iz bor)
   *       "clean" — yashil (tozalash oldidan; almashtirilgan belgilar rangni meros qiladi)
   * Qaytaradi: {spots:[{name, text, marks, rules}], total, usedSelection}
   */
  async function markTraces(opts, scope, mode) {
    if (!canMark()) throw new Error(tr("errNoMark"));
    var color = mode === "clean" ? COLORS.clean : COLORS.mark;
    return Word.run(async function (ctx) {
      await clearInContext(ctx);                       // eski belgilarni olib, asl ranglarni qaytaramiz
      var s = await resolveScope(ctx, scope);
      var paras = s.range.paragraphs;
      paras.load("items/text");
      await ctx.sync();

      var work = [];
      paras.items.forEach(function (p) {
        var tr = C.traces(p.text, opts);
        if (!tr.length) return;
        var toks = p.getTextRanges([" "], false);
        toks.load("items/text,items/font/highlightColor");
        work.push({ text: p.text, tr: tr, toks: toks });
      });
      if (!work.length) return { spots: [], total: 0, usedSelection: s.usedSelection };
      await ctx.sync();

      var useBm = canBookmark();
      var stamp = Date.now().toString(36).slice(-4);
      var seq = 0, spots = [], total = 0;
      work.forEach(function (w) {
        var items = w.toks.items;
        var spans = C.mapTokens(w.text, items.map(function (t) { return t.text; }));
        C.groupTraces(w.text, w.tr, spans, mode === "clean").forEach(function (g) {
          var tok = items[g.tok];
          var name = PREFIX + stamp + (seq++).toString(36) + "_" + encodeColor(tok.font.highlightColor);
          tok.font.highlightColor = color;
          if (useBm) tok.insertBookmark(name);
          if (g.listed) {
            total++;
            spots.push({ name: useBm ? name : null, text: tok.text, marks: g.marks, rules: g.rules });
          }
        });
      });
      await ctx.sync();
      return { spots: spots, total: total, usedSelection: s.usedSelection };
    });
  }

  /** Ro'yxatdagi joyni hujjatda belgilab ko'rsatish. */
  async function selectMark(name) {
    return Word.run(async function (ctx) {
      var r = ctx.document.getBookmarkRangeOrNullObject(name);
      r.load("isNullObject");
      await ctx.sync();
      if (r.isNullObject) return false;
      r.select();
      await ctx.sync();
      return true;
    });
  }

  /** Tozalash + tozalangan joylarni yashil qilish. */
  async function cleanAndMark(opts, scope, onProgress) {
    var marked = null;
    if (canMark()) {
      (onProgress || function () {})(tr("prMarking"));
      marked = await markTraces(opts, scope, "clean");
    }
    var r = await cleanInWord(opts, scope, onProgress);
    r.spots = marked ? marked.spots : [];
    return r;
  }

  root.WordEngine = {
    analyzeInWord: analyzeInWord, cleanInWord: cleanInWord, cleanAndMark: cleanAndMark,
    markTraces: markTraces, clearMarks: clearMarks, countMarks: countMarks, selectMark: selectMark,
    canMark: canMark, canBookmark: canBookmark, COLORS: COLORS
  };
})(typeof window !== "undefined" ? window : globalThis);
