/*
 * Interfeys tarjimalari (en / ru / uz) va tilni aniqlash.
 * Til: foydalanuvchi tanlagani → Office interfeysi tili → brauzer tili → en.
 * Qoidalar nomlarining o'zbekcha matni cleaner.js'da (RULES.label/desc) turadi.
 */
(function (root) {
  "use strict";
  var KEY = "aitc.lang";
  var LANGS = ["en", "ru", "uz"];

  var STR = {
    en: {
      appName: "AI Text Cleaner",
      tagline: "Removes hidden leftovers from ChatGPT, Claude and Gemini text",
      langLabel: "Language",
      tabDoc: "Document", tabText: "Text",
      scopeLabel: "What to clean", scopeSelection: "Selection", scopeDocument: "Whole document",
      scopeHint: "If nothing is selected, the whole document is cleaned.",
      btnShow: "AI traces", btnShowTitle: "Highlight AI traces in the document in yellow",
      btnShowUnsupported: "Highlighting is not supported in this Word version (WordApi 1.3 required)",
      btnClean: "Clean", btnCleanTitle: "Clean; cleaned spots turn green",
      btnClear: "Remove colors",
      indOff: "No colored marks in the document",
      indMark: "{n} spots with AI traces shown (yellow)",
      indClean: "{n} cleaned spots (green)",
      indUnknown: "The document may contain colored marks",
      indCount: "{n} spots are color-marked",
      spotsFound: "Found spots", spotsCleaned: "Cleaned spots",
      legendTrace: "trace", legendClean: "cleaned",
      spotsHint: "Click a spot to jump to it in Word. ¤ hidden character, ⍽ non-breaking space, ␣ space, → tab.",
      emptySpot: "(blank)", more: "Show more ({n})",
      txtIn: "Input text", txtInPlaceholder: "Paste AI-generated text here…",
      btnInsert: "Insert into document", btnInserted: "Inserted ✓",
      txtOut: "Cleaned text", btnCopy: "Copy", btnCopied: "Copied ✓",
      rulesTitle: "Cleaning rules", btnReset: "Restore defaults",
      footer: "Your text is never sent anywhere — everything runs on your computer.",
      lnkPrivacy: "Privacy", lnkTerms: "Terms", lnkSupport: "Support",
      resFoundSel: "Found in selection", resFoundDoc: "Found in document",
      resCleaned: "Cleaned", resResult: "Result", total: "Total",
      stSearching: "Searching for AI traces…",
      stTracesFound: "Yellow spots are AI traces. Click “Clean” to turn them green.",
      stNoTraces: "No AI traces found.",
      stCleaning: "Cleaning…",
      stChanges: "{n} changes made. Green spots are the cleaned ones.",
      stNothing: "Nothing to clean.",
      noteMissed: "Word search could not reach {n} characters (e.g. inside a field or text box). You can clean them in “Text” mode.",
      noteTrack: "All changes are visible under Review.",
      noteUndo: "Press Ctrl+Z to undo.",
      stColorsRemoved: "Colors removed, your original highlighting restored.",
      stNoColors: "No colors to remove.",
      error: "Error: {msg}",
      prParagraphs: "Checking paragraphs…",
      prReplacing: "Replacing characters ({n} kinds)…",
      prLookalike: "Checking look-alike letters…",
      prMarking: "Marking spots to clean…",
      errNoMark: "This Word version does not support highlighting (WordApi 1.3 required).",
      // Lenta tugmalari bildirishnomalari
      ntTitle: "AI Text Cleaner",
      ntOk: "OK",
      ntQuickClean: "{n} changes made in the {where}. Cleaned spots are highlighted in green.",
      ntQuickCleanNone: "Nothing to clean in the {where}.",
      ntTraces: "{n} spots with AI traces highlighted in yellow in the {where}.",
      ntTracesNone: "No AI traces found in the {where}.",
      ntCleared: "Colors removed from {n} spots. Your original highlighting was restored.",
      ntClearedNone: "There were no colors to remove.",
      ntWhereSel: "selection", ntWhereDoc: "document",
      ntError: "Something went wrong: {msg}"
    },
    ru: {
      appName: "AI Очиститель текста",
      tagline: "Удаляет скрытые следы из текста ChatGPT, Claude и Gemini",
      langLabel: "Язык",
      tabDoc: "Документ", tabText: "Текст",
      scopeLabel: "Что очистить", scopeSelection: "Выделенное", scopeDocument: "Весь документ",
      scopeHint: "Если ничего не выделено, очищается весь документ.",
      btnShow: "Следы AI", btnShowTitle: "Подсветить следы AI в документе жёлтым",
      btnShowUnsupported: "Подсветка не поддерживается в этой версии Word (нужен WordApi 1.3)",
      btnClean: "Очистить", btnCleanTitle: "Очистить; очищенные места станут зелёными",
      btnClear: "Убрать цвета",
      indOff: "В документе нет цветных пометок",
      indMark: "Следы AI показаны в {n} местах (жёлтый)",
      indClean: "Очищено мест: {n} (зелёный)",
      indUnknown: "В документе могут быть цветные пометки",
      indCount: "Цветом отмечено мест: {n}",
      spotsFound: "Найденные места", spotsCleaned: "Очищенные места",
      legendTrace: "след", legendClean: "очищено",
      spotsHint: "Нажмите на место — Word перейдёт к нему. ¤ скрытый символ, ⍽ неразрывный пробел, ␣ пробел, → табуляция.",
      emptySpot: "(пусто)", more: "Показать ещё ({n})",
      txtIn: "Исходный текст", txtInPlaceholder: "Вставьте сюда текст от AI…",
      btnInsert: "Вставить в документ", btnInserted: "Вставлено ✓",
      txtOut: "Очищенный текст", btnCopy: "Копировать", btnCopied: "Скопировано ✓",
      rulesTitle: "Правила очистки", btnReset: "Вернуть настройки по умолчанию",
      footer: "Текст никуда не отправляется — всё выполняется на вашем компьютере.",
      lnkPrivacy: "Конфиденциальность", lnkTerms: "Условия", lnkSupport: "Поддержка",
      resFoundSel: "Найдено в выделенном", resFoundDoc: "Найдено в документе",
      resCleaned: "Очищено", resResult: "Результат", total: "Итого",
      stSearching: "Поиск следов AI…",
      stTracesFound: "Жёлтые места — следы AI. Нажмите «Очистить», и они станут зелёными.",
      stNoTraces: "Следы AI не найдены.",
      stCleaning: "Очистка…",
      stChanges: "Внесено изменений: {n}. Зелёные места — очищенные.",
      stNothing: "Нечего очищать.",
      noteMissed: "Поиск Word не нашёл символов: {n} (например, внутри поля или надписи). Их можно очистить в режиме «Текст».",
      noteTrack: "Все изменения видны на вкладке «Рецензирование».",
      noteUndo: "Отменить — Ctrl+Z.",
      stColorsRemoved: "Цвета убраны, прежняя подсветка восстановлена.",
      stNoColors: "Нет цветов для удаления.",
      error: "Ошибка: {msg}",
      prParagraphs: "Проверка абзацев…",
      prReplacing: "Замена символов (видов: {n})…",
      prLookalike: "Проверка похожих букв…",
      prMarking: "Отмечаем места для очистки…",
      errNoMark: "Эта версия Word не поддерживает подсветку (нужен WordApi 1.3).",
      ntTitle: "AI Очиститель текста",
      ntOk: "OK",
      ntQuickClean: "Внесено изменений: {n} ({where}). Очищенные места подсвечены зелёным.",
      ntQuickCleanNone: "Нечего очищать ({where}).",
      ntTraces: "Следы AI подсвечены жёлтым в {n} местах ({where}).",
      ntTracesNone: "Следы AI не найдены ({where}).",
      ntCleared: "Цвета убраны в {n} местах. Прежняя подсветка восстановлена.",
      ntClearedNone: "Цветов для удаления не было.",
      ntWhereSel: "выделенное", ntWhereDoc: "весь документ",
      ntError: "Что-то пошло не так: {msg}"
    },
    uz: {
      appName: "AI Matn Tozalagich",
      tagline: "ChatGPT, Claude, Gemini matnidagi yashirin qoldiqlarni tozalaydi",
      langLabel: "Til",
      tabDoc: "Hujjat", tabText: "Matn",
      scopeLabel: "Qayerni tozalash", scopeSelection: "Belgilangan qism", scopeDocument: "Butun hujjat",
      scopeHint: "Hech narsa belgilanmagan bo‘lsa, butun hujjat tozalanadi.",
      btnShow: "AI izlari", btnShowTitle: "AI izlarini hujjatda sariq rang bilan ko‘rsatish",
      btnShowUnsupported: "Bu Word versiyasida ranglash qo‘llanmaydi (WordApi 1.3 kerak)",
      btnClean: "Tozalash", btnCleanTitle: "Tozalash; tozalangan joylar yashil bo‘ladi",
      btnClear: "Ranglarni o‘chirish",
      indOff: "Hujjatda rangli belgilar yo‘q",
      indMark: "{n} ta joyda AI izi ko‘rsatilgan (sariq)",
      indClean: "{n} ta tozalangan joy (yashil)",
      indUnknown: "Hujjatda rangli belgilar bo‘lishi mumkin",
      indCount: "{n} ta joy rangli belgilangan",
      spotsFound: "Topilgan joylar", spotsCleaned: "Tozalangan joylar",
      legendTrace: "iz", legendClean: "tozalangan",
      spotsHint: "Joyni bosing — Word shu yerga o‘tadi. ¤ yashirin belgi, ⍽ uzilmas bo‘shliq, ␣ bo‘shliq, → tab.",
      emptySpot: "(bo‘sh joy)", more: "Yana ko‘rsatish ({n})",
      txtIn: "Kiruvchi matn", txtInPlaceholder: "AI chiqargan matnni shu yerga qo‘ying…",
      btnInsert: "Hujjatga qo‘yish", btnInserted: "Qo‘yildi ✓",
      txtOut: "Tozalangan matn", btnCopy: "Nusxalash", btnCopied: "Nusxalandi ✓",
      rulesTitle: "Tozalash qoidalari", btnReset: "Standart sozlamalarga qaytarish",
      footer: "Matn hech qayerga yuborilmaydi — hammasi kompyuteringizda bajariladi.",
      lnkPrivacy: "Maxfiylik", lnkTerms: "Shartlar", lnkSupport: "Yordam",
      resFoundSel: "Belgilangan qismda topildi", resFoundDoc: "Hujjatda topildi",
      resCleaned: "Tozalandi", resResult: "Natija", total: "Jami",
      stSearching: "AI izlari qidirilmoqda…",
      stTracesFound: "Sariq joylar — AI izlari. «Tozalash» bosilsa, ular yashil bo‘ladi.",
      stNoTraces: "AI izlari topilmadi.",
      stCleaning: "Tozalanmoqda…",
      stChanges: "{n} ta o‘zgarish kiritildi. Yashil joylar — tozalanganlar.",
      stNothing: "Tozalanadigan narsa topilmadi.",
      noteMissed: "{n} ta belgini Word qidiruvi topa olmadi (masalan, maydon yoki matn qutisi ichida). Ularni «Matn» rejimida tozalashingiz mumkin.",
      noteTrack: "Barcha o‘zgarishlar «Рецензирование / Review» bo‘limida ko‘rinadi.",
      noteUndo: "Bekor qilish uchun Ctrl+Z.",
      stColorsRemoved: "Ranglar o‘chirildi, avvalgi bo‘yoqlar tiklandi.",
      stNoColors: "O‘chiriladigan rang yo‘q.",
      error: "Xatolik: {msg}",
      prParagraphs: "Paragraflar tekshirilmoqda…",
      prReplacing: "Belgilar almashtirilmoqda ({n} tur)…",
      prLookalike: "O‘xshash harflar tekshirilmoqda…",
      prMarking: "Tozalanadigan joylar belgilanmoqda…",
      errNoMark: "Bu Word versiyasi ranglash funksiyasini qo‘llamaydi (WordApi 1.3 kerak).",
      ntTitle: "AI Matn Tozalagich",
      ntOk: "OK",
      ntQuickClean: "{where}: {n} ta o‘zgarish kiritildi. Tozalangan joylar yashil rangda.",
      ntQuickCleanNone: "{where}: tozalanadigan narsa topilmadi.",
      ntTraces: "{where}: {n} ta joyda AI izi sariq rang bilan ko‘rsatildi.",
      ntTracesNone: "{where}: AI izlari topilmadi.",
      ntCleared: "{n} ta joydan rang olindi, avvalgi bo‘yoqlar tiklandi.",
      ntClearedNone: "O‘chiriladigan rang yo‘q edi.",
      ntWhereSel: "Belgilangan qism", ntWhereDoc: "Butun hujjat",
      ntError: "Xatolik yuz berdi: {msg}"
    }
  };

  // Qoidalar (en/ru); uz matni cleaner.js'dan olinadi
  var RULES = {
    en: {
      hidden: ["Hidden characters", "Zero-width space, BOM, soft hyphen, direction marks, Unicode tags (emoji are kept)"],
      nbsp: ["Non-breaking spaces", "Replace NBSP and narrow spaces with a regular space"],
      dashes: ["Dashes", "Turn — and – into a plain hyphen (-)"],
      quotes: ["Curly quotes", "“ ” ‘ ’ → \" and '"],
      ellipsis: ["Ellipsis", "Turn … into three dots (...)"],
      trailing: ["Trailing spaces", "Extra spaces and tabs at the end of each paragraph"],
      asterisks: ["Asterisks (*)", "Remove all * characters left over from Markdown"],
      headings: ["Markdown headings", "Remove # … ###### at the start of a paragraph"],
      lookalike: ["Look-alike letters", "Fix Cyrillic/Greek letters mixed into Latin words"],
      nfkc: ["Unicode normalization (NFKC)", "Ｈｅｌｌｏ, ﬁ, 𝐇𝐞𝐥𝐥𝐨 → plain form (², ™ change too)"],
      uzApostrophe: ["Keep Uzbek o‘, g‘ and the glottal stop", "Leaves ‘ ’ in words like o‘, g‘, ma’no untouched"],
      guillemets: ["Also replace «» quotes", "« » → \" and ‹ › → '"],
      trackChanges: ["Use Track Changes", "Every fix is marked for review in Word"]
    },
    ru: {
      hidden: ["Скрытые символы", "Пробел нулевой ширины, BOM, мягкий перенос, знаки направления, теги Unicode (эмодзи сохраняются)"],
      nbsp: ["Неразрывные пробелы", "Заменить NBSP и узкие пробелы обычным пробелом"],
      dashes: ["Тире", "Превратить — и – в обычный дефис (-)"],
      quotes: ["Фигурные кавычки", "“ ” ‘ ’ → \" и '"],
      ellipsis: ["Многоточие", "Превратить … в три точки (...)"],
      trailing: ["Пробелы в конце строк", "Лишние пробелы и табуляции в конце каждого абзаца"],
      asterisks: ["Звёздочки (*)", "Удалить все символы *, оставшиеся от Markdown"],
      headings: ["Заголовки Markdown", "Удалить # … ###### в начале абзаца"],
      lookalike: ["Похожие буквы", "Исправить кириллицу/греческие буквы внутри латинских слов"],
      nfkc: ["Нормализация Unicode (NFKC)", "Ｈｅｌｌｏ, ﬁ, 𝐇𝐞𝐥𝐥𝐨 → обычная форма (², ™ тоже меняются)"],
      uzApostrophe: ["Сохранять узбекские o‘, g‘ и тутук", "Не трогает ‘ ’ в словах вроде o‘, g‘, ma’no"],
      guillemets: ["Заменять и кавычки «»", "« » → \" и ‹ › → '"],
      trackChanges: ["В режиме исправлений", "Каждое исправление отмечается для рецензирования в Word"]
    }
  };

  function norm(code) {
    var c = String(code || "").toLowerCase();
    for (var i = 0; i < LANGS.length; i++) if (c.indexOf(LANGS[i]) === 0) return LANGS[i];
    return null;
  }

  function stored() {
    try { return norm(root.localStorage.getItem(KEY)); } catch (e) { return null; }
  }

  function officeLang() {
    try { return norm(root.Office.context.displayLanguage); } catch (e) { return null; }
  }

  var current = stored() || norm(root.navigator && root.navigator.language) || "en";

  /** Office tayyor bo'lgach chaqiriladi: Office tili brauzer tilidan ustun. */
  function detect() {
    current = stored() || officeLang() || norm(root.navigator && root.navigator.language) || "en";
    return current;
  }

  /** Tilni o'rnatish; persist=false bo'lsa tanlov eslab qolinmaydi (masalan, dialog oynada). */
  function setLang(code, persist) {
    current = norm(code) || "en";
    if (persist !== false) {
      try { root.localStorage.setItem(KEY, current); } catch (e) { /* e'tiborsiz */ }
    }
    return current;
  }

  function t(key, params) {
    var s = STR[current][key];
    if (s === undefined) s = STR.en[key];
    if (s === undefined) return key;
    return s.replace(/\{(\w+)\}/g, function (m, k) {
      return params && params[k] !== undefined ? String(params[k]) : m;
    });
  }

  /** Qoida nomi va izohi joriy tilda: {label, desc}. */
  function rule(r) {
    var tr = RULES[current] && RULES[current][r.id];
    return tr ? { label: tr[0], desc: tr[1] } : { label: r.label, desc: r.desc };
  }

  function num(n) {
    try { return Number(n).toLocaleString(current); } catch (e) { return String(n); }
  }

  /** data-i18n, data-i18n-title, data-i18n-placeholder atributlari bo'yicha sahifani tarjima qilish. */
  function apply(doc) {
    doc.documentElement.lang = current;
    doc.querySelectorAll("[data-i18n]").forEach(function (el) { el.textContent = t(el.dataset.i18n); });
    doc.querySelectorAll("[data-i18n-title]").forEach(function (el) { el.title = t(el.dataset.i18nTitle); });
    doc.querySelectorAll("[data-i18n-placeholder]").forEach(function (el) { el.placeholder = t(el.dataset.i18nPlaceholder); });
  }

  root.I18N = {
    LANGS: LANGS, STR: STR, RULES: RULES,
    t: t, rule: rule, num: num, apply: apply, detect: detect, setLang: setLang,
    lang: function () { return current; }
  };
})(typeof window !== "undefined" ? window : globalThis);
