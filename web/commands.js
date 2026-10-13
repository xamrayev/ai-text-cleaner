/* Lenta (ribbon) tugmalari — panelni ochmasdan ishlaydi, saqlangan sozlamalardan foydalanadi.
 * Natija kichik dialog oynada ko'rsatiladi (Word'da function command uchun boshqa bildirishnoma API yo'q). */
(function () {
  "use strict";
  var t = window.I18N.t, num = window.I18N.num;
  var NOTICE_MS = 6000;   // bildirishnoma o'zi yopiladigan vaqt

  function log(msg) { console.log("[AI Text Cleaner] " + msg); }

  function where(usedSelection) { return t(usedSelection ? "ntWhereSel" : "ntWhereDoc"); }

  /** Natijani dialogda ko'rsatadi; dialog yopilganda (yoki ochilmasa) resolve bo'ladi. */
  function notify(text, kind) {
    log(text);
    return new Promise(function (resolve) {
      var ui = Office.context && Office.context.ui;
      var ok = Office.context.requirements && Office.context.requirements.isSetSupported("DialogApi", "1.1");
      if (!ui || !ok) return resolve();
      var url = new URL("notice.html", location.href);
      url.searchParams.set("lang", window.I18N.lang());
      url.searchParams.set("msg", text);
      if (kind) url.searchParams.set("kind", kind);
      ui.displayDialogAsync(url.href, { height: 24, width: 28, displayInIframe: true }, function (res) {
        if (res.status !== Office.AsyncResultStatus.Succeeded) { log("Dialog: " + res.error.message); return resolve(); }
        var dlg = res.value, done = false;
        function close() {
          if (done) return;
          done = true;
          try { dlg.close(); } catch (e) { /* allaqachon yopilgan */ }
          resolve();
        }
        dlg.addEventHandler(Office.EventType.DialogMessageReceived, close);
        dlg.addEventHandler(Office.EventType.DialogEventReceived, close);   // foydalanuvchi X ni bosdi
        setTimeout(close, NOTICE_MS);
      });
    });
  }

  function command(fn) {
    return async function (event) {
      try {
        await notify(await fn());
      } catch (e) {
        await notify(t("ntError", { msg: e.message || e }), "err");
      } finally {
        event.completed();
      }
    };
  }

  // Tez tozalash: tozalangan joylar yashil bo'ladi
  var quickClean = command(async function () {
    var r = await WordEngine.cleanAndMark(AISettings.load(), "selection");
    var n = Math.round(AICleaner.total(r.counts));
    return n ? t("ntQuickClean", { n: num(n), where: where(r.usedSelection) })
             : t("ntQuickCleanNone", { where: where(r.usedSelection) });
  });

  // AI izlarini sariq rang bilan ko'rsatish
  var showTraces = command(async function () {
    var r = await WordEngine.markTraces(AISettings.load(), "selection", "mark");
    return r.total ? t("ntTraces", { n: num(r.total), where: where(r.usedSelection) })
                   : t("ntTracesNone", { where: where(r.usedSelection) });
  });

  // Barcha ranglarni o'chirish (asl bo'yoqlar tiklanadi)
  var clearColors = command(async function () {
    var n = await WordEngine.clearMarks();
    return n ? t("ntCleared", { n: num(n) }) : t("ntClearedNone");
  });

  // Eski va yangi Office versiyalari uchun ikkala usulda ro'yxatdan o'tkazamiz
  window.quickClean = quickClean;
  window.showTraces = showTraces;
  window.clearColors = clearColors;
  Office.onReady(function () {
    window.I18N.detect();
    if (Office.actions && Office.actions.associate) {
      Office.actions.associate("quickClean", quickClean);
      Office.actions.associate("showTraces", showTraces);
      Office.actions.associate("clearColors", clearColors);
    }
  });
})();
