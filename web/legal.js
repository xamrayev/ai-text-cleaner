/* Huquqiy sahifalar: ?lang=en|ru|uz bo'yicha bitta til bo'limini ko'rsatadi. */
(function () {
  "use strict";
  var LANGS = ["en", "ru", "uz"];
  var q = new URLSearchParams(location.search).get("lang");
  var nav = (navigator.language || "").slice(0, 2).toLowerCase();
  var lang = LANGS.indexOf(q) >= 0 ? q : LANGS.indexOf(nav) >= 0 ? nav : "en";
  document.documentElement.className += " js";
  document.documentElement.lang = lang;
  document.querySelectorAll("section[lang]").forEach(function (s) { s.classList.toggle("on", s.lang === lang); });
  document.querySelectorAll("nav.langs a").forEach(function (a) {
    a.href = "?lang=" + a.dataset.lang;
    a.setAttribute("aria-current", String(a.dataset.lang === lang));
  });
  document.querySelectorAll("a[data-page]").forEach(function (a) { a.href = a.dataset.page + "?lang=" + lang; });
})();
