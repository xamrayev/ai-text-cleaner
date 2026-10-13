/* Sozlamalarni brauzer xotirasida saqlash (taskpane va lenta tugmasi uchun umumiy). */
(function (root) {
  "use strict";
  var KEY = "aitc.options.v1";

  function load() {
    var opts = root.AICleaner.defaultOptions();
    try {
      var saved = JSON.parse(root.localStorage.getItem(KEY) || "{}");
      Object.keys(opts).forEach(function (k) { if (typeof saved[k] === "boolean") opts[k] = saved[k]; });
    } catch (e) { /* xotira mavjud emas — standart sozlamalar */ }
    return opts;
  }

  function save(opts) {
    try { root.localStorage.setItem(KEY, JSON.stringify(opts)); } catch (e) { /* e'tiborsiz */ }
  }

  root.AISettings = { load: load, save: save };
})(window);
