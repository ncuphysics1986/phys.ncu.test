/* ════════════════════════════════════════════════════════════════
   boot.js — app.js 版本載入器
   ────────────────────────────────────────────────────────────────
   作用：先以 no-cache 讀取 app-version.txt，再據此載入對應版本的 app.js。
   解決快取問題：只要改 app-version.txt 的版本，瀏覽器就會抓新的 app.js，
   不必每次強制重載（版本檔本身以 no-cache 讀取，故修改後立即生效）。

   app-version.txt 內容可為兩種寫法：
     1) 純版本字串（最簡單），例如：
          2026-06-15
        → 載入 app.js?v=2026-06-15
     2) JSON（可指定載入「不同檔名」以回退舊版），例如：
          {"file":"app-2026-06-01.js","v":"2026-06-01"}
        → 載入 app-2026-06-01.js?v=2026-06-01

   若版本檔讀取失敗，會退回直接載入 app.js，網站不致空白。
═════════════════════════════════════════════════════════════════ */
(function () {
  function load(file, v) {
    var s = document.createElement('script');
    s.src = (file || 'app.js') + (v ? '?v=' + encodeURIComponent(v) : '');
    document.body.appendChild(s);
  }
  fetch('app-version.txt?_=' + Date.now(), { cache: 'no-cache' })
    .then(function (r) { return r.ok ? r.text() : ''; })
    .then(function (txt) {
      var file = 'app.js', v = '';
      txt = (txt || '').trim();
      if (txt.charAt(0) === '{') {
        try { var c = JSON.parse(txt); file = c.file || file; v = c.v || ''; } catch (e) {}
      } else if (txt) {
        v = txt;
      }
      load(file, v);
    })
    .catch(function () { load('app.js', ''); });
})();
