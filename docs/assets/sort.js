window.UGT_STAGE_ORDER = ["Wishlist", "Development", "Deployment", "Pre-Production", "Production", "Terminated"];

(function () {
  var STAGES = window.UGT_STAGE_ORDER || [];
  var coll = new Intl.Collator(["ka", "en"], { numeric: true, sensitivity: "base" });
  function cellText(tr, i) { var c = tr.cells[i]; if (!c) return ""; var t = c.textContent.trim(); return (t === "—" || t === "-") ? "" : t; }
  function key(v, type) {
    if (v === "" || /^(unscheduled|unassigned)$/i.test(v)) return null;
    if (type === "stage") { var i = STAGES.indexOf(v); return i < 0 ? 900 : i; }
    if (type === "due") { var m = v.match(/(\d{4})\s*Q([1-4])/i); return m ? (+m[1]) * 10 + (+m[2]) : null; }
    if (type === "num") { var n = parseFloat(v.replace(/[^\d.\-]/g, "")); return isNaN(n) ? null : n; }
    return v;
  }
  function typeOf(th, table, i) {
    var h = (th.firstChild && th.firstChild.nodeType === 3 ? th.firstChild.textContent : th.textContent).replace(/[↕▲▼]/g, "").trim().toLowerCase();
    if (h.indexOf("stage") === 0) return "stage";
    if (h === "due" || h.indexOf("due") === 0 || h.indexOf("quarter") >= 0) return "due";
    var rows = table.tBodies[0] ? table.tBodies[0].rows : [], n = 0, num = 0;
    for (var r = 0; r < rows.length && n < 30; r++) { var v = cellText(rows[r], i); if (v) { n++; if (/^#?\d+(\.\d+)?$/.test(v)) num++; } }
    return n && num === n ? "num" : "text";
  }
  function setup(table) {
    if (table.dataset.sortable) return; var thead = table.tHead, tb = table.tBodies[0];
    if (!thead || !tb) return; table.dataset.sortable = "1";
    var st = { col: -1, dir: 0 }, mo = null;
    function stamp() { Array.prototype.forEach.call(tb.rows, function (r, i) { if (r.dataset.oi == null) r.dataset.oi = i; }); }
    function apply() {
      if (mo) mo.disconnect(); stamp();
      var rows = Array.prototype.slice.call(tb.rows);
      if (st.dir === 0) rows.sort(function (a, b) { return a.dataset.oi - b.dataset.oi; });
      else {
        var t = st.type;
        rows.sort(function (a, b) {
          var ka = key(cellText(a, st.col), t), kb = key(cellText(b, st.col), t);
          if (ka === null && kb === null) return a.dataset.oi - b.dataset.oi;
          if (ka === null) return 1; if (kb === null) return -1;
          var c = (typeof ka === "number" && typeof kb === "number") ? ka - kb : coll.compare(String(ka), String(kb));
          return (c || (a.dataset.oi - b.dataset.oi)) * (c ? st.dir : 1);
        });
      }
      rows.forEach(function (r) { tb.appendChild(r); });
      Array.prototype.forEach.call(thead.rows[0].cells, function (th, i) {
        th.setAttribute("aria-sort", i === st.col && st.dir ? (st.dir > 0 ? "ascending" : "descending") : "none");
        var a = th.querySelector(".sort-ind"); if (a) a.textContent = i === st.col && st.dir ? (st.dir > 0 ? "▲" : "▼") : "↕";
      });
      if (mo) { mo.takeRecords(); mo.observe(tb, { childList: true }); }
    }
    Array.prototype.forEach.call(thead.rows[0].cells, function (th, i) {
      th.classList.add("sortable-th");
      var s = document.createElement("span"); s.className = "sort-ind"; s.textContent = "↕"; th.appendChild(s);
      th.addEventListener("click", function (e) {
        if (e.target.closest(".col-resizer") || document.body.classList.contains("col-resizing")) return;
        if (st.col !== i) { st = { col: i, dir: 1, type: typeOf(th, table, i) }; }
        else st.dir = st.dir === 1 ? -1 : 0;
        apply();
      });
    });
    mo = new MutationObserver(function () { { Array.prototype.forEach.call(tb.rows, function (r) { delete r.dataset.oi; }); if (st.dir) apply(); else stamp(); } });
    mo.observe(tb, { childList: true });
  }
  function boot() { document.querySelectorAll("table").forEach(function (t) { if (!t.classList.contains("matrix") || true) setup(t); }); }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", boot); else boot();
})();
