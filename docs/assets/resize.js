
/* Drag-to-resize table columns. Widths persist per table type in localStorage.
   Double-click a handle to reset that table to automatic widths. */
(function () {
  var PREFIX = "ugt-colw:v1:";
  function pageKind() {
    var f = (location.pathname.split("/").pop() || "index.html").replace(/\.html$/, "");
    if (/^owner-/.test(f)) return "owner";
    if (/^product-\d+/.test(f)) return "product";
    return f || "index";
  }
  function headCells(t) {
    var row = t.tHead && t.tHead.rows[t.tHead.rows.length - 1];
    return row ? Array.prototype.slice.call(row.cells) : [];
  }
  function sig(t) {
    return headCells(t).map(function (th) { return (th.textContent || "").trim(); }).join("|");
  }
  function keyOf(t) { return PREFIX + pageKind() + ":" + sig(t); }
  function load(k) { try { return JSON.parse(localStorage.getItem(k) || "null"); } catch (e) { return null; } }
  function save(k, w) { try { localStorage.setItem(k, JSON.stringify(w)); } catch (e) {} }
  function clear(k) { try { localStorage.removeItem(k); } catch (e) {} }
  function sameTables(t) {
    var k = keyOf(t);
    return Array.prototype.filter.call(document.querySelectorAll("table.resizable"), function (x) { return keyOf(x) === k; });
  }
  function ensureCols(t) {
    var cg = t.querySelector("colgroup.rz");
    var n = headCells(t).length;
    if (!cg) { cg = document.createElement("colgroup"); cg.className = "rz"; t.insertBefore(cg, t.firstChild); }
    while (cg.children.length < n) cg.appendChild(document.createElement("col"));
    return Array.prototype.slice.call(cg.children, 0, n);
  }
  function apply(t, widths) {
    var cols = ensureCols(t), sum = 0;
    cols.forEach(function (c, i) { var w = Math.max(48, Math.round(widths[i] || 120)); c.style.width = w + "px"; sum += w; });
    t.style.width = sum + "px";
    t.classList.add("rz-fixed");
  }
  function measure(t) {
    t.classList.remove("rz-fixed");
    t.style.width = "";
    ensureCols(t).forEach(function (c) { c.style.width = ""; });
    var avail = t.parentElement ? t.parentElement.clientWidth : 0;
    var w = headCells(t).map(function (th) { return Math.min(480, Math.max(64, th.getBoundingClientRect().width)); });
    var sum = w.reduce(function (a, b) { return a + b; }, 0);
    if (avail && sum < avail) { var extra = (avail - sum) / w.length; w = w.map(function (x) { return x + extra; }); }
    return w;
  }
  function init(t) {
    if (t.dataset.rzReady) return;
    if (!t.offsetWidth) return;              // hidden (e.g. Cards view) — retried when it becomes visible
    t.dataset.rzReady = "1";
    var saved = load(keyOf(t));
    var cells = headCells(t);
    apply(t, saved && saved.length === cells.length ? saved : measure(t));
    cells.forEach(function (th, i) {
      if (th.querySelector(".col-resizer")) return;
      var h = document.createElement("span");
      h.className = "col-resizer";
      h.setAttribute("role", "separator");
      h.setAttribute("aria-orientation", "vertical");
      h.title = "Drag to resize · double-click to reset widths";
      th.appendChild(h);
      h.addEventListener("pointerdown", function (e) {
        e.preventDefault(); e.stopPropagation();
        var cols = ensureCols(t);
        var startX = e.clientX, startW = cols[i].getBoundingClientRect().width || parseFloat(cols[i].style.width) || 120;
        var widths = cols.map(function (c) { return parseFloat(c.style.width) || c.getBoundingClientRect().width; });
        h.setPointerCapture(e.pointerId);
        document.body.classList.add("rz-dragging");
        function move(ev) { widths[i] = Math.max(48, startW + ev.clientX - startX); sameTables(t).forEach(function (x) { apply(x, widths); }); }
        function up() {
          h.removeEventListener("pointermove", move); h.removeEventListener("pointerup", up); h.removeEventListener("pointercancel", up);
          document.body.classList.remove("rz-dragging");
          save(keyOf(t), widths.map(Math.round));
        }
        h.addEventListener("pointermove", move); h.addEventListener("pointerup", up); h.addEventListener("pointercancel", up);
      });
      h.addEventListener("click", function (e) { e.stopPropagation(); });
      h.addEventListener("dblclick", function (e) {
        e.preventDefault(); e.stopPropagation();
        clear(keyOf(t));
        sameTables(t).forEach(function (x) { if (x.offsetWidth) apply(x, measure(x)); });
      });
    });
  }
  function setup() {
    var tables = document.querySelectorAll(".table-wrap > table:not(.matrix)");
    var ro = window.ResizeObserver ? new ResizeObserver(function (es) { es.forEach(function (en) { init(en.target); }); }) : null;
    Array.prototype.forEach.call(tables, function (t) {
      if (!headCells(t).length) return;
      t.classList.add("resizable");
      init(t);
      if (ro) ro.observe(t);
    });
    window.ugtResizeInit = function () { Array.prototype.forEach.call(document.querySelectorAll("table.resizable"), init); };
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", setup); else setup();
})();
