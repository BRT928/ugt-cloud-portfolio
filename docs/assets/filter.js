
(function () {
  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) { return {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]; }); }
  function init(box) {
    var state = {}, q = "";
    var items = Array.prototype.slice.call(document.querySelectorAll("[data-fitem]"));
    var rows = Array.prototype.slice.call(box.querySelectorAll(".filter-row[data-fkey]"));
    var params = new URLSearchParams(location.search);
    function vals(el, key) { var v = el.getAttribute("data-f-" + key); return v ? v.split("|") : []; }
    rows.forEach(function (row) {
      var key = row.getAttribute("data-fkey");
      var order = JSON.parse(row.getAttribute("data-forder") || "[]");
      var tail = JSON.parse(row.getAttribute("data-ftail") || "[]");
      var labels = JSON.parse(row.getAttribute("data-flabels") || "{}");
      var set = {};
      items.forEach(function (it) { vals(it, key).forEach(function (v) { if (v) set[v] = 1; }); });
      var all = Object.keys(set);
      var list = order.filter(function (v) { return set[v]; })
        .concat(all.filter(function (v) { return order.indexOf(v) < 0 && tail.indexOf(v) < 0; }).sort(function (a, b) { return a.localeCompare(b); }))
        .concat(tail.filter(function (v) { return set[v]; }));
      state[key] = params.get(key) && set[params.get(key)] ? params.get(key) : "";
      var html = '<button type="button" class="chip" data-v="">All</button>';
      list.forEach(function (v) { html += '<button type="button" class="chip" data-v="' + esc(v) + '">' + esc(labels[v] || v) + '</button>'; });
      row.insertAdjacentHTML("beforeend", html);
      row.addEventListener("click", function (e) {
        var b = e.target.closest(".chip[data-v]"); if (!b) return;
        var v = b.getAttribute("data-v");
        state[key] = (state[key] === v && v !== "") ? "" : v;
        apply();
      });
    });
    var input = box.querySelector("input[data-fsearch]");
    if (input) input.addEventListener("input", function () { q = input.value.trim().toLowerCase(); apply(); });
    function apply() {
      rows.forEach(function (row) {
        var key = row.getAttribute("data-fkey");
        row.querySelectorAll(".chip[data-v]").forEach(function (b) {
          var on = b.getAttribute("data-v") === state[key];
          b.classList.toggle("active", on); b.setAttribute("aria-pressed", on);
        });
      });
      var seen = {};
      items.forEach(function (it) {
        var ok = true;
        for (var k in state) { if (state[k] && vals(it, k).indexOf(state[k]) < 0) { ok = false; break; } }
        if (ok && q && (it.getAttribute("data-ftext") || "").indexOf(q) < 0) ok = false;
        it.classList.toggle("fhide", !ok);
        if (ok) seen[it.getAttribute("data-n") || Math.random()] = 1;
      });
      document.querySelectorAll("[data-fsection]").forEach(function (sec) {
        sec.classList.toggle("fhide", !sec.querySelector("[data-fitem]:not(.fhide)"));
      });
      var n = Object.keys(seen).length, c = box.querySelector(".fcount");
      if (c) c.textContent = n + " " + (box.getAttribute("data-fnoun") || "item") + (n === 1 ? "" : "s") + " in view.";
      var empty = document.getElementById("fempty"); if (empty) empty.hidden = n > 0;
      var u = new URL(location.href);
      for (var k2 in state) { if (state[k2]) u.searchParams.set(k2, state[k2]); else u.searchParams.delete(k2); }
      history.replaceState(history.state, "", u);
    }
    apply();
  }
  function boot() { document.querySelectorAll(".ffilters").forEach(init); }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", boot); else boot();
})();
