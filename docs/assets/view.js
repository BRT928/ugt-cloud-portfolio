
(function () {
  var KEY = "ugt-portfolio-view";
  var VALID = { cards: 1, table: 1, both: 1 };
  function read() {
    var q = new URLSearchParams(location.search).get("view");
    if (q && VALID[q]) return q;
    try {
      var s = localStorage.getItem(KEY);
      if (s && VALID[s]) return s;
    } catch (e) {}
    return "both";
  }
  function apply(mode) {
    document.body.classList.remove("view-cards", "view-table", "view-both");
    document.body.classList.add("view-" + mode);
    document.querySelectorAll("[data-view-toggle] button[data-view]").forEach(function (btn) {
      btn.classList.toggle("active", btn.getAttribute("data-view") === mode);
    });
  }
  function set(mode, updateUrl) {
    if (!VALID[mode]) mode = "both";
    try { localStorage.setItem(KEY, mode); } catch (e) {}
    apply(mode);
    if (updateUrl !== false) {
      var u = new URL(location.href);
      if (mode === "both") u.searchParams.delete("view");
      else u.searchParams.set("view", mode);
      history.replaceState({}, "", u);
    }
  }
  function appendViewParam(href) {
    var mode = read();
    if (mode === "both" || !href) return href;
    try {
      var u = new URL(href, location.href);
      u.searchParams.set("view", mode);
      // keep relative for same-directory pages
      var file = u.pathname.split("/").pop() || "index.html";
      return file + u.search + u.hash;
    } catch (e) {
      return href;
    }
  }
  window.ugtView = { read: read, set: set, apply: apply, appendViewParam: appendViewParam };
  function boot() {
    var q = new URLSearchParams(location.search).get("view");
    if (q && VALID[q]) {
      try { localStorage.setItem(KEY, q); } catch (e) {}
    }
    apply(read());
    document.querySelectorAll("[data-view-toggle]").forEach(function (root) {
      root.addEventListener("click", function (e) {
        var btn = e.target.closest("button[data-view]");
        if (!btn) return;
        set(btn.getAttribute("data-view"));
      });
    });
    document.querySelectorAll("a[href]").forEach(function (a) {
      var href = a.getAttribute("href") || "";
      if (!/products\.html|roadmap\.html|owner-/.test(href)) return;
      a.addEventListener("click", function () {
        var next = appendViewParam(a.getAttribute("href"));
        if (next) a.setAttribute("href", next);
      });
    });
    document.querySelectorAll("section.section[data-collapsible]").forEach(function (sec) {
      var head = sec.querySelector(".section-head");
      if (!head) return;
      head.addEventListener("click", function () {
        sec.classList.toggle("is-collapsed");
        head.setAttribute("aria-expanded", sec.classList.contains("is-collapsed") ? "false" : "true");
      });
    });
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", boot);
  else boot();
})();
