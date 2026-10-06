#!/usr/bin/env python3
"""Build the full UGT Cloud Portfolio static site from Excel."""

from __future__ import annotations

import json
import re
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from zoneinfo import ZoneInfo

import openpyxl

ROOT = Path(__file__).resolve().parent
EXCEL = ROOT / "UGT_Product_Portfolio_2026.xlsx"
STATE = ROOT / "state.json"
OUT_JSON = ROOT / "products.json"
SITE = ROOT / "docs"  # GitHub Pages /docs

STAGE_FIX = {"Pre-Prodcution": "Pre-Production"}
HEAD_FIX = {"ALL": "Cross-portfolio"}
STAGES = [
    "Production",
    "Pre-Production",
    "Development",
    "Pre-Development",
    "Pipeline",
    "Terminated",
]
PLATFORM_COLS = [
    "Universal Cloud",
    "Public Cloud",
    "Private Cloud",
    "Bare Metal",
    "SMB",
    "Standalone",
]
OWNER_BLURBS = {
    "TGA": ("Commercial Product Owner", "Universal Cloud, Bare Metal, storage, network and security"),
    "GMA": ("Commercial Product Owner", "Private Cloud platforms and professional services"),
    "GNA": ("Commercial Product Owner", "SMB packaged cloud and VPS"),
    "SME": ("Technical Product Owner", "Private Cloud, storage, bare metal and connectivity"),
    "AJA": ("Technical Product Owner", "Compute, network, security, co-location and services"),
    "GBU": ("Technical Product Owner", "Backup, GPU and SMB compute"),
    "VMD": ("Technical Product Owner", "Cloud Metal and Canonical MAAS"),
    "ZAN": ("Technical Product Owner", "Microsoft 365 backup"),
}
GLOSSARY = [
    ("Head Product", "The platform or commercial environment where the product is offered. A product may have more than one head product."),
    ("Universal Cloud", "UGT’s universal cloud / virtual data centre platform."),
    ("Public Cloud", "Public cloud environment / platform."),
    ("Private Cloud", "Dedicated private cloud environment."),
    ("Bare Metal", "Physical dedicated server and appliance portfolio."),
    ("SMB", "SMB-focused packaged cloud and services portfolio."),
    ("Standalone", "Product that can be sold independently of the cloud platforms above."),
    ("ALL / Cross-portfolio", "Cross-portfolio service applicable across platforms."),
    ("CPO — Commercial Product Owner", "The person accountable for the commercial offer: positioning, packaging and go-to-market."),
    ("TPO — Technical Product Owner", "The person accountable for technical delivery, architecture and operational readiness."),
    ("Production", "Live / operational product."),
    ("Pre-Production", "Final preparation before production."),
    ("Development", "Actively being built."),
    ("Pre-Development", "Concept, discovery or preparation before development."),
    ("Pipeline", "Planned product / backlog item."),
    ("Terminated", "No longer active in the portfolio."),
    ("In Delivery", "Executive roll-up of Pre-Production plus Development."),
    ("Due Date", "Target year and quarter from the register, for example 2026 Q4. Unscheduled means the team has not set one yet."),
]


def clean_text(v) -> str:
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v).replace("\xa0", " ")).strip()


def split_heads(raw) -> list[str]:
    if raw is None:
        return []
    parts = re.split(r"[\n,/]+", str(raw).replace("\r", ""))
    out = []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        out.append(HEAD_FIX.get(p, p))
    return out


def owner_or_none(v) -> str | None:
    s = clean_text(v)
    if not s or s.upper() in {"N/A", "NA", "-", "—"}:
        return None
    return s


def tech_display(tech: str, sub: str) -> str:
    return " ".join(b for b in (tech, sub) if b)


def load_products(path: Path) -> list[dict]:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["UGT Product Portfolio ALL"]
    headers = [ws.cell(3, c).value for c in range(1, 11)]
    products = []
    for r in range(4, ws.max_row + 1):
        row = {headers[c - 1]: ws.cell(r, c).value for c in range(1, 11)}
        name = clean_text(row.get("Product"))
        if not name and row.get("N") is None:
            continue
        stage_raw = clean_text(row.get("Stage"))
        stage = STAGE_FIX.get(stage_raw, stage_raw)
        heads = split_heads(row.get("Head Product"))
        if not heads:
            heads = ["Unplaced"]
        due = clean_text(row.get("Due Date")) or None
        tech = clean_text(row.get("Platform/Technology/Vendor"))
        sub = clean_text(row.get("Sub Products"))
        n = row.get("N")
        try:
            n = int(n)
        except (TypeError, ValueError):
            n = len(products) + 1
        products.append(
            {
                "n": n,
                "product": name,
                "headProducts": heads,
                "category": clean_text(row.get("Category")),
                "technology": tech_display(tech, sub),
                "stage": stage,
                "due": due,
                "cpo": owner_or_none(row.get("CPO")),
                "tpo": owner_or_none(row.get("TPO")),
            }
        )
    return products


def data_updated_label() -> str:
    """Asia/Tbilisi timestamp from state.json lastModifiedDateTime, else now."""
    tz = ZoneInfo("Asia/Tbilisi")
    iso = None
    if STATE.exists():
        try:
            iso = json.loads(STATE.read_text()).get("lastModifiedDateTime")
        except Exception:
            iso = None
    if iso:
        dt = datetime.fromisoformat(iso.replace("Z", "+00:00")).astimezone(tz)
    else:
        dt = datetime.now(tz)
    return dt.strftime("%d %b %Y, %H:%M GET")


def esc(s) -> str:
    if s is None:
        s = ""
    return (
        str(s)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&#39;")
    )


def badge_class(stage: str) -> str:
    return "badge-" + stage.lower().replace(" ", "-")


def owner_chip_html(code: str | None, role: str, base: str = ".") -> str:
    if not code:
        return (
            f'<span class="owner"><span class="avatar muted">—</span>'
            f'<span><div>Unassigned</div><div class="role">{role}</div></span></span>'
        )
    return (
        f'<a class="owner" href="{base}/owner-{esc(code)}.html">'
        f'<span class="avatar">{esc(code)}</span>'
        f'<span><div>{esc(code)}</div><div class="role">{role}</div></span></a>'
    )


def product_card_html(p: dict, base: str = ".") -> str:
    heads = "".join(f'<span class="pill">{esc(h)}</span>' for h in p["headProducts"])
    due = p["due"] or "Unscheduled"
    tech = f" · {esc(p['technology'])}" if p.get("technology") else ""
    return f"""
    <a class="card" href="{base}/product-{p['n']}.html">
      <div class="card-top">
        <div class="card-title">{esc(p['product'])}</div>
        <span class="badge {badge_class(p['stage'])}">{esc(p['stage'])}</span>
      </div>
      <div class="meta">{esc(p['category'] or '—')}</div>
      <div class="heads">{heads}</div>
      <div class="meta">{esc(due)}{tech}</div>
      <div class="owners">{owner_chip_html(p['cpo'], 'CPO', base)}{owner_chip_html(p['tpo'], 'TPO', base)}</div>
    </a>"""


def product_table_rows(products: list[dict], base: str = ".") -> str:
    rows = []
    for p in products:
        rows.append(
            f"<tr>"
            f'<td><a href="{base}/product-{p["n"]}.html"><strong>{esc(p["product"])}</strong></a></td>'
            f"<td>{esc(p['category'] or '—')}</td>"
            f"<td>{esc(', '.join(p['headProducts']))}</td>"
            f'<td><span class="badge {badge_class(p["stage"])}">{esc(p["stage"])}</span></td>'
            f"<td>{esc(p['due'] or 'Unscheduled')}</td>"
            f"<td>{esc(p['cpo'] or 'Unassigned')}</td>"
            f"<td>{esc(p['tpo'] or 'Unassigned')}</td>"
            f"<td>{esc(p['technology'] or '—')}</td>"
            f"</tr>"
        )
    return "\n".join(rows)


SHARED_CSS = r"""
:root {
  --bg: #eef5fb; --rail: #e4eef8; --card: #ffffff; --fg: #152033;
  --muted: #5b6b7c; --border: #d5e0ec; --accent: #1e6fd9; --accent-fg: #fff;
  --secondary: #e8f0f9; --warn: #9a3412;
  --shadow: 0 0 0 1px rgba(21,32,51,.06), 0 1px 2px rgba(21,32,51,.04);
  --shadow-hover: 0 0 0 1px rgba(30,111,217,.18), 0 8px 24px rgba(21,32,51,.08);
  --radius: 12px;
}
* { box-sizing: border-box; }
html, body { margin: 0; padding: 0; }
body {
  font-family: Figtree, ui-sans-serif, system-ui, sans-serif;
  background: var(--bg); color: var(--fg); min-height: 100vh; line-height: 1.45;
}
a { color: inherit; }
.font-display { font-family: Fraunces, Georgia, serif; }
.app { display: grid; grid-template-columns: 16.5rem 1fr; min-height: 100vh; }
@media (max-width: 960px) {
  .app { grid-template-columns: 1fr; }
  aside.rail { display: none !important; }
}
aside.rail {
  position: sticky; top: 0; height: 100vh; background: var(--rail);
  border-right: 1px solid var(--border); padding: 1.25rem;
  display: flex; flex-direction: column;
}
.brand { display: flex; align-items: center; gap: .75rem; text-decoration: none; color: inherit; }
.logo {
  width: 2.25rem; height: 2.25rem; border-radius: 8px; background: var(--accent);
  color: var(--accent-fg); display: grid; place-items: center; font-weight: 700; font-size: .85rem;
}
.brand .sub { font-size: 11px; letter-spacing: .18em; text-transform: uppercase; color: var(--muted); }
nav.side { margin-top: 2rem; display: flex; flex-direction: column; gap: .25rem; flex: 1; }
nav.side a {
  display: flex; align-items: center; gap: .75rem; min-height: 2.75rem;
  padding: 0 .75rem; border-radius: 8px; text-decoration: none;
  color: var(--muted); font-size: .9rem;
}
nav.side a.active, nav.side a:hover { background: var(--secondary); color: var(--fg); }
.snap { background: var(--secondary); border-radius: 10px; padding: 1rem; }
.snap .label { font-size: 11px; letter-spacing: .16em; text-transform: uppercase; color: var(--muted); }
.snap .num { font-family: Fraunces, serif; font-size: 1.85rem; margin-top: .25rem; }
.snap .hint { font-size: .875rem; color: var(--muted); margin-top: .25rem; }
.updated { font-size: .7rem; color: var(--muted); margin-top: .6rem; line-height: 1.3; }
header.top {
  position: sticky; top: 0; z-index: 20; display: flex; align-items: center; gap: .75rem;
  padding: .75rem 2rem; border-bottom: 1px solid var(--border);
  background: rgba(238,245,251,.92); backdrop-filter: blur(8px);
}
header.top .stats { color: var(--muted); font-size: .875rem; }
.content { padding: 1.5rem 2rem 3rem; }
h1 { font-family: Fraunces, serif; font-weight: 500; font-size: 2.5rem; margin: 0; letter-spacing: -0.02em; }
h2 { font-family: Fraunces, serif; font-weight: 500; font-size: 1.65rem; margin: 0; }
h3 { margin: 0; font-size: 1rem; }
.lede { color: var(--muted); margin: .75rem 0 0; max-width: 42rem; }
.eyebrow { font-size: 11px; letter-spacing: .18em; text-transform: uppercase; color: var(--muted); margin: 0 0 .35rem; }
.kpi-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(7.5rem, 1fr)); gap: .75rem; margin-top: 1.5rem; }
.kpi {
  background: var(--card); border-radius: var(--radius); padding: 1rem; box-shadow: var(--shadow);
  text-decoration: none; transition: box-shadow .15s;
}
.kpi:hover { box-shadow: var(--shadow-hover); }
.kpi .k { font-size: 11px; letter-spacing: .14em; text-transform: uppercase; color: var(--muted); }
.kpi .v { font-family: Fraunces, serif; font-size: 2rem; margin-top: .4rem; tabular-nums: true; }
.panel { background: var(--card); border-radius: var(--radius); padding: 1.25rem; box-shadow: var(--shadow); }
.panel + .panel, .mt { margin-top: 1rem; }
.bar-row { display: grid; grid-template-columns: 9.5rem 1fr 2rem; gap: .75rem; align-items: center; text-decoration: none; border-radius: 8px; padding: .15rem 0; }
.bar-row:hover { background: var(--secondary); }
.bar-track { height: .5rem; background: var(--secondary); border-radius: 999px; overflow: hidden; }
.bar-fill { display: block; height: 100%; background: var(--accent); border-radius: 999px; }
.muted { color: var(--muted); }
.grid-2 { display: grid; gap: 1rem; }
@media (min-width: 900px) { .grid-2 { grid-template-columns: 1fr 1fr; } }
.cat-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(7rem, 1fr)); gap: .75rem; }
.cat-cell { background: var(--secondary); border-radius: 10px; padding: .75rem; }
.cat-cell .v { font-family: Fraunces, serif; font-size: 1.5rem; }
.owner-list { list-style: none; margin: .75rem 0 0; padding: 0; }
.owner-list li { display: flex; justify-content: space-between; align-items: center; padding: .55rem 0; border-bottom: 1px solid var(--border); }
.owner-list li:last-child { border-bottom: 0; }
.gap-list { list-style: none; margin: .75rem 0 0; padding: 0; }
.gap-list li { display: flex; justify-content: space-between; gap: .75rem; padding: .35rem 0; font-size: .9rem; }
.filters { margin-top: 1.5rem; display: flex; flex-direction: column; gap: 1rem; }
.filter-row { display: flex; flex-wrap: wrap; gap: .5rem; align-items: center; }
.filter-label { font-size: 11px; letter-spacing: .14em; text-transform: uppercase; color: var(--muted); min-width: 7.5rem; }
.chip {
  border: 0; background: var(--card); box-shadow: var(--shadow); border-radius: 999px;
  padding: .45rem .85rem; font: inherit; font-size: .8rem; color: var(--muted); cursor: pointer;
}
.chip:hover { color: var(--fg); }
.chip.active { background: var(--accent); color: #fff; box-shadow: none; }
.selects { display: flex; flex-wrap: wrap; gap: .75rem; }
.selects label {
  display: flex; flex-direction: column; gap: .35rem;
  font-size: 11px; letter-spacing: .14em; text-transform: uppercase; color: var(--muted);
}
.selects select {
  font: inherit; font-size: .875rem; text-transform: none; letter-spacing: 0;
  padding: .55rem .75rem; border-radius: 8px; border: 1px solid var(--border);
  background: var(--card); color: var(--fg); min-width: 10rem;
}
.count { margin-top: 1.25rem; color: var(--muted); font-size: .9rem; }
.cards { margin-top: 1.25rem; display: grid; grid-template-columns: repeat(auto-fill, minmax(17rem, 1fr)); gap: .9rem; }
.card {
  background: var(--card); border-radius: var(--radius); box-shadow: var(--shadow);
  padding: 1rem 1.05rem; display: flex; flex-direction: column; gap: .55rem;
  text-decoration: none; transition: box-shadow .15s;
}
.card:hover { box-shadow: var(--shadow-hover); }
.card-top { display: flex; justify-content: space-between; gap: .5rem; align-items: flex-start; }
.card-title { font-weight: 600; font-size: .95rem; }
.badge {
  display: inline-flex; align-items: center; border-radius: 999px;
  padding: .2rem .55rem; font-size: .7rem; font-weight: 600; white-space: nowrap;
}
.badge-production { color: #166534; background: #dcfce7; }
.badge-pre-production { color: #9a3412; background: #ffedd5; }
.badge-development { color: #1e40af; background: #dbeafe; }
.badge-pre-development { color: #6b21a8; background: #f3e8ff; }
.badge-pipeline { color: #3f3f46; background: #f4f4f5; }
.badge-terminated { color: #991b1b; background: #fee2e2; }
.meta { font-size: .8rem; color: var(--muted); }
.heads { display: flex; flex-wrap: wrap; gap: .35rem; }
.pill { background: var(--secondary); color: var(--fg); border-radius: 6px; padding: .15rem .45rem; font-size: .72rem; }
.owners { display: flex; gap: .5rem; flex-wrap: wrap; margin-top: auto; padding-top: .35rem; }
.owner {
  display: inline-flex; align-items: center; gap: .4rem; font-size: .78rem;
  text-decoration: none; color: var(--fg);
}
.avatar {
  width: 1.5rem; height: 1.5rem; border-radius: 999px; background: var(--accent); color: #fff;
  display: grid; place-items: center; font-size: 9px; font-weight: 700;
}
.avatar.muted { background: #94a3b8; }
.owner .role { font-size: 9px; letter-spacing: .12em; text-transform: uppercase; color: var(--muted); }
.table-wrap {
  margin-top: 2rem; overflow: auto; background: var(--card);
  border-radius: var(--radius); box-shadow: var(--shadow);
}
table { width: 100%; border-collapse: collapse; font-size: .85rem; }
th, td { padding: .7rem .85rem; text-align: left; border-bottom: 1px solid var(--border); vertical-align: top; }
th {
  font-size: 11px; letter-spacing: .12em; text-transform: uppercase; color: var(--muted);
  font-weight: 600; background: #f7fafc; position: sticky; top: 0;
}
tr:last-child td { border-bottom: 0; }
.matrix-wrap { overflow: auto; margin-top: 1rem; }
.matrix { font-size: .8rem; }
.matrix th, .matrix td { text-align: center; white-space: nowrap; }
.matrix td:first-child, .matrix th:first-child { text-align: left; position: sticky; left: 0; background: var(--card); }
.mark { color: var(--accent); font-weight: 700; }
.section { margin-top: 2.5rem; }
.glossary dt { font-weight: 600; margin-top: 1rem; }
.glossary dd { margin: .35rem 0 0; color: var(--muted); max-width: 40rem; }
.footer { margin-top: 2rem; color: var(--muted); font-size: .8rem; }
.detail-grid { display: grid; gap: 1rem; margin-top: 1.5rem; }
@media (min-width: 800px) { .detail-grid { grid-template-columns: 2fr 1fr; } }
.dl { display: grid; grid-template-columns: 9rem 1fr; gap: .5rem .75rem; font-size: .9rem; }
.dl dt { color: var(--muted); }
.owner-card {
  display: block; background: var(--card); border-radius: var(--radius); padding: 1.1rem;
  box-shadow: var(--shadow); text-decoration: none; transition: box-shadow .15s;
}
.owner-card:hover { box-shadow: var(--shadow-hover); }
.owner-card .big { font-family: Fraunces, serif; font-size: 1.75rem; }
.search-box {
  margin-left: auto; display: flex; align-items: center; gap: .5rem;
  background: var(--card); box-shadow: var(--shadow); border-radius: 8px;
  padding: .55rem .85rem; min-width: 14rem; max-width: 22rem; width: 100%;
}
.search-box input { border: 0; outline: 0; background: transparent; width: 100%; font: inherit; font-size: .875rem; color: var(--fg); }
.empty { margin-top: 2rem; padding: 2rem; text-align: center; color: var(--muted); background: var(--card); border-radius: var(--radius); box-shadow: var(--shadow); }
"""


def shell(page: str, title: str, body: str, products: list[dict], updated: str, extra_head: str = "", extra_js: str = "") -> str:
    live = sum(1 for p in products if p["stage"] == "Production")
    delivery = sum(1 for p in products if p["stage"] in ("Pre-Production", "Development"))
    pipe = sum(1 for p in products if p["stage"] == "Pipeline")
    nav = [
        ("index.html", "Overview", "overview"),
        ("products.html", "Products", "products"),
        ("platforms.html", "Platforms", "platforms"),
        ("roadmap.html", "Roadmap", "roadmap"),
        ("owners.html", "Owners", "owners"),
        ("glossary.html", "Legend", "glossary"),
    ]
    nav_html = "".join(
        f'<a class="{"active" if page == key else ""}" href="{href}">{label}</a>'
        for href, label, key in nav
    )
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>{esc(title)} · UGT Cloud Portfolio</title>
<meta name="description" content="UGT Cloud product portfolio — products, platforms, lifecycle stages, and CPO / TPO ownership."/>
<link rel="preconnect" href="https://fonts.googleapis.com"/>
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin/>
<link href="https://fonts.googleapis.com/css2?family=Figtree:wght@400;500;600;700&family=Fraunces:opsz,wght@9..144,500;9..144,600&display=swap" rel="stylesheet"/>
<link rel="stylesheet" href="assets/site.css"/>
{extra_head}
</head>
<body>
<div class="app">
  <aside class="rail">
    <a class="brand" href="index.html">
      <div class="logo">U</div>
      <div>
        <div class="font-display" style="font-size:1.15rem">UGT Cloud</div>
        <div class="sub">Portfolio</div>
      </div>
    </a>
    <nav class="side">{nav_html}</nav>
    <div class="snap">
      <div class="label">Snapshot</div>
      <div class="num">{len(products)}</div>
      <div class="hint">products in the 2026 register</div>
      <div class="updated">Data last updated<br/>{esc(updated)}</div>
    </div>
  </aside>
  <div>
    <header class="top">
      <div class="stats">{live} live · {delivery} in delivery · {pipe} pipeline</div>
    </header>
    <div class="content">
{body}
      <p class="footer">Built from UGT_Product_Portfolio_2026.xlsx · sheet “UGT Product Portfolio ALL”.</p>
    </div>
  </div>
</div>
{extra_js}
</body>
</html>
"""


def bar_list(items: list[tuple[str, int, str]], max_n: int | None = None) -> str:
    m = max_n or max((n for _, n, _ in items), default=1) or 1
    bits = []
    for label, n, href in items:
        pct = 100.0 * n / m
        bits.append(
            f'<a class="bar-row" href="{href}">'
            f'<span class="muted" style="font-size:.875rem;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">{esc(label)}</span>'
            f'<span class="bar-track"><span class="bar-fill" style="width:{pct:.1f}%"></span></span>'
            f'<span style="text-align:right;font-size:.875rem">{n}</span></a>'
        )
    return "\n".join(bits)


def build_overview(products: list[dict], updated: str) -> str:
    by_stage = Counter(p["stage"] for p in products)
    by_due = Counter(p["due"] for p in products if p["due"])
    by_cat = Counter(p["category"] for p in products if p["category"])
    by_head = Counter()
    for p in products:
        for h in p["headProducts"]:
            if h not in ("Unplaced", "Cross-portfolio"):
                by_head[h] += 1
            elif h == "Cross-portfolio":
                by_head[h] += 1
    # Prefer platform order
    head_order = [h for h in PLATFORM_COLS if h in by_head] + [
        h for h in by_head if h not in PLATFORM_COLS
    ]
    cpo_c = Counter(p["cpo"] for p in products if p["cpo"])
    tpo_c = Counter(p["tpo"] for p in products if p["tpo"])
    no_cpo = [p for p in products if not p["cpo"]]
    no_tpo = [p for p in products if not p["tpo"]]
    scheduled = sum(1 for p in products if p["due"])
    delivery = by_stage.get("Pre-Production", 0) + by_stage.get("Development", 0)
    owners_n = len(set(list(cpo_c) + list(tpo_c)))

    kpis = [
        ("Total products", len(products), "products.html"),
        ("Production", by_stage.get("Production", 0), "products.html?stage=Production"),
        ("In delivery", delivery, "roadmap.html"),
        ("Scheduled", scheduled, "roadmap.html"),
        ("Pipeline", by_stage.get("Pipeline", 0), "products.html?stage=Pipeline"),
        ("Pre-development", by_stage.get("Pre-Development", 0), "products.html?stage=Pre-Development"),
        ("Terminated", by_stage.get("Terminated", 0), "products.html?stage=Terminated"),
    ]
    kpi_html = "".join(
        f'<a class="kpi" href="{href}"><div class="k">{esc(k)}</div><div class="v">{v}</div></a>'
        for k, v, href in kpis
    )
    stage_bars = bar_list([(s, by_stage.get(s, 0), f"products.html?stage={s}") for s in STAGES], by_stage.get("Production", 1))
    due_items = sorted(by_due.items())
    due_bars = bar_list([(d, n, f"products.html?due={d.replace(' ', '+')}") for d, n in due_items], max(by_due.values(), default=1))
    head_bars = bar_list([(h, by_head[h], f"products.html?platform={h.replace(' ', '+')}") for h in head_order], max(by_head.values(), default=1))
    cats = "".join(
        f'<div class="cat-cell"><div class="v">{n}</div><div class="muted" style="font-size:.85rem;margin-top:.25rem">{esc(c)}</div></div>'
        for c, n in by_cat.most_common()
    )

    def owner_side(counter, role_path):
        items = []
        for code, n in counter.most_common():
            items.append(
                f'<li><a class="owner" href="owner-{esc(code)}.html">'
                f'<span class="avatar">{esc(code)}</span>'
                f'<span><div style="font-weight:500">{esc(code)}</div>'
                f'<div class="role">{role_path}</div></span></a>'
                f'<span class="muted">{n}</span></li>'
            )
        return "<ul class='owner-list'>" + "".join(items) + "</ul>"

    gaps_cpo = "".join(
        f'<li><a href="product-{p["n"]}.html">{esc(p["product"])}</a>'
        f'<span class="badge {badge_class(p["stage"])}">{esc(p["stage"])}</span></li>'
        for p in no_cpo
    )
    gaps_tpo = "".join(
        f'<li><a href="product-{p["n"]}.html">{esc(p["product"])}</a>'
        f'<span class="badge {badge_class(p["stage"])}">{esc(p["stage"])}</span></li>'
        for p in no_tpo
    )

    body = f"""
      <p class="eyebrow">2026 register</p>
      <h1>Product portfolio at a glance</h1>
      <p class="lede">Every UGT Cloud offer, the platform it sits on, and the people who own it commercially and technically.</p>
      <div class="kpi-grid">{kpi_html}</div>
      <div class="grid-2 mt" style="margin-top:1.5rem">
        <div class="panel">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:1rem">
            <h2>Lifecycle</h2>
            <a class="muted" href="roadmap.html">↗</a>
          </div>
          {stage_bars}
        </div>
        <div class="panel">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:1rem">
            <h2>Scheduled</h2>
            <a class="muted" href="roadmap.html">↗</a>
          </div>
          {due_bars}
          <p class="muted" style="margin-top:1rem;font-size:.875rem">{scheduled} products have a due date.</p>
        </div>
      </div>
      <div class="panel mt">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:1rem">
          <h2>Head products</h2>
          <a class="muted" href="platforms.html">↗</a>
        </div>
        {head_bars}
      </div>
      <div class="panel mt">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:1rem">
          <h2>Categories</h2>
          <a class="muted" href="products.html">↗</a>
        </div>
        <div class="cat-grid">{cats}</div>
      </div>
      <div class="section">
        <h2 style="font-size:1.85rem">Ownership</h2>
        <p class="muted" style="margin:.35rem 0 1rem;font-size:.9rem">CPO is commercial product owner. TPO is technical product owner.</p>
        <div class="grid-2">
          <div class="panel"><div class="eyebrow">Commercial</div>{owner_side(cpo_c, "CPO")}</div>
          <div class="panel"><div class="eyebrow">Technical</div>{owner_side(tpo_c, "TPO")}</div>
        </div>
      </div>
      <div class="panel section">
        <h2>Ownership gaps</h2>
        <p class="muted" style="margin:.35rem 0 1rem;font-size:.9rem">Lines still marked N/A in the register — useful for staffing conversations.</p>
        <div class="grid-2">
          <div>
            <h3>No commercial owner <span class="muted">({len(no_cpo)})</span></h3>
            <ul class="gap-list">{gaps_cpo}</ul>
          </div>
          <div>
            <h3>No technical owner <span class="muted">({len(no_tpo)})</span></h3>
            <ul class="gap-list">{gaps_tpo}</ul>
          </div>
        </div>
      </div>
      <p class="muted" style="margin-top:2rem;font-size:.9rem">{scheduled} products have a quarter on the roadmap. {owners_n} owners across the portfolio.</p>
"""
    return shell("overview", "UGT Cloud Portfolio", body, products, updated)


PRODUCTS_JS = r"""
<script>
const PRODUCTS = __PRODUCTS__;
const STAGES = __STAGES__;
const state = { stage: "", due: "", head: "", category: "", cpo: "", tpo: "", q: "" };

function badgeClass(stage) {
  return "badge-" + String(stage || "").toLowerCase().replace(/\s+/g, "-");
}
function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
}
function ownerChip(code, role) {
  if (!code) return `<span class="owner"><span class="avatar muted">—</span><span><div>Unassigned</div><div class="role">${role}</div></span></span>`;
  return `<a class="owner" href="owner-${esc(code)}.html"><span class="avatar">${esc(code)}</span><span><div>${esc(code)}</div><div class="role">${role}</div></span></a>`;
}
function matches(p) {
  if (state.stage && p.stage !== state.stage) return false;
  if (state.due === "Unscheduled") { if (p.due) return false; }
  else if (state.due && p.due !== state.due) return false;
  if (state.head && !(p.headProducts || []).includes(state.head)) return false;
  if (state.category && p.category !== state.category) return false;
  if (state.cpo === "__unassigned__") { if (p.cpo) return false; }
  else if (state.cpo && p.cpo !== state.cpo) return false;
  if (state.tpo === "__unassigned__") { if (p.tpo) return false; }
  else if (state.tpo && p.tpo !== state.tpo) return false;
  if (state.q) {
    const hay = [p.product, p.category, p.stage, p.due || "", p.technology || "", p.cpo || "", p.tpo || "", ...(p.headProducts || [])].join(" ").toLowerCase();
    if (!hay.includes(state.q)) return false;
  }
  return true;
}
function renderChips(id, values, key, allLabel, extra) {
  const el = document.getElementById(id);
  const label = el.querySelector(".filter-label");
  let html = "";
  html += `<button type="button" class="chip ${state[key]===""?"active":""}" data-key="${key}" data-val="">${allLabel}</button>`;
  for (const v of values) html += `<button type="button" class="chip ${state[key]===v?"active":""}" data-key="${key}" data-val="${esc(v)}">${esc(v)}</button>`;
  if (extra) for (const [lab, val] of extra) html += `<button type="button" class="chip ${state[key]===val?"active":""}" data-key="${key}" data-val="${esc(val)}">${esc(lab)}</button>`;
  el.innerHTML = "";
  if (label) el.appendChild(label);
  else { const s=document.createElement("span"); s.className="filter-label"; s.textContent=key==="stage"?"Stage":"Due date"; el.appendChild(s); }
  el.insertAdjacentHTML("beforeend", html);
}
function render() {
  const rows = PRODUCTS.filter(matches);
  document.getElementById("view-count").textContent = `${rows.length} line${rows.length===1?"":"s"} in view.`;
  const cards = document.getElementById("cards");
  const tbody = document.getElementById("tbody");
  const empty = document.getElementById("empty");
  if (!rows.length) { cards.innerHTML=""; tbody.innerHTML=""; empty.hidden=false; return; }
  empty.hidden = true;
  cards.innerHTML = rows.map(p => `
    <a class="card" href="product-${p.n}.html">
      <div class="card-top"><div class="card-title">${esc(p.product)}</div><span class="badge ${badgeClass(p.stage)}">${esc(p.stage)}</span></div>
      <div class="meta">${esc(p.category || "—")}</div>
      <div class="heads">${(p.headProducts||[]).map(h=>`<span class="pill">${esc(h)}</span>`).join("")}</div>
      <div class="meta">${esc(p.due || "Unscheduled")}${p.technology ? " · " + esc(p.technology) : ""}</div>
      <div class="owners">${ownerChip(p.cpo,"CPO")}${ownerChip(p.tpo,"TPO")}</div>
    </a>`).join("");
  tbody.innerHTML = rows.map(p => `
    <tr>
      <td><a href="product-${p.n}.html"><strong>${esc(p.product)}</strong></a></td>
      <td>${esc(p.category || "—")}</td>
      <td>${esc((p.headProducts||[]).join(", "))}</td>
      <td><span class="badge ${badgeClass(p.stage)}">${esc(p.stage)}</span></td>
      <td>${esc(p.due || "Unscheduled")}</td>
      <td>${esc(p.cpo || "Unassigned")}</td>
      <td>${esc(p.tpo || "Unassigned")}</td>
      <td>${esc(p.technology || "—")}</td>
    </tr>`).join("");
}
function bind() {
  const dues = [...new Set(PRODUCTS.map(p => p.due).filter(Boolean))].sort();
  const heads = [...new Set(PRODUCTS.flatMap(p => p.headProducts).filter(h => h && h !== "Unplaced"))].sort((a,b)=>a.localeCompare(b));
  const cats = [...new Set(PRODUCTS.map(p => p.category).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
  const cpos = [...new Set(PRODUCTS.map(p => p.cpo).filter(Boolean))].sort();
  const tpos = [...new Set(PRODUCTS.map(p => p.tpo).filter(Boolean))].sort();
  function fill(id, values, withUnassigned) {
    const sel = document.getElementById(id);
    const keep = withUnassigned ? 2 : 1;
    const opts = [...sel.querySelectorAll("option")].slice(0, keep);
    sel.innerHTML = ""; opts.forEach(o => sel.appendChild(o));
    for (const v of values) { const o=document.createElement("option"); o.value=v; o.textContent=v; sel.appendChild(o); }
  }
  fill("f-head", heads, false); fill("f-cat", cats, false); fill("f-cpo", cpos, true); fill("f-tpo", tpos, true);
  const params = new URLSearchParams(location.search);
  if (params.get("due")) state.due = params.get("due");
  if (params.get("stage")) state.stage = params.get("stage");
  if (params.get("platform")) state.head = params.get("platform");
  if (params.get("category")) state.category = params.get("category");
  if (params.get("cpo")) state.cpo = params.get("cpo");
  if (params.get("tpo")) state.tpo = params.get("tpo");
  renderChips("stage-filters", STAGES, "stage", "All");
  renderChips("due-filters", dues, "due", "All", [["Unscheduled","Unscheduled"]]);
  if (state.head) document.getElementById("f-head").value = state.head;
  if (state.category) document.getElementById("f-cat").value = state.category;
  if (state.cpo) document.getElementById("f-cpo").value = state.cpo;
  if (state.tpo) document.getElementById("f-tpo").value = state.tpo;
  document.getElementById("stage-filters").addEventListener("click", e => {
    const btn = e.target.closest(".chip"); if (!btn) return;
    state.stage = btn.dataset.val; renderChips("stage-filters", STAGES, "stage", "All"); render();
  });
  document.getElementById("due-filters").addEventListener("click", e => {
    const btn = e.target.closest(".chip"); if (!btn) return;
    state.due = btn.dataset.val; renderChips("due-filters", dues, "due", "All", [["Unscheduled","Unscheduled"]]); render();
  });
  document.getElementById("f-head").addEventListener("change", e => { state.head = e.target.value; render(); });
  document.getElementById("f-cat").addEventListener("change", e => { state.category = e.target.value; render(); });
  document.getElementById("f-cpo").addEventListener("change", e => { state.cpo = e.target.value; render(); });
  document.getElementById("f-tpo").addEventListener("change", e => { state.tpo = e.target.value; render(); });
  document.getElementById("q").addEventListener("input", e => { state.q = e.target.value.trim().toLowerCase(); render(); });
  render();
}
bind();
</script>
"""


def build_products(products: list[dict], updated: str) -> str:
    body = """
      <p class="eyebrow">2026 register</p>
      <h1>Products</h1>
      <p class="lede">Filter by stage, quarter, platform, category or owner.</p>
      <div class="filters">
        <div class="filter-row" id="stage-filters"><span class="filter-label">Stage</span></div>
        <div class="filter-row" id="due-filters"><span class="filter-label">Due date</span></div>
        <div class="selects">
          <label>Head product<select id="f-head"><option value="">All</option></select></label>
          <label>Category<select id="f-cat"><option value="">All</option></select></label>
          <label>Commercial owner<select id="f-cpo"><option value="">All</option><option value="__unassigned__">Unassigned</option></select></label>
          <label>Technical owner<select id="f-tpo"><option value="">All</option><option value="__unassigned__">Unassigned</option></select></label>
        </div>
      </div>
      <p class="count" id="view-count"></p>
      <div class="cards" id="cards"></div>
      <div class="table-wrap"><table>
        <thead><tr><th>Product</th><th>Category</th><th>Head product</th><th>Stage</th><th>Due</th><th>CPO</th><th>TPO</th><th>Technology</th></tr></thead>
        <tbody id="tbody"></tbody>
      </table></div>
      <div class="empty" id="empty" hidden>No products match these filters.</div>
"""
    # inject search into header via extra - simpler to add search in body top
    body = """
      <div style="display:flex;flex-wrap:wrap;gap:1rem;align-items:flex-end;justify-content:space-between">
        <div>
          <p class="eyebrow">2026 register</p>
          <h1>Products</h1>
          <p class="lede">Filter by stage, quarter, platform, category or owner.</p>
        </div>
        <label class="search-box"><span aria-hidden="true">⌕</span>
          <input id="q" type="search" placeholder="Search products, owners, platforms" autocomplete="off"/>
        </label>
      </div>
""" + body.split("</p>", 1)[-1] if False else body
    # rewrite cleanly
    body = """
      <div style="display:flex;flex-wrap:wrap;gap:1rem;align-items:flex-end;justify-content:space-between">
        <div>
          <p class="eyebrow">2026 register</p>
          <h1>Products</h1>
          <p class="lede">Filter by stage, quarter, platform, category or owner.</p>
        </div>
        <label class="search-box"><span aria-hidden="true">⌕</span>
          <input id="q" type="search" placeholder="Search products, owners, platforms" autocomplete="off"/>
        </label>
      </div>
      <div class="filters">
        <div class="filter-row" id="stage-filters"><span class="filter-label">Stage</span></div>
        <div class="filter-row" id="due-filters"><span class="filter-label">Due date</span></div>
        <div class="selects">
          <label>Head product<select id="f-head"><option value="">All</option></select></label>
          <label>Category<select id="f-cat"><option value="">All</option></select></label>
          <label>Commercial owner<select id="f-cpo"><option value="">All</option><option value="__unassigned__">Unassigned</option></select></label>
          <label>Technical owner<select id="f-tpo"><option value="">All</option><option value="__unassigned__">Unassigned</option></select></label>
        </div>
      </div>
      <p class="count" id="view-count"></p>
      <div class="cards" id="cards"></div>
      <div class="table-wrap"><table>
        <thead><tr><th>Product</th><th>Category</th><th>Head product</th><th>Stage</th><th>Due</th><th>CPO</th><th>TPO</th><th>Technology</th></tr></thead>
        <tbody id="tbody"></tbody>
      </table></div>
      <div class="empty" id="empty" hidden>No products match these filters.</div>
"""
    js = PRODUCTS_JS.replace("__PRODUCTS__", json.dumps(products, ensure_ascii=False)).replace(
        "__STAGES__", json.dumps(STAGES)
    )
    return shell("products", "Products", body, products, updated, extra_js=js)


def build_platforms(products: list[dict], updated: str) -> str:
    counts = Counter()
    for p in products:
        for h in p["headProducts"]:
            if h in PLATFORM_COLS:
                counts[h] += 1
    chips = ['<a class="chip" href="products.html">All</a>'] + [
        f'<a class="chip" href="products.html?platform={h.replace(" ", "+")}">{esc(h)} · {counts.get(h, 0)}</a>'
        for h in PLATFORM_COLS
    ]
    rows = []
    for p in products:
        marks = "".join(
            f'<td class="mark">{"●" if h in p["headProducts"] else ""}</td>' for h in PLATFORM_COLS
        )
        rows.append(
            f'<tr><td><a href="product-{p["n"]}.html">{esc(p["product"])}</a> '
            f'<span class="muted">{esc(p["category"])}</span></td>{marks}'
            f'<td><span class="badge {badge_class(p["stage"])}">{esc(p["stage"])}</span></td></tr>'
        )
    head_cells = "".join(f"<th>{esc(h.split()[0] if h!='Bare Metal' else 'Bare Metal')}</th>" for h in PLATFORM_COLS)
    # cleaner headers
    head_cells = "".join(
        f"<th>{esc({'Universal Cloud':'Universal','Public Cloud':'Public','Private Cloud':'Private','Bare Metal':'Bare Metal','SMB':'SMB','Standalone':'Standalone'}[h])}</th>"
        for h in PLATFORM_COLS
    )
    body = f"""
      <p class="eyebrow">2026 register</p>
      <h1>Product × platform</h1>
      <p class="lede">Which offers sit on each head product. A filled mark means the line is associated with that platform.</p>
      <div class="filter-row" style="margin-top:1.25rem"><span class="filter-label">Head product</span>{''.join(chips)}</div>
      <div class="panel mt">
        <ul style="list-style:none;margin:0;padding:0;display:flex;flex-wrap:wrap;gap:1rem">
          {''.join(f'<li><strong>{counts.get(h,0)}</strong> <span class="muted">{esc(h)}</span></li>' for h in PLATFORM_COLS)}
        </ul>
      </div>
      <div class="table-wrap matrix-wrap"><table class="matrix">
        <thead><tr><th>Product</th>{head_cells}<th>Stage</th></tr></thead>
        <tbody>{''.join(rows)}</tbody>
      </table></div>
"""
    return shell("platforms", "Platforms", body, products, updated)


def build_roadmap(products: list[dict], updated: str) -> str:
    dues = sorted({p["due"] for p in products if p["due"]})
    sections = []
    for due in dues:
        group = [p for p in products if p["due"] == due]
        sections.append(
            f'<section class="section" id="due-{esc(due).replace(" ","-")}">'
            f"<h2>{esc(due)}</h2>"
            f'<p class="muted">Target from the register. <strong>{len(group)}</strong></p>'
            f'<div class="cards">{"".join(product_card_html(p) for p in group)}</div>'
            f'<div class="table-wrap"><table><thead><tr><th>Product</th><th>Category</th><th>Head product</th><th>Stage</th><th>Due</th><th>CPO</th><th>TPO</th><th>Technology</th></tr></thead>'
            f"<tbody>{product_table_rows(group)}</tbody></table></div></section>"
        )
    # delivery stages first then production/pipeline/terminated like Grok
    stage_order = ["Pre-Production", "Development", "Pre-Development", "Pipeline", "Production", "Terminated"]
    blurbs = {
        "Pre-Production": "Final preparation before going live.",
        "Development": "Actively being built.",
        "Pre-Development": "Discovery and preparation before build.",
        "Pipeline": "Planned backlog — not yet in delivery.",
        "Production": "Live and operational.",
        "Terminated": "Removed from the active portfolio.",
    }
    for stage in stage_order:
        group = [p for p in products if p["stage"] == stage]
        if not group:
            continue
        sections.append(
            f'<section class="section" id="stage-{stage.lower().replace(" ","-")}">'
            f"<h2>{esc(stage)}</h2>"
            f'<p class="muted">{esc(blurbs.get(stage, ""))} <strong>{len(group)}</strong></p>'
            f'<div class="cards">{"".join(product_card_html(p) for p in group)}</div>'
            f'<div class="table-wrap"><table><thead><tr><th>Product</th><th>Category</th><th>Head product</th><th>Stage</th><th>Due</th><th>CPO</th><th>TPO</th><th>Technology</th></tr></thead>'
            f"<tbody>{product_table_rows(group)}</tbody></table></div></section>"
        )
    body = f"""
      <p class="eyebrow">2026 register</p>
      <h1>Roadmap</h1>
      <p class="lede">Year-quarter targets first, then the same register by lifecycle stage.</p>
      <p class="muted" style="margin-top:1rem">By quarter</p>
      {''.join(sections)}
"""
    return shell("roadmap", "Roadmap", body, products, updated)


def build_owners_index(products: list[dict], updated: str) -> str:
    cpos = sorted({p["cpo"] for p in products if p["cpo"]})
    tpos = sorted({p["tpo"] for p in products if p["tpo"]})
    no_cpo = sum(1 for p in products if not p["cpo"])
    no_tpo = sum(1 for p in products if not p["tpo"])

    def cards(codes, kind):
        out = []
        for code in codes:
            title, blurb = OWNER_BLURBS.get(code, ("Product Owner", ""))
            as_cpo = sum(1 for p in products if p["cpo"] == code)
            as_tpo = sum(1 for p in products if p["tpo"] == code)
            owned = [p for p in products if p["cpo"] == code or p["tpo"] == code]
            stages = Counter(p["stage"] for p in owned)
            stage_pills = " ".join(
                f'<span class="badge {badge_class(s)}">{esc(s)}</span>' for s in STAGES if stages.get(s)
            )
            out.append(
                f'<a class="owner-card" href="owner-{esc(code)}.html">'
                f'<div style="display:flex;gap:.75rem;align-items:center">'
                f'<span class="avatar" style="width:2.5rem;height:2.5rem;font-size:.75rem">{esc(code)}</span>'
                f'<div><div class="big">{esc(code)}</div><div class="muted" style="font-size:.8rem">{esc(title)}</div></div></div>'
                f'<p class="muted" style="margin:.75rem 0;font-size:.9rem">{esc(blurb)}</p>'
                f'<div style="font-family:Fraunces,serif;font-size:2rem">{as_cpo if kind=="cpo" else as_tpo}</div>'
                f'<div class="muted" style="font-size:.8rem">{as_cpo} as CPO · {as_tpo} as TPO</div>'
                f'<div style="margin-top:.75rem;display:flex;flex-wrap:wrap;gap:.35rem">{stage_pills}</div>'
                f"</a>"
            )
        return "".join(out)

    body = f"""
      <p class="eyebrow">2026 register</p>
      <h1>Product owners</h1>
      <p class="lede">CPO — commercial product owner. TPO — technical product owner. Open anyone to see the lines they own.</p>
      <h2 class="section">Commercial</h2>
      <div class="cards" style="margin-top:1rem">{cards(cpos, "cpo")}</div>
      <h2 class="section">Technical</h2>
      <div class="cards" style="margin-top:1rem">{cards(tpos, "tpo")}</div>
      <div class="panel section">
        <h2>Unassigned</h2>
        <p class="muted">{no_cpo} lines without a CPO. {no_tpo} without a TPO.</p>
        <p style="margin-top:.75rem">{' '.join(f'<a href="product-{p["n"]}.html">{esc(p["product"])}</a>' for p in products if not p["cpo"] or not p["tpo"])}</p>
      </div>
"""
    return shell("owners", "Owners", body, products, updated)


def build_owner_page(code: str, products: list[dict], updated: str) -> str:
    title, blurb = OWNER_BLURBS.get(code, ("Product Owner", ""))
    owned = [p for p in products if p["cpo"] == code or p["tpo"] == code]
    as_cpo = [p for p in products if p["cpo"] == code]
    as_tpo = [p for p in products if p["tpo"] == code]
    stages = Counter(p["stage"] for p in owned)
    cats = Counter(p["category"] for p in owned if p["category"])
    body = f"""
      <p class="eyebrow"><a href="owners.html" class="muted">Owners</a> / {esc(code)}</p>
      <h1>{esc(code)}</h1>
      <p class="lede">{esc(blurb)}. {len(owned)} product lines in this register.</p>
      <div class="grid-2" style="margin-top:1.5rem">
        <div class="panel">
          <h2>Lifecycle mix</h2>
          {bar_list([(s, stages.get(s, 0), f"products.html?stage={s}") for s in STAGES if stages.get(s)], max(stages.values(), default=1))}
        </div>
        <div class="panel">
          <h2>Categories</h2>
          <div class="cat-grid" style="margin-top:1rem">
            {''.join(f'<div class="cat-cell"><div class="v">{n}</div><div class="muted" style="font-size:.85rem">{esc(c)}</div></div>' for c,n in cats.most_common())}
          </div>
        </div>
      </div>
"""
    if as_cpo:
        body += f"""
      <section class="section">
        <h2>As commercial owner</h2>
        <p class="muted">Accountable for the offer, packaging and go-to-market. · {len(as_cpo)} as CPO · {len(as_tpo)} as TPO</p>
        <div class="cards">{''.join(product_card_html(p) for p in as_cpo)}</div>
        <div class="table-wrap"><table><thead><tr><th>Product</th><th>Category</th><th>Head product</th><th>Stage</th><th>Due</th><th>CPO</th><th>TPO</th><th>Technology</th></tr></thead>
        <tbody>{product_table_rows(as_cpo)}</tbody></table></div>
      </section>
"""
    if as_tpo:
        body += f"""
      <section class="section">
        <h2>As technical owner</h2>
        <p class="muted">Accountable for technical delivery, architecture and operational readiness.</p>
        <div class="cards">{''.join(product_card_html(p) for p in as_tpo)}</div>
        <div class="table-wrap"><table><thead><tr><th>Product</th><th>Category</th><th>Head product</th><th>Stage</th><th>Due</th><th>CPO</th><th>TPO</th><th>Technology</th></tr></thead>
        <tbody>{product_table_rows(as_tpo)}</tbody></table></div>
      </section>
"""
    return shell("owners", code, body, products, updated)


def build_product_page(p: dict, products: list[dict], updated: str) -> str:
    body = f"""
      <p class="eyebrow"><a class="muted" href="products.html">Products</a> / {esc(p['product'])}</p>
      <div style="display:flex;flex-wrap:wrap;gap:1rem;align-items:flex-start;justify-content:space-between">
        <div>
          <h1>{esc(p['product'])}</h1>
          <p class="lede">{esc(p['category'] or '—')} · line #{p['n']}</p>
        </div>
        <span class="badge {badge_class(p['stage'])}" style="font-size:.85rem;padding:.4rem .75rem">{esc(p['stage'])}</span>
      </div>
      <div class="detail-grid">
        <div class="panel">
          <h2>Register details</h2>
          <dl class="dl" style="margin-top:1rem">
            <dt>Head product</dt><dd>{esc(', '.join(p['headProducts']))}</dd>
            <dt>Category</dt><dd>{esc(p['category'] or '—')}</dd>
            <dt>Stage</dt><dd>{esc(p['stage'])}</dd>
            <dt>Due</dt><dd>{esc(p['due'] or 'Unscheduled')}</dd>
            <dt>Technology</dt><dd>{esc(p['technology'] or '—')}</dd>
          </dl>
        </div>
        <div class="panel">
          <h2>Ownership</h2>
          <div class="owners" style="margin-top:1rem;flex-direction:column;align-items:flex-start;gap:.75rem">
            {owner_chip_html(p['cpo'], 'CPO')}
            {owner_chip_html(p['tpo'], 'TPO')}
          </div>
        </div>
      </div>
"""
    # related same head
    related = [
        q for q in products
        if q["n"] != p["n"] and set(q["headProducts"]) & set(p["headProducts"])
    ][:6]
    if related:
        body += f"""
      <section class="section">
        <h2>Related on same platform</h2>
        <div class="cards">{''.join(product_card_html(q) for q in related)}</div>
      </section>
"""
    return shell("products", p["product"], body, products, updated)


def build_glossary(products: list[dict], updated: str) -> str:
    items = "".join(f"<dt>{esc(k)}</dt><dd>{esc(v)}</dd>" for k, v in GLOSSARY)
    body = f"""
      <p class="eyebrow">2026 register</p>
      <h1>Legend</h1>
      <p class="lede">Portfolio language used in the register. Head product is the commercial environment; CPO and TPO are people.</p>
      <dl class="glossary">{items}</dl>
"""
    return shell("glossary", "Legend", body, products, updated)


def main() -> None:
    products = load_products(EXCEL)
    updated = data_updated_label()
    OUT_JSON.write_text(json.dumps(products, ensure_ascii=False, indent=2) + "\n")
    SITE.mkdir(parents=True, exist_ok=True)
    (SITE / "assets").mkdir(exist_ok=True)
    (SITE / "assets" / "site.css").write_text(SHARED_CSS)
    (SITE / "index.html").write_text(build_overview(products, updated), encoding="utf-8")
    (SITE / "products.html").write_text(build_products(products, updated), encoding="utf-8")
    (SITE / "platforms.html").write_text(build_platforms(products, updated), encoding="utf-8")
    (SITE / "roadmap.html").write_text(build_roadmap(products, updated), encoding="utf-8")
    (SITE / "owners.html").write_text(build_owners_index(products, updated), encoding="utf-8")
    (SITE / "glossary.html").write_text(build_glossary(products, updated), encoding="utf-8")
    codes = sorted({p["cpo"] for p in products if p["cpo"]} | {p["tpo"] for p in products if p["tpo"]})
    for code in codes:
        (SITE / f"owner-{code}.html").write_text(build_owner_page(code, products, updated), encoding="utf-8")
    for p in products:
        (SITE / f"product-{p['n']}.html").write_text(build_product_page(p, products, updated), encoding="utf-8")
    # also write data.js for sync tooling consumers
    (SITE / "assets" / "data.js").write_text(
        "window.UGT_PRODUCTS = " + json.dumps(products, ensure_ascii=False) + ";\n"
        + "window.UGT_UPDATED = " + json.dumps(updated) + ";\n",
        encoding="utf-8",
    )
    print(f"Built {len(products)} products → {SITE} (updated: {updated})")


if __name__ == "__main__":
    main()
