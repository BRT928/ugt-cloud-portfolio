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
    ncols = ws.max_column
    headers = [clean_text(ws.cell(3, c).value) or f"col{c}" for c in range(1, ncols + 1)]
    products = []
    for r in range(4, ws.max_row + 1):
        row = {headers[c - 1]: ws.cell(r, c).value for c in range(1, ncols + 1)}
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



OUT_TGA_JSON = ROOT / "tga_nni.json"


def name_key(s: str) -> str:
    s = clean_text(s).lower()
    s = s.replace("/", " ")
    s = s.replace("co-location", "colocation").replace("co location", "colocation")
    s = s.replace("marketplace", "marketplace").replace("market place", "marketplace")
    s = re.sub(r"[^a-z0-9ა-ჰ\s]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def load_tga_nni(path: Path) -> list[dict]:
    """Load TGA & NNI sheet. Keep Georgian text; include orphan status rows."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = None
    for cand in ("TGA & NNI", "NNI", "TGA&NNI", "TGA and NNI"):
        if cand in wb.sheetnames:
            ws = wb[cand]
            break
    if ws is None:
        raise KeyError(f"Status sheet not found (tried 'TGA & NNI', 'NNI'); sheets: {wb.sheetnames}")
    raw_headers = [ws.cell(1, c).value for c in range(1, 13)]
    # Normalise header names (leading space on Status Update)
    headers = [clean_text(h) if h else f"col{c}" for c, h in enumerate(raw_headers, 1)]
    # Map known aliases
    rename = {
        "Status Update": "statusUpdate",
        "Head Product": "headProduct",
        "Platform/Technology/Vendor": "technology",
        "Sub Products": "subProducts",
        "Due Date": "due",
        "where can it be used": "whereUsed",
        "Product": "product",
        "Category": "category",
        "Stage": "stage",
        "CPO": "cpo",
        "TPO": "tpo",
        "N": "n",
    }
    rows = []
    for r in range(2, ws.max_row + 1):
        raw = {headers[c - 1]: ws.cell(r, c).value for c in range(1, 13)}
        mapped = {}
        for k, v in raw.items():
            key = rename.get(k, k)
            if key in ("statusUpdate", "whereUsed", "product", "headProduct", "technology", "subProducts", "category", "stage", "due", "cpo", "tpo"):
                # Preserve Georgian; only collapse whitespace / nbsp
                if v is None:
                    mapped[key] = ""
                else:
                    mapped[key] = re.sub(r"[ \t\xa0]+", " ", str(v)).replace("\r\n", "\n").strip()
            elif key == "n":
                try:
                    mapped[key] = int(v) if v is not None and str(v).strip() != "" else None
                except (TypeError, ValueError):
                    mapped[key] = None
            else:
                mapped[key] = v
        has_product = bool(mapped.get("product"))
        has_status = bool(mapped.get("statusUpdate"))
        has_where = bool(mapped.get("whereUsed"))
        # Skip totally empty / where-only spreadsheet debris without product or status
        if not has_product and not has_status:
            continue
        stage = mapped.get("stage") or ""
        mapped["stage"] = STAGE_FIX.get(stage, stage) if stage else ""
        heads = split_heads(mapped.get("headProduct") or "")
        mapped["headProducts"] = heads
        mapped["technology"] = tech_display(mapped.get("technology") or "", mapped.get("subProducts") or "")
        mapped["excelRow"] = r
        mapped["id"] = f"tga-{r}"
        mapped["matchedProductN"] = None
        rows.append(mapped)
    return rows


def attach_tga_notes(products: list[dict], tga_rows: list[dict]) -> None:
    """Match TGA rows to ALL products by name + head (+ tech for duplicates). Mutates both."""
    for p in products:
        p["statusUpdate"] = None
        p["whereUsed"] = None
        p["tgaId"] = None

    def score(t: dict, p: dict) -> int:
        tn, pn = name_key(t.get("product") or ""), name_key(p.get("product") or "")
        if not tn or not pn:
            return 0
        if tn != pn and tn not in pn and pn not in tn:
            return 0
        sc = 10 if tn == pn else 6
        th = {name_key(h) for h in (t.get("headProducts") or [])}
        ph = {name_key(h) for h in p["headProducts"]}
        if th & ph:
            sc += 5
        tt = name_key(t.get("technology") or "")
        pt = name_key(p.get("technology") or "")
        if tt and pt and (tt in pt or pt in tt):
            sc += 3
        return sc

    used_products = set()
    for t in tga_rows:
        if not t.get("product"):
            continue
        scored = sorted(((score(t, p), p) for p in products), key=lambda x: -x[0])
        best_sc, best = scored[0] if scored else (0, None)
        if best_sc <= 0 or best is None:
            continue
        # Prefer unmatched product when ties
        candidates = [p for sc, p in scored if sc == best_sc]
        pick = next((p for p in candidates if p["n"] not in used_products), candidates[0])
        used_products.add(pick["n"])
        t["matchedProductN"] = pick["n"]
        # Prefer non-empty fields; don't overwrite stronger existing note
        if t.get("statusUpdate"):
            pick["statusUpdate"] = t["statusUpdate"]
        if t.get("whereUsed"):
            pick["whereUsed"] = t["whereUsed"]
        pick["tgaId"] = t["id"]


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
    note = ""
    if p.get("statusUpdate"):
        note = f'<div class="note-flag" title="{esc(p["statusUpdate"])}">📝 Status note</div>'
    elif p.get("whereUsed"):
        note = f'<div class="note-flag muted-flag" title="{esc(p["whereUsed"])}">Where: {esc(p["whereUsed"])}</div>'
    return f"""
    <article class="card">
      <a class="card-stretch" href="{base}/product-{p['n']}.html" aria-label="{esc(p['product'])}"></a>
      <div class="card-top">
        <div class="card-title">{esc(p['product'])}</div>
        <span class="badge {badge_class(p['stage'])}">{esc(p['stage'])}</span>
      </div>
      <div class="meta">{esc(p['category'] or '—')}</div>
      <div class="heads">{heads}</div>
      <div class="meta">{esc(due)}{tech}</div>
      {note}
      <div class="owners">{owner_chip_html(p['cpo'], 'CPO', base)}{owner_chip_html(p['tpo'], 'TPO', base)}</div>
    </article>"""


def product_table_rows(products: list[dict], base: str = ".", include_notes: bool = True) -> str:
    rows = []
    for p in products:
        note_cell = ""
        if include_notes:
            bits = []
            if p.get("statusUpdate"):
                bits.append(f'<span class="ka">{esc(p["statusUpdate"])}</span>')
            if p.get("whereUsed"):
                bits.append(f'<span class="muted">{esc(p["whereUsed"])}</span>')
            note_cell = f"<td>{'<br/>'.join(bits) if bits else '—'}</td>"
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
            f"{note_cell}"
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
  font-family: Figtree, "Noto Sans Georgian", "Noto Sans", ui-sans-serif, system-ui, sans-serif;
  background: var(--bg); color: var(--fg); min-height: 100vh; line-height: 1.45;
}
.ka { font-family: "Noto Sans Georgian", Figtree, sans-serif; }
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
  position: relative;
  background: var(--card); border-radius: var(--radius); box-shadow: var(--shadow);
  padding: 1rem 1.05rem; display: flex; flex-direction: column; gap: .55rem;
  text-decoration: none; transition: box-shadow .15s;
}
.card:hover { box-shadow: var(--shadow-hover); }
.card-stretch { position: absolute; inset: 0; z-index: 0; border-radius: inherit; }
.card .owners, .card .note-flag { position: relative; z-index: 1; }
.note-flag {
  font-size: .72rem; color: #9a3412; background: #ffedd5; border-radius: 6px;
  padding: .25rem .45rem; width: fit-content; max-width: 100%;
  overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
}
.note-flag.muted-flag { color: var(--muted); background: var(--secondary); }
.note-panel .ka { font-size: 1rem; line-height: 1.55; white-space: pre-wrap; }
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

.count-row {
  margin-top: 1.25rem; display: flex; flex-wrap: wrap; gap: .75rem 1.25rem;
  align-items: center; justify-content: space-between;
}
.count-row .count { margin-top: 0; }
.view-toggle {
  display: inline-flex; background: var(--card); box-shadow: var(--shadow);
  border-radius: 999px; padding: .2rem; gap: .15rem;
}
.view-toggle button {
  border: 0; background: transparent; font: inherit; font-size: .8rem;
  padding: .4rem .85rem; border-radius: 999px; color: var(--muted); cursor: pointer;
}
.view-toggle button:hover { color: var(--fg); }
.view-toggle button.active { background: var(--accent); color: #fff; }
body.view-cards .view-table-target { display: none !important; }
body.view-table .view-cards-target { display: none !important; }
.section-head {
  display: flex; align-items: center; justify-content: space-between; gap: 1rem;
  width: 100%; border: 0; background: transparent; padding: 0; cursor: pointer;
  font: inherit; text-align: left; color: inherit;
}
.section-head h2 { margin: 0; }
.section-head .chev { color: var(--muted); font-size: .9rem; transition: transform .15s; }
.section.is-collapsed .chev { transform: rotate(-90deg); }
.section.is-collapsed .section-body { display: none; }
"""



VIEW_TOGGLE_HTML = """
<div class="count-row">
  <p class="count" id="view-count">{count}</p>
  <div class="view-toggle" data-view-toggle role="group" aria-label="Layout">
    <button type="button" data-view="cards">Cards</button>
    <button type="button" data-view="table">Table</button>
    <button type="button" data-view="both">Both</button>
  </div>
</div>
"""

VIEW_JS = r"""
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
"""


def view_toggle_bar(count_html: str = "", count_id: str = "view-count") -> str:
    return f"""
<div class="count-row">
  <p class="count" id="{count_id}">{count_html}</p>
  <div class="view-toggle" data-view-toggle role="group" aria-label="Layout">
    <button type="button" data-view="cards">Cards</button>
    <button type="button" data-view="table">Table</button>
    <button type="button" data-view="both">Both</button>
  </div>
</div>
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
        ("status.html", "TGA & NNI", "status"),
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
<link href="https://fonts.googleapis.com/css2?family=Figtree:wght@400;500;600;700&family=Fraunces:opsz,wght@9..144,500;9..144,600&family=Noto+Sans+Georgian:wght@400;500;600;700&display=swap" rel="stylesheet"/>
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
<script src="assets/view.js"></script>
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
    const hay = [p.product, p.category, p.stage, p.due || "", p.technology || "", p.cpo || "", p.tpo || "", p.statusUpdate || "", p.whereUsed || "", ...(p.headProducts || [])].join(" ").toLowerCase();
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
function syncUrl() {
  const u = new URL(location.href);
  const map = { due: state.due, stage: state.stage, platform: state.head, category: state.category, cpo: state.cpo, tpo: state.tpo };
  for (const [k, v] of Object.entries(map)) {
    if (v) u.searchParams.set(k, v);
    else u.searchParams.delete(k);
  }
  const view = (window.ugtView && ugtView.read()) || "both";
  if (view && view !== "both") u.searchParams.set("view", view);
  else u.searchParams.delete("view");
  history.replaceState({}, "", u);
}
function render() {
  const rows = PRODUCTS.filter(matches);
  document.getElementById("view-count").textContent = `${rows.length} line${rows.length===1?"":"s"} in view.`;
  const cards = document.getElementById("cards");
  const tbody = document.getElementById("tbody");
  const empty = document.getElementById("empty");
  if (!rows.length) { cards.innerHTML=""; tbody.innerHTML=""; empty.hidden=false; return; }
  empty.hidden = true;
  cards.innerHTML = rows.map(p => {
    let note = "";
    if (p.statusUpdate) note = `<div class="note-flag" title="${esc(p.statusUpdate)}">📝 Status note</div>`;
    else if (p.whereUsed) note = `<div class="note-flag muted-flag" title="${esc(p.whereUsed)}">Where: ${esc(p.whereUsed)}</div>`;
    return `
    <article class="card">
      <a class="card-stretch" href="product-${p.n}.html" aria-label="${esc(p.product)}"></a>
      <div class="card-top"><div class="card-title">${esc(p.product)}</div><span class="badge ${badgeClass(p.stage)}">${esc(p.stage)}</span></div>
      <div class="meta">${esc(p.category || "—")}</div>
      <div class="heads">${(p.headProducts||[]).map(h=>`<span class="pill">${esc(h)}</span>`).join("")}</div>
      <div class="meta">${esc(p.due || "Unscheduled")}${p.technology ? " · " + esc(p.technology) : ""}</div>
      ${note}
      <div class="owners">${ownerChip(p.cpo,"CPO")}${ownerChip(p.tpo,"TPO")}</div>
    </article>`;
  }).join("");
  tbody.innerHTML = rows.map(p => {
    const bits = [];
    if (p.statusUpdate) bits.push(`<span class="ka">${esc(p.statusUpdate)}</span>`);
    if (p.whereUsed) bits.push(`<span class="muted">${esc(p.whereUsed)}</span>`);
    return `
    <tr>
      <td><a href="product-${p.n}.html"><strong>${esc(p.product)}</strong></a></td>
      <td>${esc(p.category || "—")}</td>
      <td>${esc((p.headProducts||[]).join(", "))}</td>
      <td><span class="badge ${badgeClass(p.stage)}">${esc(p.stage)}</span></td>
      <td>${esc(p.due || "Unscheduled")}</td>
      <td>${esc(p.cpo || "Unassigned")}</td>
      <td>${esc(p.tpo || "Unassigned")}</td>
      <td>${esc(p.technology || "—")}</td>
      <td>${bits.length ? bits.join("<br/>") : "—"}</td>
    </tr>`;
  }).join("");
  syncUrl();
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
""" + view_toggle_bar() + """
      <div class="cards view-cards-target" id="cards"></div>
      <div class="table-wrap view-table-target"><table>
        <thead><tr><th>Product</th><th>Category</th><th>Head product</th><th>Stage</th><th>Due</th><th>CPO</th><th>TPO</th><th>Technology</th><th>TGA &amp; NNI</th></tr></thead>
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
            f'<section class="section" data-collapsible id="due-{esc(due).replace(" ","-")}">'
            f'<button type="button" class="section-head" aria-expanded="true"><h2>{esc(due)}</h2><span class="chev">▾</span></button>'
            f'<div class="section-body">'
            f'<p class="muted">Target from the register. <strong>{len(group)}</strong></p>'
            f'<div class="cards view-cards-target">{"".join(product_card_html(p) for p in group)}</div>'
            f'<div class="table-wrap view-table-target"><table><thead><tr><th>Product</th><th>Category</th><th>Head product</th><th>Stage</th><th>Due</th><th>CPO</th><th>TPO</th><th>Technology</th><th>TGA &amp; NNI</th></tr></thead>'
            f"<tbody>{product_table_rows(group)}</tbody></table></div></div></section>"
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
            f'<section class="section" data-collapsible id="stage-{stage.lower().replace(" ","-")}">'
            f'<button type="button" class="section-head" aria-expanded="true"><h2>{esc(stage)}</h2><span class="chev">▾</span></button>'
            f'<div class="section-body">'
            f'<p class="muted">{esc(blurbs.get(stage, ""))} <strong>{len(group)}</strong></p>'
            f'<div class="cards view-cards-target">{"".join(product_card_html(p) for p in group)}</div>'
            f'<div class="table-wrap view-table-target"><table><thead><tr><th>Product</th><th>Category</th><th>Head product</th><th>Stage</th><th>Due</th><th>CPO</th><th>TPO</th><th>Technology</th><th>TGA &amp; NNI</th></tr></thead>'
            f"<tbody>{product_table_rows(group)}</tbody></table></div></div></section>"
        )
    body = f"""
      <p class="eyebrow">2026 register</p>
      <h1>Roadmap</h1>
      <p class="lede">Year-quarter targets first, then the same register by lifecycle stage.</p>
""" + view_toggle_bar(count_html="Cards / table layout") + f"""
      <p class="muted" style="margin-top:1rem">By quarter — click a heading to collapse</p>
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
""" + view_toggle_bar(count_html=f"{len(owned)} lines · Cards / table") + """
"""
    if as_cpo:
        body += f"""
      <section class="section" data-collapsible>
        <button type="button" class="section-head" aria-expanded="true"><h2>As commercial owner</h2><span class="chev">▾</span></button>
        <div class="section-body">
        <p class="muted">Accountable for the offer, packaging and go-to-market. · {len(as_cpo)} as CPO · {len(as_tpo)} as TPO</p>
        <div class="cards view-cards-target">{''.join(product_card_html(p) for p in as_cpo)}</div>
        <div class="table-wrap view-table-target"><table><thead><tr><th>Product</th><th>Category</th><th>Head product</th><th>Stage</th><th>Due</th><th>CPO</th><th>TPO</th><th>Technology</th><th>TGA &amp; NNI</th></tr></thead>
        <tbody>{product_table_rows(as_cpo)}</tbody></table></div>
        </div>
      </section>
"""
    if as_tpo:
        body += f"""
      <section class="section" data-collapsible>
        <button type="button" class="section-head" aria-expanded="true"><h2>As technical owner</h2><span class="chev">▾</span></button>
        <div class="section-body">
        <p class="muted">Accountable for technical delivery, architecture and operational readiness.</p>
        <div class="cards view-cards-target">{''.join(product_card_html(p) for p in as_tpo)}</div>
        <div class="table-wrap view-table-target"><table><thead><tr><th>Product</th><th>Category</th><th>Head product</th><th>Stage</th><th>Due</th><th>CPO</th><th>TPO</th><th>Technology</th><th>TGA &amp; NNI</th></tr></thead>
        <tbody>{product_table_rows(as_tpo)}</tbody></table></div>
        </div>
      </section>
"""
    return shell("owners", code, body, products, updated)


def build_product_page(p: dict, products: list[dict], updated: str) -> str:
    tga_bits = []
    if p.get("statusUpdate"):
        tga_bits.append(f'<div><div class="eyebrow">Status update</div><p class="ka note-panel" style="margin:.5rem 0 0">{esc(p["statusUpdate"])}</p></div>')
    if p.get("whereUsed"):
        tga_bits.append(f'<div style="margin-top:1rem"><div class="eyebrow">Where can it be used</div><p style="margin:.5rem 0 0">{esc(p["whereUsed"])}</p></div>')
    if tga_bits:
        tga_panel = (
            '<div class="panel section note-panel"><h2>TGA &amp; NNI</h2>'
            '<p class="muted" style="margin:.35rem 0 1rem;font-size:.875rem">From the TGA &amp; NNI sheet.</p>'
            + "".join(tga_bits)
            + '<p style="margin-top:1rem;font-size:.85rem"><a href="status.html">All status notes →</a></p></div>'
        )
    else:
        tga_panel = (
            '<div class="panel section"><h2>TGA &amp; NNI</h2>'
            '<p class="muted" style="margin:.5rem 0 0">No status note or usage note on the TGA &amp; NNI sheet for this line. '
            '<a href="status.html">Browse all notes</a>.</p></div>'
        )
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
      {tga_panel}
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


def build_status_notes(products: list[dict], tga_rows: list[dict], updated: str) -> str:
    with_status = sum(1 for t in tga_rows if t.get("statusUpdate"))
    with_where = sum(1 for t in tga_rows if t.get("whereUsed"))
    matched = sum(1 for t in tga_rows if t.get("matchedProductN") is not None)
    unmatched = [t for t in tga_rows if t.get("product") and t.get("matchedProductN") is None]
    orphans = [t for t in tga_rows if not t.get("product")]
    data_json = json.dumps(tga_rows, ensure_ascii=False)
    body = f"""
      <p class="eyebrow">TGA &amp; NNI sheet</p>
      <h1>Status notes</h1>
      <p class="lede">Working notes and “where can it be used” from the TGA &amp; NNI register. Georgian text is kept as written.</p>
      <div class="kpi-grid" style="margin-top:1.25rem">
        <div class="kpi"><div class="k">Rows</div><div class="v">{len(tga_rows)}</div></div>
        <div class="kpi"><div class="k">With status</div><div class="v">{with_status}</div></div>
        <div class="kpi"><div class="k">With where</div><div class="v">{with_where}</div></div>
        <div class="kpi"><div class="k">Matched to ALL</div><div class="v">{matched}</div></div>
      </div>
      <div class="filters" style="margin-top:1.5rem">
        <div class="filter-row" id="note-filters">
          <span class="filter-label">Show</span>
          <button type="button" class="chip active" data-val="">All</button>
          <button type="button" class="chip" data-val="status">Has status note</button>
          <button type="button" class="chip" data-val="where">Has where-used</button>
          <button type="button" class="chip" data-val="unmatched">Unmatched</button>
        </div>
        <div class="selects">
          <label>Stage<select id="f-stage"><option value="">All</option></select></label>
          <label>Head product<select id="f-head"><option value="">All</option></select></label>
          <label>Category<select id="f-cat"><option value="">All</option></select></label>
          <label class="search-box" style="margin-left:0;min-width:14rem">Search
            <input id="q" type="search" placeholder="Search notes (incl. Georgian)" autocomplete="off"/>
          </label>
        </div>
      </div>
      <p class="count" id="view-count"></p>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>N</th><th>Product</th><th>Status update</th><th>Where used</th>
              <th>Head</th><th>Category</th><th>Stage</th><th>Due</th><th>CPO</th><th>TPO</th><th>ALL link</th>
            </tr>
          </thead>
          <tbody id="tbody"></tbody>
        </table>
      </div>
      <div class="empty" id="empty" hidden>No rows match these filters.</div>
"""
    js = f"""
<script>
const ROWS = {data_json};
const state = {{ filter: "", stage: "", head: "", category: "", q: "" }};
function esc(s) {{
  return String(s ?? "").replace(/[&<>"']/g, c => ({{"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}}[c]));
}}
function badgeClass(stage) {{
  return "badge-" + String(stage || "").toLowerCase().replace(/\\s+/g, "-");
}}
function matches(t) {{
  if (state.filter === "status" && !t.statusUpdate) return false;
  if (state.filter === "where" && !t.whereUsed) return false;
  if (state.filter === "unmatched" && (t.matchedProductN != null || !t.product)) return false;
  if (state.stage && t.stage !== state.stage) return false;
  if (state.head && !(t.headProducts || []).includes(state.head)) return false;
  if (state.category && t.category !== state.category) return false;
  if (state.q) {{
    const hay = [t.product, t.statusUpdate, t.whereUsed, t.category, t.stage, t.due, t.cpo, t.tpo, ...(t.headProducts||[])].join(" ").toLowerCase();
    if (!hay.includes(state.q)) return false;
  }}
  return true;
}}
function render() {{
  const rows = ROWS.filter(matches);
  document.getElementById("view-count").textContent = `${{rows.length}} of ${{ROWS.length}} rows in view.`;
  const empty = document.getElementById("empty");
  const tbody = document.getElementById("tbody");
  if (!rows.length) {{ tbody.innerHTML = ""; empty.hidden = false; return; }}
  empty.hidden = true;
  tbody.innerHTML = rows.map(t => {{
    const link = t.matchedProductN != null
      ? `<a href="product-${{t.matchedProductN}}.html">#${{t.matchedProductN}}</a>`
      : (t.product ? `<span class="muted">Unmatched</span>` : `<span class="muted">Orphan note</span>`);
    const title = t.product
      ? (t.matchedProductN != null ? `<a href="product-${{t.matchedProductN}}.html"><strong>${{esc(t.product)}}</strong></a>` : `<strong>${{esc(t.product)}}</strong>`)
      : `<span class="muted">(no product)</span>`;
    return `<tr>
      <td>${{t.n ?? "—"}}</td>
      <td>${{title}}</td>
      <td class="ka">${{esc(t.statusUpdate || "—")}}</td>
      <td>${{esc(t.whereUsed || "—")}}</td>
      <td>${{esc((t.headProducts||[]).join(", ") || "—")}}</td>
      <td>${{esc(t.category || "—")}}</td>
      <td>${{t.stage ? `<span class="badge ${{badgeClass(t.stage)}}">${{esc(t.stage)}}</span>` : "—"}}</td>
      <td>${{esc(t.due || "—")}}</td>
      <td>${{esc(t.cpo || "—")}}</td>
      <td>${{esc(t.tpo || "—")}}</td>
      <td>${{link}}</td>
    </tr>`;
  }}).join("");
}}
function bind() {{
  const stages = [...new Set(ROWS.map(r => r.stage).filter(Boolean))];
  const heads = [...new Set(ROWS.flatMap(r => r.headProducts || []).filter(Boolean))].sort();
  const cats = [...new Set(ROWS.map(r => r.category).filter(Boolean))].sort();
  function fill(id, values) {{
    const sel = document.getElementById(id);
    for (const v of values) {{ const o=document.createElement("option"); o.value=v; o.textContent=v; sel.appendChild(o); }}
  }}
  fill("f-stage", stages); fill("f-head", heads); fill("f-cat", cats);
  document.getElementById("note-filters").addEventListener("click", e => {{
    const btn = e.target.closest(".chip"); if (!btn) return;
    state.filter = btn.dataset.val;
    document.querySelectorAll("#note-filters .chip").forEach(c => c.classList.toggle("active", c === btn));
    render();
  }});
  document.getElementById("f-stage").addEventListener("change", e => {{ state.stage = e.target.value; render(); }});
  document.getElementById("f-head").addEventListener("change", e => {{ state.head = e.target.value; render(); }});
  document.getElementById("f-cat").addEventListener("change", e => {{ state.category = e.target.value; render(); }});
  document.getElementById("q").addEventListener("input", e => {{ state.q = e.target.value.trim().toLowerCase(); render(); }});
  render();
}}
bind();
</script>
"""
    return shell("status", "TGA & NNI", body, products, updated, extra_js=js)



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
    tga_rows = load_tga_nni(EXCEL)
    attach_tga_notes(products, tga_rows)
    updated = data_updated_label()
    OUT_JSON.write_text(json.dumps(products, ensure_ascii=False, indent=2) + "\n")
    OUT_TGA_JSON.write_text(json.dumps(tga_rows, ensure_ascii=False, indent=2) + "\n")
    SITE.mkdir(parents=True, exist_ok=True)
    (SITE / "assets").mkdir(exist_ok=True)
    (SITE / "assets" / "site.css").write_text(SHARED_CSS)
    (SITE / "assets" / "view.js").write_text(VIEW_JS)
    (SITE / "index.html").write_text(build_overview(products, updated), encoding="utf-8")
    (SITE / "products.html").write_text(build_products(products, updated), encoding="utf-8")
    (SITE / "platforms.html").write_text(build_platforms(products, updated), encoding="utf-8")
    (SITE / "roadmap.html").write_text(build_roadmap(products, updated), encoding="utf-8")
    (SITE / "owners.html").write_text(build_owners_index(products, updated), encoding="utf-8")
    (SITE / "status.html").write_text(build_status_notes(products, tga_rows, updated), encoding="utf-8")
    (SITE / "glossary.html").write_text(build_glossary(products, updated), encoding="utf-8")
    codes = sorted({p["cpo"] for p in products if p["cpo"]} | {p["tpo"] for p in products if p["tpo"]})
    for code in codes:
        (SITE / f"owner-{code}.html").write_text(build_owner_page(code, products, updated), encoding="utf-8")
    for p in products:
        (SITE / f"product-{p['n']}.html").write_text(build_product_page(p, products, updated), encoding="utf-8")
    (SITE / "assets" / "data.js").write_text(
        "window.UGT_PRODUCTS = " + json.dumps(products, ensure_ascii=False) + ";\n"
        + "window.UGT_TGA_NNI = " + json.dumps(tga_rows, ensure_ascii=False) + ";\n"
        + "window.UGT_UPDATED = " + json.dumps(updated) + ";\n",
        encoding="utf-8",
    )
    matched = sum(1 for t in tga_rows if t.get("matchedProductN") is not None)
    print(f"Built {len(products)} products, {len(tga_rows)} TGA&NNI rows ({matched} matched) → {SITE} (updated: {updated})")


if __name__ == "__main__":
    main()
