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
OUT_PROJECTS_JSON = ROOT / "projects.json"
CURRENT_PROJECTS: list[dict] = []
SITE = ROOT / "docs"  # GitHub Pages /docs

STAGE_FIX = {"Pre-Prodcution": "Pre-Production"}
HEAD_FIX = {"ALL": "Cross-portfolio"}
# Preferred lifecycle order; actual STAGES is rebuilt from the Excel in main().
STAGE_ORDER = [
    "Production",
    "Pre-Production",
    "Deployment",
    "Development",
    "Pre-Development",
    "Wishlist",
    "Pipeline",
    "Terminated",
]
STAGES = list(STAGE_ORDER)
STAGE_COLORS: dict[str, str] = {}  # stage -> "#RRGGBB", read from the Excel Stage cell fills
FALLBACK_COLORS = ["#8E7CC3", "#E06666", "#6FA8DC", "#F6B26B", "#93C47D", "#C27BA0", "#76A5AF", "#FFD966", "#A4C2F4", "#B4A7D6"]
DELIVERY_STAGES = ("Pre-Production", "Deployment", "Development", "Pre-Development")
BACKLOG_STAGES = ("Wishlist", "Pipeline")


def text_on(hex_color: str) -> str:
    h = hex_color.lstrip("#")
    r, g, b = (int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
    lin = lambda c: c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    L = 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b)
    # pick whichever of black / white contrasts better
    return "#111111" if (L + 0.05) / 0.05 >= 1.05 / (L + 0.05) else "#ffffff"


def cell_fill_hex(cell) -> str | None:
    try:
        f = cell.fill
        if f is None or f.fill_type != "solid":
            return None
        rgb = f.fgColor.rgb
        if isinstance(rgb, str) and len(rgb) in (6, 8) and rgb.upper() not in ("00000000", "FFFFFFFF"):
            return "#" + rgb[-6:].upper()
    except Exception:
        pass
    return None


def finalize_stages(products: list[dict]) -> None:
    global STAGES
    present = []
    for p in products:
        if p["stage"] and p["stage"] not in present:
            present.append(p["stage"])
    STAGES = [s for s in STAGE_ORDER if s in present] + sorted(s for s in present if s not in STAGE_ORDER)
    used = {c.upper() for c in STAGE_COLORS.values()}
    pool = [c for c in FALLBACK_COLORS if c.upper() not in used]
    for s in STAGES:
        if s not in STAGE_COLORS:
            STAGE_COLORS[s] = pool.pop(0) if pool else "#9CA3AF"


SOFT_STAGE = {  # stage -> (background, text, dot)
    "Production":     ("#E3F4EA", "#1E6B3E", "#2F9E5B"),
    "Pre-Production": ("#EDF6E1", "#4A6B1F", "#7FB241"),
    "Deployment":     ("#E2F1FA", "#185E85", "#2E94C9"),
    "Development":    ("#FBF3D5", "#7A5C0A", "#D9A916"),
    "Wishlist":       ("#FCEBDD", "#8A4512", "#E07B2E"),
    "Terminated":     ("#FBE4E4", "#962B2B", "#D64545"),
    "Pre-Development":("#EFE8F8", "#5B3A8C", "#8E68C9"),
    "Pipeline":       ("#ECEEF2", "#454B57", "#8A93A3"),
}
SOFT_FALLBACK = [("#E6EEFB", "#2F4F8F", "#5B7FD1"), ("#F6E6F0", "#7D2F5E", "#C0569A"),
                 ("#E2F3F1", "#1F625B", "#3BA597"), ("#F3EDE4", "#6B4E2A", "#A9824F")]


def stage_css() -> str:
    out = ["/* soft stage palette (hue family matches the Excel fills) */",
           ".badge { display:inline-flex; align-items:center; gap:.35rem; font-weight:600; border-radius:999px; }",
           ".badge::before { content:''; width:.45rem; height:.45rem; border-radius:50%; background:var(--dot, currentColor); flex:none; }"]
    i = 0
    for s in STAGES:
        if s in SOFT_STAGE:
            bg, fg, dot = SOFT_STAGE[s]
        else:
            bg, fg, dot = SOFT_FALLBACK[i % len(SOFT_FALLBACK)]; i += 1
        out.append(f".{badge_class(s)} {{ background:{bg}; color:{fg}; --dot:{dot}; border:1px solid {dot}33; }}")
    return "\n".join(out) + "\n"
# All platform values used in filters, the matrix, and membership parsing.
PLATFORM_COLS = [
    "Universal Cloud",
    "Public Cloud",
    "Private Cloud",
    "Colocation",
    "SMB",
    "Bare Metal",
    "Standalone",
    "Cross-portfolio",
]
# Sidebar dropdown + dedicated platform pages only (not Standalone or Cross-portfolio / ALL).
NAV_PLATFORMS = [
    "Universal Cloud",
    "Public Cloud",
    "Private Cloud",
    "Colocation",
    "SMB",
    "Bare Metal",
]
# Set True later to pull Status Update / where-used from the NNI sheet again.
# When re-enabling, attach_tga_notes fills only blank ALL fields and does not overwrite them.
USE_NNI = False

PDDF_PHASE_LABELS = {
    "HD": "High Level Design",
    "Discovery": "Discovery",
    "Solution": "Solution",
    "Pilot": "Pilot",
}

# Used only when the Excel Description cell is empty. Tornike's later Excel text wins.
FALLBACK_DESCRIPTIONS = {
    "GPU VPS": "GPU VPS — იზოლირებული ვირტუალური სერვერი SMB-სთვის, საკუთარი გამოყოფილი GPU-ით.",
    "AI Router": "AI Router — UGT Cloud-ის როუტერი, სადაც კლიენტებს შეუძლიათ სხვადასხვა AI მოდელის მიბმა და გამოყენება ლარში გადახდით, მათ შორის UGT Cloud-ის საკუთარი მოდელის ტოკენების გამოყენებაც.",
    "AI Factory": "AI Factory — UGT Cloud-ის საკუთარი მოდელი ჩვენს GPU-ებზე; კლიენტებს შეუძლიათ მისი API-ს მიბმა და ჩვენი ტოკენების გამოყენება.",
}
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
    ("Platform", "The commercial environment the line belongs to. Sub-products are used only inside that platform. Standalone products can be sold independently and may be used on several platforms."),
    ("Head Product", "Earlier name for Platform. A product may have more than one platform."),
    ("Universal Cloud", "UGT’s universal cloud / virtual data centre platform."),
    ("Public Cloud", "Public cloud environment / platform."),
    ("Private Cloud", "Dedicated private cloud environment."),
    ("Colocation", "UGT data-centre colocation offers."),
    ("Bare Metal", "Physical dedicated server and appliance portfolio."),
    ("SMB", "SMB-focused packaged cloud and services portfolio."),
    ("Standalone", "Product that can be sold independently of the cloud platforms above."),
    ("ALL / Cross-portfolio", "Cross-portfolio service applicable across platforms."),
    ("Type — Product", "A sellable offer in its own right."),
    ("Type — Sub-Prod", "A capability used only inside its parent platform, not sold as a standalone offer."),
    ("Internal projects", "Internal tools, portal work and platform projects tracked in a second table on the ALL sheet. They are not commercial products and are listed under Internal projects."),
    ("Category", "Commercial grouping, for example Cloud Platform, VPS or Object Storages."),
    ("Class", "Technical class, for example Compute, Network, Storage or Backup."),
    ("Features", "Short tags for notable capabilities of the line."),
    ("CPO — Commercial Product Owner", "The person accountable for the commercial offer: positioning, packaging and go-to-market."),
    ("TPO — Technical Product Owner", "The person accountable for technical delivery, architecture and operational readiness."),
    ("Production", "Live / operational product."),
    ("Pre-Production", "Final preparation before production."),
    ("Development", "Actively being built."),
    ("Pre-Development", "Concept, discovery or preparation before development."),
    ("Deployment", "Being rolled out / deployed towards production."),
    ("Wishlist", "Wanted or planned product, not yet in delivery."),
    ("Pipeline", "Planned product / backlog item."),
    ("Terminated", "No longer active in the portfolio."),
    ("In Delivery", "Executive roll-up of Pre-Production, Deployment and Development."),
    ("Due Date", "Target year and quarter from the register, for example 2026 Q4. Unscheduled means the team has not set one yet."),
    ("PDDF Phase", "Product definition and delivery phase from the register."),
    ("PDDF — Discovery", "Early exploration of the offer."),
    ("PDDF — HD (High Level Design)", "High Level Design — the solution is being designed at a high level."),
    ("PDDF — Solution", "The solution is defined and being prepared."),
    ("PDDF — Pilot", "A pilot or limited rollout is under way."),
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


def clean_multiline(v) -> str:
    """Trim a free-text cell but keep the author's line breaks and wording verbatim."""
    if v is None:
        return ""
    s = str(v).replace("\xa0", " ").replace("\r\n", "\n").replace("\r", "\n")
    lines = [re.sub(r"[ \t]+", " ", ln).strip() for ln in s.split("\n")]
    out, blank = [], False
    for ln in lines:
        if not ln:
            if out and not blank:
                out.append("")
            blank = True
            continue
        out.append(ln)
        blank = False
    return "\n".join(out).strip()


# Header aliases (matched case-insensitively, by name — never by column position)
HEADER_ALIASES = {
    "Description": ("description", "descriptions", "product description", "აღწერა", "აღწერილობა"),
    "Product": ("product", "name", "product name"),
    "Head Product": ("head product", "head products", "head category", "platform", "platforms"),
    "Category": ("class",),
    "Group": ("category",),
    "Type": ("type",),
    "PortalName": ("product name in portal", "product name in the portal", "portal name", "portal product name", "name in portal"),
    "Platform/Technology/Vendor": ("platform/technology/vendor", "technology/vendor", "technology / vendor"),
    "Status Update": ("status update", "status"),
    "Features": ("features", "feature"),
    "where can it be used": ("where can it be used", "where used"),
}


def find_header_row(ws, must_have=("n",), scan_rows: int = 15) -> int:
    for r in range(1, min(scan_rows, ws.max_row) + 1):
        vals = {clean_text(ws.cell(r, c).value).lower() for c in range(1, ws.max_column + 1)}
        if all(m in vals for m in must_have):
            return r
    return 3


# Tokens that mark a repeated column-header row in the ALL sheet (second table, etc.).
TABLE_HEADER_TOKENS = {
    "n", "name", "description", "type", "platform", "category", "class",
    "technology/vendor", "technology / vendor", "platform/technology/vendor",
    "features", "pddf phase", "phase", "stage", "due date", "cpo", "tpo",
    "where can it be used", "product name in portal", "status update",
}


def is_table_header_row(ws, r: int, ncols: int) -> bool:
    """True when a row is a column-header bar (Name + several header labels), not a data row."""
    vals = [clean_text(ws.cell(r, c).value).lower() for c in range(1, ncols + 1)]
    vals = [v for v in vals if v]
    if "name" not in vals:
        return False
    hits = sum(1 for v in vals if v in TABLE_HEADER_TOKENS)
    return hits >= 4


def find_table_header_rows(ws) -> list[int]:
    """Every header bar on the ALL sheet, in row order. First = products, second = internal projects."""
    ncols = ws.max_column
    found = []
    for r in range(1, ws.max_row + 1):
        if is_table_header_row(ws, r, ncols):
            found.append(r)
    return found


def project_slug(name: str, used: set[str] | None = None) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", (name or "").lower()).strip("-") or "project"
    if used is not None:
        base = slug
        i = 2
        while slug in used:
            slug = f"{base}-{i}"
            i += 1
        used.add(slug)
    return slug


def item_href(p: dict, base: str = ".") -> str:
    if p.get("href"):
        href = p["href"]
    elif p.get("kind") == "project":
        href = f"project-{p['id']}.html"
    else:
        href = f"product-{p['n']}.html"
    if base and base not in (".", ""):
        return f"{base.rstrip('/')}/{href}"
    return href


def canonical_header(h: str) -> str:
    low = (h or "").lower().strip()
    # The internal-projects table labels this column " Phase" (no PDDF prefix).
    if low in ("phase", "pddf phase", "pddf"):
        return "PDDF"
    if re.match(r"^p+d+f", re.sub(r"[^a-z]", "", low)):
        return "PDDF"
    for canon, aliases in HEADER_ALIASES.items():
        if low in aliases:
            return canon
    return h


def pddf_link(ws, r, headers) -> str | None:
    """Hyperlink on the PDDF cell, or a URL typed into it."""
    if "PDDF" not in headers:
        return None
    c = ws.cell(r, headers.index("PDDF") + 1)
    if c.hyperlink is not None and c.hyperlink.target:
        return c.hyperlink.target
    v = clean_text(c.value)
    return v if re.match(r"^https?://", v or "") else None


def pddf_label(code: str | None) -> str:
    if not code:
        return ""
    pretty = PDDF_PHASE_LABELS.get(code)
    if pretty and pretty != code:
        return f"{code} — {pretty}"
    return pretty or code


def pddf_html(p: dict) -> str:
    if not p.get("pddf") and not p.get("pddfUrl"):
        return "—"
    label = pddf_label(p.get("pddf")) or p.get("pddf") or "PDF"
    if p.get("pddfUrl"):
        if re.match(r"^https?://", p.get("pddf") or ""):
            label = "📄 PDF"
        return f'<a href="{esc(p["pddfUrl"])}" target="_blank" rel="noopener">📄 {esc(label)}</a>' if not label.startswith("📄") else f'<a href="{esc(p["pddfUrl"])}" target="_blank" rel="noopener">{label}</a>'
    return esc(label)


def display_spelling(s: str) -> str:
    """Site-only spelling fixes; the Excel is left unchanged."""
    if not s:
        return s
    return re.sub(r"\bStandart\b", "Standard", s)


def split_features(raw) -> list[str]:
    if raw is None:
        return []
    parts = re.split(r"[\n;,/]+", str(raw).replace("\r", ""))
    return [display_spelling(p.strip()) for p in parts if p and p.strip()]


def platform_slug(name: str) -> str:
    return "platform-" + re.sub(r"[^a-z0-9]+", "-", (name or "").lower()).strip("-")


def platform_nav_list(products: list[dict]) -> list[str]:
    """Platforms shown in the sidebar dropdown and as dedicated pages."""
    present = {h for p in products for h in (p.get("headProducts") or []) if h and h != "Unplaced"}
    return [h for h in NAV_PLATFORMS if h in present]


def platform_matrix_list(products: list[dict]) -> list[str]:
    """Columns on the All platforms matrix, including Standalone and Cross-portfolio."""
    present = {h for p in products for h in (p.get("headProducts") or []) if h and h != "Unplaced"}
    return [h for h in PLATFORM_COLS if h in present] + sorted(h for h in present if h not in PLATFORM_COLS)


def parse_where_platforms(text: str) -> tuple[list[str], bool]:
    """Return (platforms mentioned, is_everywhere). Tolerates missing spaces, e.g. Private CloudBare Metal."""
    raw = clean_text(text)
    if not raw:
        return [], False
    low = raw.lower()
    everywhere = "everywhere" in low or "every platform" in low
    known = list(PLATFORM_COLS)
    found = []
    for name in sorted(known, key=len, reverse=True):
        pat = r"\s*".join(re.escape(part) for part in name.split())
        if re.search(pat, raw, re.I):
            found.append(name)
    # keep PLATFORM_COLS order
    found = [h for h in known if h in found]
    return found, everywhere


def product_on_platform(p: dict, platform: str) -> str | None:
    """'native' if the Platform column is this platform, 'available' if where-used names it.

    "Everywhere, in every platform" is display-only and does not add a product to a platform page.
    """
    if platform in (p.get("headProducts") or []):
        return "native"
    if platform in (p.get("wherePlatforms") or []):
        return "available"
    return None


def type_badge_html(p: dict) -> str:
    t = p.get("type") or ""
    if not t:
        return ""
    if t == "Sub-Prod":
        title = "Sub-product of " + ", ".join(p["headProducts"]) if p.get("headProducts") else "Sub-product of this platform"
    elif t in ("Product",):
        title = "Standalone or sellable product"
    else:
        title = t
    return f'<span class="type-badge type-{badge_class(t).removeprefix("badge-")}" title="{esc(title)}">{esc(t)}</span>'


def feature_tags_html(p: dict) -> str:
    feats = p.get("features") or []
    if not feats:
        return ""
    return '<div class="feat-tags">' + "".join(f'<span class="feat-tag">{esc(f)}</span>' for f in feats) + "</div>"


def _parse_register_row(ws, r: int, headers: list[str], kind: str, seq: int, slug_used: set[str]) -> dict | None:
    """Turn one ALL-sheet data row into a product or internal-project record. Skip blank and header rows."""
    ncols = len(headers)
    if is_table_header_row(ws, r, ws.max_column):
        return None
    row = {}
    for c in range(1, ncols + 1):
        row.setdefault(headers[c - 1], ws.cell(r, c).value)
    name = clean_text(row.get("Product"))
    if not name and row.get("N") is None:
        return None
    # Belt and braces: never ingest a repeated header as a data row (Name / Type / Stage ...).
    typ = clean_text(row.get("Type")) or None
    stage_raw = clean_text(row.get("Stage"))
    if name.lower() == "name" and (typ or "").lower() == "type" and (stage_raw or "").lower() == "stage":
        return None
    stage = STAGE_FIX.get(stage_raw, stage_raw)
    if stage and stage not in STAGE_COLORS and "Stage" in headers:
        col = headers.index("Stage") + 1
        hx = cell_fill_hex(ws.cell(r, col))
        if hx:
            STAGE_COLORS[stage] = hx
    heads = split_heads(row.get("Head Product"))
    if not heads:
        heads = [] if kind == "project" else ["Unplaced"]
    due = clean_text(row.get("Due Date")) or None
    tech = clean_text(row.get("Platform/Technology/Vendor"))
    sub = clean_text(row.get("Sub Products"))
    raw_n = row.get("N")
    try:
        n = int(raw_n)
    except (TypeError, ValueError):
        n = seq
    desc = clean_multiline(row.get("Description")) or None
    if kind != "project" and not desc:
        # Build-time fallback only when the Excel cell is empty; later Excel text wins.
        desc = FALLBACK_DESCRIPTIONS.get(name) or FALLBACK_DESCRIPTIONS.get(display_spelling(name))
    group = clean_text(row.get("Group")) or None
    klass = clean_text(row.get("Category")) or None
    where_raw = clean_multiline(row.get("where can it be used")) or None
    where_platforms, where_everywhere = parse_where_platforms(where_raw or "")
    status = clean_multiline(row.get("Status Update")) or None
    rec = {
        "n": n,
        "product": display_spelling(name),
        "headProducts": heads,
        "type": typ,
        "group": group,
        "category": klass,  # technical Class (kept as category for older links)
        "cls": klass,
        "technology": display_spelling(tech_display(tech, sub)),
        "stage": stage,
        "due": due,
        "cpo": owner_or_none(row.get("CPO")),
        "tpo": owner_or_none(row.get("TPO")),
        "description": desc,
        "portalName": display_spelling(clean_text(row.get("PortalName"))) or None,
        "features": split_features(row.get("Features")),
        "pddf": clean_text(row.get("PDDF")) or None,
        "pddfUrl": pddf_link(ws, r, headers),
        "statusUpdate": status,
        "whereUsed": where_raw,
        "wherePlatforms": where_platforms,
        "whereEverywhere": where_everywhere,
        "kind": kind,
    }
    if kind == "project":
        rec["id"] = project_slug(rec["product"], slug_used)
        rec["href"] = f"project-{rec['id']}.html"
    else:
        rec["href"] = f"product-{rec['n']}.html"
    return rec


def load_register(path: Path) -> tuple[list[dict], list[dict]]:
    """Read the ALL sheet as two tables: products (first header) and internal projects (second header).

    The header row of each table is never ingested. Future rows added under the second header
    are picked up automatically as internal projects.
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["UGT Product Portfolio ALL"]
    ncols = ws.max_column
    header_rows = find_table_header_rows(ws)
    if not header_rows:
        header_rows = [find_header_row(ws)]
    product_hdr = header_rows[0]
    project_hdr = header_rows[1] if len(header_rows) > 1 else None
    product_end = project_hdr if project_hdr else (ws.max_row + 1)

    def headers_at(hdr_row: int) -> list[str]:
        return [canonical_header(clean_text(ws.cell(hdr_row, c).value)) or f"col{c}" for c in range(1, ncols + 1)]

    products = []
    prod_headers = headers_at(product_hdr)
    for r in range(product_hdr + 1, product_end):
        rec = _parse_register_row(ws, r, prod_headers, "product", len(products) + 1, set())
        if rec:
            products.append(rec)

    projects = []
    if project_hdr:
        proj_headers = headers_at(project_hdr)
        slugs: set[str] = set()
        for r in range(project_hdr + 1, ws.max_row + 1):
            rec = _parse_register_row(ws, r, proj_headers, "project", len(projects) + 1, slugs)
            if rec:
                projects.append(rec)
    return products, projects


def load_products(path: Path) -> list[dict]:
    products, _ = load_register(path)
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
    raw_headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    # Normalise header names (leading space on Status Update)
    headers = [canonical_header(clean_text(h)) if h else f"col{c}" for c, h in enumerate(raw_headers, 1)]
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
        "PortalName": "portalName",
    }
    rows = []
    for r in range(2, ws.max_row + 1):
        raw = {}
        for c in range(1, len(headers) + 1):
            raw.setdefault(headers[c - 1], ws.cell(r, c).value)
        mapped = {}
        for k, v in raw.items():
            key = rename.get(k, k)
            if key in ("statusUpdate", "whereUsed", "portalName", "product", "headProduct", "technology", "subProducts", "category", "stage", "due", "cpo", "tpo"):
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
        p.setdefault("statusUpdate", None)
        p.setdefault("whereUsed", None)
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
        if t.get("statusUpdate") and not pick.get("statusUpdate"):
            pick["statusUpdate"] = t["statusUpdate"]
        if t.get("whereUsed") and not pick.get("whereUsed"):
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
    return "badge-" + re.sub(r"[^a-z0-9]+", "-", (stage or "").lower()).strip("-")


def fattrs(p: dict) -> str:
    """data-* attributes consumed by assets/filter.js"""
    heads = p.get("headProducts") or []
    text = " ".join(str(x) for x in [p["product"], p.get("type") or "", p.get("group") or "", p.get("cls") or "",
                    p.get("category") or "", p["stage"], p.get("due") or "",
                    p.get("technology") or "", p.get("cpo") or "", p.get("tpo") or "", p.get("description") or "",
                    p.get("portalName") or "", p.get("pddf") or "", " ".join(p.get("features") or []),
                    " ".join(heads), p.get("whereUsed") or "", p.get("statusUpdate") or ""]).lower()
    typ = p.get("type") or "Unspecified"
    nid = p.get("n") if p.get("n") is not None else p.get("id") or ""
    return (f' data-fitem data-n="{esc(str(nid))}" data-f-stage="{esc(p["stage"])}" data-f-due="{esc(p.get("due") or "Unscheduled")}"'
            f' data-f-head="{esc("|".join(heads))}" data-f-type="{esc(typ)}"'
            f' data-f-group="{esc(p.get("group") or "")}" data-f-cls="{esc(p.get("cls") or "")}"'
            f' data-f-category="{esc(p.get("group") or p.get("category") or "")}"'
            f' data-f-pddf="{esc(p.get("pddf") or "Not set")}"'
            f' data-f-cpo="{esc(p.get("cpo") or "Unassigned")}" data-f-tpo="{esc(p.get("tpo") or "Unassigned")}"'
            f' data-ftext="{esc(text)}"')


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


def product_card_html(p: dict, base: str = ".", platform: str | None = None) -> str:
    heads = "".join(f'<span class="pill">{esc(h)}</span>' for h in p["headProducts"])
    due = p["due"] or "Unscheduled"
    tech = f" · {esc(p['technology'])}" if p.get("technology") else ""
    note = ""
    if p.get("statusUpdate"):
        note = f'<div class="note-flag" title="{esc(p["statusUpdate"])}">📝 Status note</div>'
    where = ""
    if p.get("whereUsed"):
        where = (f'<div class="where-used"><span class="where-label">Where can it be used</span>'
                 f'{esc(p["whereUsed"])}</div>')
    desc = ""
    if p.get("description"):
        desc = f'<p class="card-desc ka" title="{esc(p["description"])}">{esc(p["description"])}</p>'
    portal = f'<div class="card-portal" title="Product name in portal">🏷 {esc(p["portalName"])}</div>' if p.get("portalName") else ""
    belong = ""
    if platform and (p.get("type") == "Sub-Prod") and platform in (p.get("headProducts") or []):
        belong = f'<div class="subprod-flag">Sub-product of {esc(platform)}</div>'
    elif platform and product_on_platform(p, platform) == "available":
        belong = f'<div class="available-flag">Also available on {esc(platform)}</div>'
    group = p.get("group") or ""
    klass = p.get("cls") or p.get("category") or ""
    meta_bits = " · ".join(b for b in (group, klass) if b) or "—"
    return f"""
    <article class="card"{fattrs(p)}>
      <a class="card-stretch" href="{item_href(p, base)}" aria-label="{esc(p['product'])}"></a>
      <div class="card-top">
        <div><div class="card-title">{esc(p['product'])}</div>{portal}{belong}</div>
        <span class="card-badges">{type_badge_html(p)}<span class="badge {badge_class(p['stage'])}">{esc(p['stage'])}</span></span>
      </div>
      <div class="meta">{esc(meta_bits)}</div>
      <div class="heads">{heads}</div>
      <div class="meta">{esc(due)}{tech}</div>
      {feature_tags_html(p)}
      {desc}
      {where}
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
            note_cell = f"<td>{'<br/>'.join(bits) if bits else '—'}</td>"
        typ = type_badge_html(p) or "—"
        feats = ", ".join(p.get("features") or []) or "—"
        rows.append(
            f"<tr{fattrs(p)}>"
            f'<td><a href="{item_href(p, base)}"><strong>{esc(p["product"])}</strong></a></td>'
            f"<td>{typ}</td>"
            f"<td>{esc(p.get('group') or '—')}</td>"
            f"<td>{esc(p.get('cls') or p.get('category') or '—')}</td>"
            f"<td>{esc(', '.join(p['headProducts']))}</td>"
            f'<td><span class="badge {badge_class(p["stage"])}">{esc(p["stage"])}</span></td>'
            f"<td>{esc(p['due'] or 'Unscheduled')}</td>"
            f"<td>{esc(p['cpo'] or 'Unassigned')}</td>"
            f"<td>{esc(p['tpo'] or 'Unassigned')}</td>"
            f"<td>{esc(p['technology'] or '—')}</td>"
            f"<td>{esc(feats)}</td>"
            f"<td>{esc(p.get('portalName') or '—')}</td>"
            f"<td>{pddf_html(p)}</td>"
            f"<td>{esc(p.get('whereUsed') or '—')}</td>"
            f"{note_cell}"
            f"</tr>"
        )
    return "\n".join(rows)


PRODUCT_TABLE_HEAD = (
    "<th>Product</th><th>Type</th><th>Category</th><th>Class</th><th>Platform</th>"
    "<th>Stage</th><th>Due</th><th>CPO</th><th>TPO</th><th>Technology</th>"
    "<th>Features</th><th>Portal name</th><th>PDDF phase</th>"
    "<th>Where can it be used</th><th>Notes</th>"
)


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
/* Mobile nav lives in the top header. Never show it on desktop — nav.side { display:flex }
   would otherwise override the hidden attribute and overlay the page. */
#mobile-nav { display: none !important; margin-top: 0; }
@media (max-width: 960px) {
  .app { grid-template-columns: 1fr; }
  aside.rail { display: none !important; }
  .menu-btn { display: inline-flex !important; }
  header.top { flex-wrap: wrap; }
  #mobile-nav { display: none !important; width: 100%; flex: 1 0 100%; padding: .5rem 0 0; }
  #mobile-nav.open { display: flex !important; flex-direction: column; gap: .25rem; }
}
aside.rail {
  position: sticky; top: 0; height: 100vh; background: var(--rail);
  border-right: 1px solid var(--border); padding: 1.25rem;
  display: flex; flex-direction: column;
}
.brand { display: flex; flex-direction: column; align-items: flex-start; gap: .4rem; text-decoration: none; color: inherit; padding: .1rem .25rem 0; }
.brand picture { display: block; line-height: 0; }
.brand-logo { display: block; width: 172px; max-width: 100%; height: auto; }
.logo {
  width: 2.25rem; height: 2.25rem; border-radius: 8px; background: var(--accent);
  color: var(--accent-fg); display: grid; place-items: center; font-weight: 700; font-size: .85rem;
}
.brand .sub { font-size: 11px; letter-spacing: .28em; text-transform: uppercase; color: var(--muted); padding-left: 2px; }
nav.side { margin-top: 2rem; display: flex; flex-direction: column; gap: .25rem; flex: 1; }
nav.side a {
  display: flex; align-items: center; gap: .75rem; min-height: 2.75rem;
  padding: 0 .75rem; border-radius: 8px; text-decoration: none;
  color: var(--muted); font-size: .9rem;
}
nav.side a.active, nav.side a:hover { background: var(--secondary); color: var(--fg); }
nav.side .nav-drop { margin: 0; }
nav.side .nav-drop > summary {
  display: flex; align-items: center; justify-content: space-between; gap: .75rem;
  min-height: 2.75rem; padding: 0 .75rem; border-radius: 8px;
  color: var(--muted); font-size: .9rem; cursor: pointer; list-style: none;
}
nav.side .nav-drop > summary::-webkit-details-marker { display: none; }
nav.side .nav-drop > summary::after { content: "▾"; font-size: .75rem; opacity: .7; }
nav.side .nav-drop[open] > summary, nav.side .nav-drop > summary:hover,
nav.side .nav-drop.active > summary { background: var(--secondary); color: var(--fg); }
nav.side .nav-drop-list { display: flex; flex-direction: column; gap: .15rem; padding: .15rem 0 .35rem .4rem; }
nav.side .nav-drop-list a { min-height: 2.25rem; font-size: .85rem; padding-left: 1rem; }
.menu-btn {
  display: none; align-items: center; gap: .4rem; border: 0; background: var(--secondary);
  color: var(--fg); border-radius: 8px; min-height: 2.4rem; padding: 0 .85rem; font: inherit; font-size: .9rem; cursor: pointer;
}
.menu-btn:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }
.card-badges { display: flex; flex-wrap: wrap; gap: .35rem; align-items: center; justify-content: flex-end; }
.type-badge {
  display: inline-flex; align-items: center; font-weight: 600; font-size: .72rem; letter-spacing: .02em;
  border-radius: 999px; padding: .2rem .55rem; background: #E7EEF6; color: #3A4A5C; border: 1px solid #C9D6E4;
}
.type-badge.type-sub-prod { background: #EDE8F6; color: #4A3A6B; border-color: #D4CAE6; }
.type-badge.type-project { background: #EDE8F6; color: #4A3A6B; border-color: #D4CAE6; }
.type-badge.type-tools, .type-badge.type-portal { background: #E7EEF6; color: #3A4A5C; border-color: #C9D6E4; }
.feat-tags { display: flex; flex-wrap: wrap; gap: .3rem; margin-top: .35rem; position: relative; z-index: 1; }
.feat-tag {
  font-size: .72rem; background: #F3F6F9; color: #4A5A6A; border: 1px solid #D5DEE7;
  border-radius: 999px; padding: .12rem .5rem;
}
.subprod-flag { font-size: .75rem; color: #5B3A8C; margin-top: .2rem; font-weight: 600; }
.available-flag { font-size: .75rem; color: var(--muted); margin-top: .2rem; }
.where-used { font-size: .8rem; color: var(--muted); margin-top: .35rem; position: relative; z-index: 1; line-height: 1.35; }
.where-used .where-label { font-size: 10px; letter-spacing: .12em; text-transform: uppercase; display: block; margin-bottom: .1rem; color: var(--muted); }
.snap { background: var(--secondary); border-radius: 10px; padding: 1rem; }
.snap .label { font-size: 11px; letter-spacing: .16em; text-transform: uppercase; color: var(--muted); }
.snap .num { font-family: Fraunces, serif; font-size: 1.85rem; margin-top: .25rem; }
.snap .hint { font-size: .875rem; color: var(--muted); margin-top: .25rem; }
.snap a.hint { color: var(--muted); text-decoration: none; display: inline-block; }
.snap a.hint:hover { color: var(--fg); }
.updated { font-size: .7rem; color: var(--muted); margin-top: .6rem; line-height: 1.3; }
.card-portal { font-size:.78rem; color: var(--accent); margin-top:.15rem; font-weight:500; }
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
.filters { margin-top: 1.75rem; display: flex; flex-direction: column; gap: 1.35rem; }
.filter-row { display: flex; flex-wrap: wrap; gap: .45rem; align-items: center; }
.filter-row > .filter-label { flex: 0 0 100%; margin-bottom: .2rem; }
.filter-label { font-size: 11px; letter-spacing: .14em; text-transform: uppercase; color: var(--muted); }
.chip {
  display: inline-flex; align-items: center; justify-content: center;
  border: 0; background: #dfe9f5; box-shadow: none; border-radius: 999px;
  min-height: 2.6rem; padding: .55rem 1rem; font: inherit; font-size: .9rem; line-height: 1.2;
  color: #3d4f63; cursor: pointer; text-decoration: none; white-space: nowrap;
  transition: background-color .12s ease, color .12s ease;
}
.chip:hover { background: #d3e1f1; color: var(--fg); }
.chip:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }
.chip.active, .chip.active:hover { background: var(--accent); color: #fff; }
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
.card-desc {
  margin: 0; font-size: .82rem; line-height: 1.45; color: var(--fg); opacity: .85;
  display: -webkit-box; -webkit-line-clamp: 2; line-clamp: 2; -webkit-box-orient: vertical;
  overflow: hidden; overflow-wrap: anywhere;
}
.desc-panel { margin-top: 1.25rem; border-left: 4px solid var(--accent); }
.desc-panel .ka { font-size: 1.05rem; line-height: 1.65; white-space: pre-line; margin: .6rem 0 0; }
td.desc-cell { max-width: 22rem; }
td.desc-cell .desc-trunc {
  display: -webkit-box; -webkit-line-clamp: 2; line-clamp: 2; -webkit-box-orient: vertical;
  overflow: hidden; font-size: .82rem; cursor: help;
}
.badge {
  display: inline-flex; align-items: center; border-radius: 999px;
  padding: .2rem .55rem; font-size: .7rem; font-weight: 600; white-space: nowrap;
}
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
table.resizable th { position: sticky; }
table.rz-fixed { table-layout: fixed; }
table.rz-fixed th, table.rz-fixed td { overflow-wrap: anywhere; word-break: normal; white-space: normal; }
.col-resizer {
  position: absolute; top: 0; right: 0; width: 10px; height: 100%; z-index: 3;
  cursor: col-resize; touch-action: none; user-select: none;
}
.col-resizer::after {
  content: ""; position: absolute; top: 22%; bottom: 22%; left: 4px; width: 2px; border-radius: 2px;
  background: var(--border); transition: background-color .12s ease;
}
.col-resizer::after { left: auto; right: 0; }
th:hover > .col-resizer::after { background: #a9bdd3; }
.col-resizer:hover::after, body.rz-dragging .col-resizer::after { background: var(--accent); }
body.rz-dragging { cursor: col-resize; user-select: none; }
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

RESIZE_JS = r"""
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
      if (!/products\.html|internal-projects\.html|roadmap\.html|owner-/.test(href)) return;
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


SORT_JS = r"""
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
"""

FILTER_CSS = r"""
th.sortable-th { cursor: pointer; user-select: none; }
th.sortable-th .sort-ind { margin-left: .3rem; font-size: .7em; opacity: .35; }
th.sortable-th[aria-sort="ascending"] .sort-ind, th.sortable-th[aria-sort="descending"] .sort-ind { opacity: 1; color: var(--accent); }
.ffilters { margin-top: 1.5rem; }
.ffilters .ftop { display:flex; flex-wrap:wrap; gap:1rem; align-items:center; justify-content:space-between; }
.fcount { font-size:.85rem; color: var(--muted); }
[data-fitem].fhide, [data-fsection].fhide { display: none !important; }
"""

FILTER_JS = r"""
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
"""

HEAD_ORDER = ["Universal Cloud", "Public Cloud", "Private Cloud", "Colocation", "SMB", "Bare Metal", "Standalone", "Cross-portfolio"]


def filter_bar(groups: list[tuple], noun: str = "product", placeholder: str = "Search") -> str:
    """groups: (key, label, order list, tail list, labels dict)"""
    rows = []
    for g in groups:
        key, label, order = g[0], g[1], g[2]
        tail = g[3] if len(g) > 3 else []
        labels = g[4] if len(g) > 4 else {}
        rows.append(
            f'<div class="filter-row" data-fkey="{key}" data-forder=\'{esc(json.dumps(order, ensure_ascii=False))}\' '
            f'data-ftail=\'{esc(json.dumps(tail, ensure_ascii=False))}\' data-flabels=\'{esc(json.dumps(labels, ensure_ascii=False))}\' '
            f'role="group" aria-label="{esc(label)}"><span class="filter-label">{esc(label)}</span></div>'
        )
    return (
        f'<div class="ffilters filters" data-fnoun="{noun}">'
        f'<div class="ftop"><span class="fcount"></span><label class="search-box"><span aria-hidden="true">⌕</span>'
        f'<input data-fsearch type="search" placeholder="{esc(placeholder)}" autocomplete="off"/></label></div>'
        + "".join(rows) + "</div>"
    )


def product_filter_groups(include_due: bool = True) -> list[tuple]:
    pddf_labels = {code: pddf_label(code) for code in PDDF_PHASE_LABELS}
    g = [("stage", "Stage", STAGES), ("type", "Type", ["Product", "Sub-Prod"], ["Unspecified"])]
    if include_due:
        g.append(("due", "Due date", [], ["Unscheduled"]))
    g += [("head", "Platform", HEAD_ORDER, ["Unplaced"]),
          ("group", "Category", []),
          ("cls", "Class", []),
          ("cpo", "Commercial owner", [], ["Unassigned"]), ("tpo", "Technical owner", [], ["Unassigned"]),
          ("pddf", "PDDF phase", [], ["Not set"], pddf_labels)]
    return g


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
    delivery = sum(1 for p in products if p["stage"] in DELIVERY_STAGES)
    pipe = sum(1 for p in products if p["stage"] in BACKLOG_STAGES)
    plat_list = platform_nav_list(products)
    def nav_links(active_page: str, drop_open: bool = False) -> str:
        items = [
            ("index.html", "Overview", "overview"),
            ("products.html", "Products", "products"),
        ]
        bits = [
            f'<a class="{"active" if active_page == key else ""}" href="{href}">{label}</a>'
            for href, label, key in items
        ]
        plat_open = drop_open or active_page in ("platforms", "platform")
        drop_cls = "nav-drop active" if plat_open else "nav-drop"
        open_attr = " open" if plat_open else ""
        links = [f'<a class="{"active" if active_page == "platforms" else ""}" href="platforms.html">All platforms</a>']
        for h in plat_list:
            slug = platform_slug(h) + ".html"
            on = active_page == f"platform:{h}"
            links.append(f'<a class="{"active" if on else ""}" href="{slug}">{esc(h)}</a>')
        bits.append(
            f'<details class="{drop_cls}"{open_attr}>'
            f"<summary>Platforms</summary>"
            f'<div class="nav-drop-list">{"".join(links)}</div></details>'
        )
        rest = [
            ("internal-projects.html", "Internal projects", "internal-projects"),
            ("roadmap.html", "Roadmap", "roadmap"),
            ("owners.html", "Owners", "owners"),
            ("status.html", "Status", "status"),
            ("glossary.html", "Legend", "glossary"),
        ]
        bits += [
            f'<a class="{"active" if active_page == key else ""}" href="{href}">{label}</a>'
            for href, label, key in rest
        ]
        return "".join(bits)

    nav_html = nav_links(page)
    mobile_html = nav_links(page, drop_open=True)
    nav_js = """
<script>
(function () {
  var btn = document.querySelector("[data-menu]");
  var nav = document.getElementById("mobile-nav");
  if (!btn || !nav) return;
  function setOpen(open) {
    nav.classList.toggle("open", open);
    nav.hidden = !open;
    btn.setAttribute("aria-expanded", open ? "true" : "false");
  }
  setOpen(false);
  btn.addEventListener("click", function () { setOpen(!nav.classList.contains("open")); });
  window.addEventListener("resize", function () {
    if (window.innerWidth > 960 && nav.classList.contains("open")) setOpen(false);
  });
})();
</script>
"""
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>{esc(title)} · UGT Cloud Portfolio</title>
<link rel="icon" href="favicon.ico" sizes="16x16 32x32 48x48"/>
<link rel="icon" type="image/png" sizes="32x32" href="assets/favicon-32.png"/>
<link rel="apple-touch-icon" sizes="180x180" href="assets/apple-touch-icon.png"/>
<meta name="theme-color" content="#153bb6"/>
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
    <a class="brand" href="index.html" aria-label="UGT Cloud Portfolio — overview">
      <picture>
        <source type="image/webp" srcset="assets/logo-344.webp 2x, assets/logo-516.webp 3x"/>
        <img class="brand-logo" src="assets/logo-344.png" srcset="assets/logo-344.png 2x, assets/logo-516.png 3x" width="172" height="70" alt="UGT Cloud"/>
      </picture>
      <span class="sub">Portfolio</span>
    </a>
    <nav class="side">{nav_html}</nav>
    <div class="snap">
      <div class="label">Snapshot</div>
      <div class="num">{len(products)}</div>
      <div class="hint">products in the 2026 register</div>
      {f'<a class="hint" href="internal-projects.html">{len(CURRENT_PROJECTS)} internal project{"s" if len(CURRENT_PROJECTS)!=1 else ""}</a>' if CURRENT_PROJECTS else ""}
      <div class="updated">Data last updated<br/>{esc(updated)}</div>
    </div>
  </aside>
  <div>
    <header class="top">
      <button type="button" class="menu-btn" data-menu aria-expanded="false" aria-controls="mobile-nav">Menu</button>
      <div class="stats">{live} live · {delivery} in delivery · {pipe} wishlist</div>
      <nav id="mobile-nav" class="side" hidden>{mobile_html}</nav>
    </header>
    <div class="content">
{body}
      <p class="footer">Built from UGT_Product_Portfolio_2026.xlsx · sheet “UGT Product Portfolio ALL” (products and internal projects).</p>
    </div>
  </div>
</div>
<script src="assets/view.js"></script>
{extra_js}
<script src="assets/resize.js"></script>
<script src="assets/filter.js"></script>
<script src="assets/sort.js"></script>
{nav_js}
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
    by_cat = Counter(p.get("group") or p.get("cls") or p.get("category") for p in products if p.get("group") or p.get("cls") or p.get("category"))
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
    delivery = sum(by_stage.get(s, 0) for s in DELIVERY_STAGES)
    owners_n = len(set(list(cpo_c) + list(tpo_c)))

    kpis = [
        ("Total products", len(products), "products.html"),
        ("Internal projects", len(CURRENT_PROJECTS), "internal-projects.html"),
        ("Production", by_stage.get("Production", 0), "products.html?stage=Production"),
        ("In delivery", delivery, "roadmap.html"),
        ("Scheduled", scheduled, "roadmap.html"),
        *[(s, by_stage.get(s, 0), f"products.html?stage={s}") for s in STAGES if s not in ("Production", "Terminated")],
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
      <p class="lede">Every UGT Cloud offer, the platform it sits on, and the people who own it commercially and technically.{(" Internal projects are listed separately.") if CURRENT_PROJECTS else ""}</p>
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
const state = { stage: "", due: "", head: "", type: "", group: "", cls: "", cpo: "", tpo: "", pddf: "", q: "" };
const PLATFORM_CONTEXT = __PLATFORM__;
const PDDF_LABELS = __PDDF_LABELS__;
function pddfLabel(code) { return (code && PDDF_LABELS[code] && PDDF_LABELS[code] !== code) ? (code + " — " + PDDF_LABELS[code]) : (PDDF_LABELS[code] || code || "—"); }
function typeBadge(p) {
  if (!p.type) return "";
  const cls = "type-badge type-" + String(p.type).toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
  const title = p.type === "Sub-Prod" ? ("Sub-product of " + (p.headProducts || []).join(", ")) : "Product";
  return `<span class="${cls}" title="${esc(title)}">${esc(p.type)}</span>`;
}
function featTags(p) {
  const feats = p.features || [];
  if (!feats.length) return "";
  return `<div class="feat-tags">${feats.map(f => `<span class="feat-tag">${esc(f)}</span>`).join("")}</div>`;
}

function badgeClass(stage) {
  return "badge-" + String(stage || "").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
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
  if (state.type === "Unspecified") { if (p.type) return false; }
  else if (state.type && p.type !== state.type) return false;
  if (state.group && p.group !== state.group) return false;
  if (state.cls && (p.cls || p.category) !== state.cls) return false;
  if (state.cpo === "__unassigned__") { if (p.cpo) return false; }
  else if (state.cpo && p.cpo !== state.cpo) return false;
  if (state.tpo === "__unassigned__") { if (p.tpo) return false; }
  else if (state.tpo && p.tpo !== state.tpo) return false;
  if (state.pddf === "__notset__") { if (p.pddf) return false; }
  else if (state.pddf && p.pddf !== state.pddf) return false;
  if (state.q) {
    const hay = [p.product, p.type || "", p.group || "", p.cls || p.category, p.stage, p.due || "", p.technology || "", p.cpo || "", p.tpo || "", p.statusUpdate || "", p.whereUsed || "", p.description || "", p.portalName || "", p.pddf || "", ...(p.features || []), ...(p.headProducts || [])].join(" ").toLowerCase();
    if (!hay.includes(state.q)) return false;
  }
  return true;
}
function renderChips(id, values, key, allLabel, extra) {
  const el = document.getElementById(id);
  if (!el) return;
  const label = el.querySelector(".filter-label");
  let html = "";
  html += `<button type="button" class="chip ${state[key]===""?"active":""}" data-key="${key}" data-val="" aria-pressed="${state[key]===""}">${allLabel}</button>`;
  for (const v of values) html += `<button type="button" class="chip ${state[key]===v?"active":""}" data-key="${key}" data-val="${esc(v)}" aria-pressed="${state[key]===v}">${esc(v)}</button>`;
  if (extra) for (const [lab, val] of extra) html += `<button type="button" class="chip ${state[key]===val?"active":""}" data-key="${key}" data-val="${esc(val)}" aria-pressed="${state[key]===val}">${esc(lab)}</button>`;
  el.innerHTML = "";
  if (label) el.appendChild(label);
  el.insertAdjacentHTML("beforeend", html);
}
function syncUrl() {
  const u = new URL(location.href);
  const map = { due: state.due, stage: state.stage, platform: state.head, type: state.type, category: state.group, class: state.cls, cpo: state.cpo, tpo: state.tpo, pddf: state.pddf };
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
    const where = p.whereUsed
      ? `<div class="where-used"><span class="where-label">Where can it be used</span>${esc(p.whereUsed)}</div>`
      : "";
    let belong = "";
    if (PLATFORM_CONTEXT && p.type === "Sub-Prod" && (p.headProducts || []).includes(PLATFORM_CONTEXT)) {
      belong = `<div class="subprod-flag">Sub-product of ${esc(PLATFORM_CONTEXT)}</div>`;
    } else if (PLATFORM_CONTEXT && !(p.headProducts || []).includes(PLATFORM_CONTEXT)) {
      belong = `<div class="available-flag">Also available on ${esc(PLATFORM_CONTEXT)}</div>`;
    }
    const meta = [p.group, p.cls || p.category].filter(Boolean).join(" · ") || "—";
    return `
    <article class="card">
      <a class="card-stretch" href="product-${p.n}.html" aria-label="${esc(p.product)}"></a>
      <div class="card-top"><div><div class="card-title">${esc(p.product)}</div>${p.portalName ? `<div class="card-portal" title="Product name in portal">🏷 ${esc(p.portalName)}</div>` : ""}${belong}</div><span class="card-badges">${typeBadge(p)}<span class="badge ${badgeClass(p.stage)}">${esc(p.stage)}</span></span></div>
      <div class="meta">${esc(meta)}</div>
      <div class="heads">${(p.headProducts||[]).map(h=>`<span class="pill">${esc(h)}</span>`).join("")}</div>
      <div class="meta">${esc(p.due || "Unscheduled")}${p.technology ? " · " + esc(p.technology) : ""}</div>
      ${featTags(p)}
      ${p.description ? `<p class="card-desc ka" title="${esc(p.description)}">${esc(p.description)}</p>` : ""}
      ${where}
      ${note}
      <div class="owners">${ownerChip(p.cpo,"CPO")}${ownerChip(p.tpo,"TPO")}</div>
    </article>`;
  }).join("");
  tbody.innerHTML = rows.map(p => {
    const bits = [];
    if (p.statusUpdate) bits.push(`<span class="ka">${esc(p.statusUpdate)}</span>`);
    return `
    <tr>
      <td><a href="product-${p.n}.html"><strong>${esc(p.product)}</strong></a></td>
      <td>${typeBadge(p) || "—"}</td>
      <td class="desc-cell">${p.description ? `<span class="ka desc-trunc" title="${esc(p.description)}">${esc(p.description)}</span>` : "—"}</td>
      <td>${esc(p.group || "—")}</td>
      <td>${esc(p.cls || p.category || "—")}</td>
      <td>${esc((p.headProducts||[]).join(", "))}</td>
      <td><span class="badge ${badgeClass(p.stage)}">${esc(p.stage)}</span></td>
      <td>${esc(p.due || "Unscheduled")}</td>
      <td>${esc(p.cpo || "Unassigned")}</td>
      <td>${esc(p.tpo || "Unassigned")}</td>
      <td>${esc(p.technology || "—")}</td>
      <td>${esc((p.features || []).join(", ") || "—")}</td>
      <td>${esc(p.portalName || "—")}</td>
      <td>${p.pddfUrl ? `<a href="${esc(p.pddfUrl)}" target="_blank" rel="noopener">📄 ${esc(/^https?:/.test(p.pddf||"") ? "PDF" : pddfLabel(p.pddf))}</a>` : esc(pddfLabel(p.pddf))}</td>
      <td>${esc(p.whereUsed || "—")}</td>
      <td>${bits.length ? bits.join("<br/>") : "—"}</td>
    </tr>`;
  }).join("");
  syncUrl();
}
function bind() {
  const HEAD_ORDER = ["Universal Cloud", "Public Cloud", "Private Cloud", "Colocation", "SMB", "Bare Metal", "Standalone", "Cross-portfolio"];
  const dues = [...new Set(PRODUCTS.map(p => p.due).filter(Boolean))].sort();
  const headSet = new Set(PRODUCTS.flatMap(p => p.headProducts || []).filter(Boolean));
  const heads = HEAD_ORDER.filter(h => headSet.has(h))
    .concat([...headSet].filter(h => !HEAD_ORDER.includes(h) && h !== "Unplaced").sort((a,b)=>a.localeCompare(b)))
    .concat(headSet.has("Unplaced") ? ["Unplaced"] : []);
  const groups = [...new Set(PRODUCTS.map(p => p.group).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
  const classes = [...new Set(PRODUCTS.map(p => p.cls || p.category).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
  const types = ["Product", "Sub-Prod"].filter(t => PRODUCTS.some(p => p.type === t));
  const unType = PRODUCTS.some(p => !p.type);
  function countBy(key) { const m = {}; PRODUCTS.forEach(p => { if (p[key]) m[p[key]] = (m[p[key]] || 0) + 1; }); return m; }
  const cpoCount = countBy("cpo");
  const cpos = Object.keys(cpoCount).sort((a,b) => (cpoCount[b]-cpoCount[a]) || a.localeCompare(b));   // biggest commercial owner first
  const tpos = Object.keys(countBy("tpo")).sort((a,b)=>a.localeCompare(b));
  const unCpo = PRODUCTS.some(p => !p.cpo) || state.cpo === "__unassigned__";
  const unTpo = PRODUCTS.some(p => !p.tpo) || state.tpo === "__unassigned__";
  const pddfCodes = [...new Set(PRODUCTS.map(p => p.pddf).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
  function renderPddf() {
    const el = document.getElementById("pddf-filters");
    if (!el) return;
    const label = el.querySelector(".filter-label");
    let html = `<button type="button" class="chip ${state.pddf===""?"active":""}" data-key="pddf" data-val="" aria-pressed="${state.pddf===""}">All</button>`;
    for (const v of pddfCodes) html += `<button type="button" class="chip ${state.pddf===v?"active":""}" data-key="pddf" data-val="${esc(v)}" aria-pressed="${state.pddf===v}">${esc(pddfLabel(v))}</button>`;
    html += `<button type="button" class="chip ${state.pddf==="__notset__"?"active":""}" data-key="pddf" data-val="__notset__" aria-pressed="${state.pddf==="__notset__"}">Not set</button>`;
    el.innerHTML = "";
    if (label) el.appendChild(label);
    el.insertAdjacentHTML("beforeend", html);
  }
  const GROUPS = {
    stage:    () => renderChips("stage-filters", STAGES, "stage", "All"),
    type:     () => renderChips("type-filters", types, "type", "All", unType ? [["Unspecified","Unspecified"]] : null),
    due:      () => renderChips("due-filters", dues, "due", "All", [["Unscheduled","Unscheduled"]]),
    head:     () => renderChips("head-filters", heads, "head", "All"),
    group:    () => renderChips("group-filters", groups, "group", "All"),
    cls:      () => renderChips("cls-filters", classes, "cls", "All"),
    cpo:      () => renderChips("cpo-filters", cpos, "cpo", "All", unCpo ? [["Unassigned","__unassigned__"]] : null),
    tpo:      () => renderChips("tpo-filters", tpos, "tpo", "All", unTpo ? [["Unassigned","__unassigned__"]] : null),
    pddf:     renderPddf,
  };
  const params = new URLSearchParams(location.search);
  if (params.get("due")) state.due = params.get("due");
  if (params.get("stage")) state.stage = params.get("stage");
  if (params.get("platform")) state.head = params.get("platform");
  if (params.get("type")) state.type = params.get("type");
  if (params.get("category")) state.group = params.get("category");
  if (params.get("class")) state.cls = params.get("class");
  if (params.get("cpo")) state.cpo = params.get("cpo");
  if (params.get("tpo")) state.tpo = params.get("tpo");
  if (params.get("pddf")) state.pddf = params.get("pddf");
  Object.values(GROUPS).forEach(fn => fn());
  document.getElementById("filters").addEventListener("click", e => {
    const btn = e.target.closest(".chip[data-key]"); if (!btn) return;
    const key = btn.dataset.key;
    if (!(key in GROUPS)) return;
    state[key] = btn.dataset.val;
    GROUPS[key]();
    render();
  });
  document.getElementById("q").addEventListener("input", e => { state.q = e.target.value.trim().toLowerCase(); render(); });
  render();
}
bind();
</script>
"""


def products_catalog_body(title: str, lede: str, eyebrow: str = "2026 register") -> str:
    return f"""
      <div style="display:flex;flex-wrap:wrap;gap:1rem;align-items:flex-end;justify-content:space-between">
        <div>
          <p class="eyebrow">{esc(eyebrow)}</p>
          <h1>{esc(title)}</h1>
          <p class="lede">{esc(lede)}</p>
        </div>
        <label class="search-box"><span aria-hidden="true">⌕</span>
          <input id="q" type="search" placeholder="Search products, descriptions, owners, platforms" autocomplete="off"/>
        </label>
      </div>
      <div class="filters" id="filters">
        <div class="filter-row" id="stage-filters" role="group" aria-label="Stage"><span class="filter-label">Stage</span></div>
        <div class="filter-row" id="type-filters" role="group" aria-label="Type"><span class="filter-label">Type</span></div>
        <div class="filter-row" id="due-filters" role="group" aria-label="Due date"><span class="filter-label">Due date</span></div>
        <div class="filter-row" id="head-filters" role="group" aria-label="Platform"><span class="filter-label">Platform</span></div>
        <div class="filter-row" id="group-filters" role="group" aria-label="Category"><span class="filter-label">Category</span></div>
        <div class="filter-row" id="cls-filters" role="group" aria-label="Class"><span class="filter-label">Class</span></div>
        <div class="filter-row" id="cpo-filters" role="group" aria-label="Commercial owner"><span class="filter-label">Commercial owner</span></div>
        <div class="filter-row" id="tpo-filters" role="group" aria-label="Technical owner"><span class="filter-label">Technical owner</span></div>
        <div class="filter-row" id="pddf-filters" role="group" aria-label="PDDF phase"><span class="filter-label">PDDF phase</span></div>
      </div>
""" + view_toggle_bar() + f"""
      <div class="cards view-cards-target" id="cards"></div>
      <div class="table-wrap view-table-target"><table>
        <thead><tr><th>Product</th><th>Type</th><th>Description</th><th>Category</th><th>Class</th><th>Platform</th><th>Stage</th><th>Due</th><th>CPO</th><th>TPO</th><th>Technology</th><th>Features</th><th>Portal name</th><th>PDDF phase</th><th>Where can it be used</th><th>Notes</th></tr></thead>
        <tbody id="tbody"></tbody>
      </table></div>
      <div class="empty" id="empty" hidden>No products match these filters.</div>
"""


def catalog_js(rows: list[dict], platform: str | None = None) -> str:
    return (
        PRODUCTS_JS
        .replace("__PRODUCTS__", json.dumps(rows, ensure_ascii=False))
        .replace("__STAGES__", json.dumps(STAGES))
        .replace("__PLATFORM__", json.dumps(platform, ensure_ascii=False))
        .replace("__PDDF_LABELS__", json.dumps(PDDF_PHASE_LABELS, ensure_ascii=False))
    )


def build_products(products: list[dict], updated: str) -> str:
    body = products_catalog_body(
        "Products",
        "Filter by type, stage, quarter, platform, category, class or owner.",
    )
    return shell("products", "Products", body, products, updated, extra_js=catalog_js(products))


def build_platform_page(platform: str, products: list[dict], updated: str) -> str:
    rows = [p for p in products if product_on_platform(p, platform)]
    # Native (belongs to this platform) first, then also-available; Sub-Prods before Products within each group.
    def sort_key(p):
        native = 0 if product_on_platform(p, platform) == "native" else 1
        typ = 0 if p.get("type") == "Sub-Prod" else 1
        return (native, typ, p.get("group") or "", p["product"])
    rows = sorted(rows, key=sort_key)
    native_n = sum(1 for p in rows if product_on_platform(p, platform) == "native")
    avail_n = len(rows) - native_n
    lede = (
        f"Offers that belong to {platform}, plus standalone products that can be used here. "
        f"{native_n} belong to this platform"
        + (f"; {avail_n} are also available here" if avail_n else "")
        + ". Sub-products are marked as belonging to this platform."
    )
    body = products_catalog_body(platform, lede, eyebrow="Platform")
    return shell(f"platform:{platform}", platform, body, products, updated, extra_js=catalog_js(rows, platform))


def build_platforms(products: list[dict], updated: str) -> str:
    plat_cols = platform_matrix_list(products)
    counts = Counter()
    for p in products:
        for h in p["headProducts"]:
            counts[h] += 1
    rows = []
    for p in products:
        marks = "".join(
            f'<td class="mark">{"●" if h in p["headProducts"] else ""}</td>' for h in plat_cols
        )
        rows.append(
            f'<tr{fattrs(p)}><td><a href="product-{p["n"]}.html">{esc(p["product"])}</a> '
            f'<span class="muted">{esc(p["category"])}</span></td>{marks}'
            f'<td><span class="badge {badge_class(p["stage"])}">{esc(p["stage"])}</span></td></tr>'
        )
    short = {
        "Universal Cloud": "Universal",
        "Public Cloud": "Public",
        "Private Cloud": "Private",
        "Bare Metal": "Bare Metal",
        "SMB": "SMB",
        "Standalone": "Standalone",
        "Colocation": "Colocation",
        "Cross-portfolio": "All",
    }
    head_cells = "".join(
        f'<th><a href="{platform_slug(h)}.html">{esc(short.get(h, h))}</a></th>'
        if h in NAV_PLATFORMS
        else f'<th>{esc(short.get(h, h))}</th>'
        for h in plat_cols
    )
    body = f"""
      <p class="eyebrow">2026 register</p>
      <h1>Product × platform</h1>
      <p class="lede">A filled mark means the line belongs to that platform. Open a platform page from the sidebar for the same filters as Products, limited to that platform.</p>
      {filter_bar(product_filter_groups(include_due=False), "product", "Search products on platforms")}
      <div class="panel mt">
        <ul style="list-style:none;margin:0;padding:0;display:flex;flex-wrap:wrap;gap:1rem">
          {''.join(f'<li><a href="{platform_slug(h)}.html"><strong>{counts.get(h,0)}</strong> <span class="muted">{esc(h)}</span></a></li>' if h in NAV_PLATFORMS else f'<li><strong>{counts.get(h,0)}</strong> <span class="muted">{esc(h)}</span></li>' for h in platform_matrix_list(products))}
        </ul>
      </div>
      <div class="table-wrap matrix-wrap"><table class="matrix">
        <thead><tr><th>Product</th>{head_cells}<th>Stage</th></tr></thead>
        <tbody>{''.join(rows)}</tbody>
      </table></div>
      <div class="empty" id="fempty" hidden>No products match these filters.</div>
"""
    return shell("platforms", "Platforms", body, products, updated)


def build_roadmap(products: list[dict], updated: str) -> str:
    dues = sorted({p["due"] for p in products if p["due"]})
    sections = []
    for due in dues:
        group = [p for p in products if p["due"] == due]
        sections.append(
            f'<section class="section" data-fsection data-collapsible id="due-{esc(due).replace(" ","-")}">'
            f'<button type="button" class="section-head" aria-expanded="true"><h2>{esc(due)}</h2><span class="chev">▾</span></button>'
            f'<div class="section-body">'
            f'<p class="muted">Target from the register. <strong>{len(group)}</strong></p>'
            f'<div class="cards view-cards-target">{"".join(product_card_html(p) for p in group)}</div>'
            f'<div class="table-wrap view-table-target"><table><thead><tr>{PRODUCT_TABLE_HEAD}</tr></thead>'
            f"<tbody>{product_table_rows(group)}</tbody></table></div></div></section>"
        )
    # delivery stages first then production/pipeline/terminated like Grok
    stage_order = [s for s in STAGES if s in DELIVERY_STAGES] + [s for s in STAGES if s not in DELIVERY_STAGES and s not in ("Production", "Terminated")] + [s for s in ("Production", "Terminated") if s in STAGES]
    blurbs = {
        "Pre-Production": "Final preparation before going live.",
        "Development": "Actively being built.",
        "Pre-Development": "Discovery and preparation before build.",
        "Pipeline": "Planned backlog — not yet in delivery.",
        "Deployment": "Being rolled out towards production.",
        "Wishlist": "Wanted, not yet in delivery.",
        "Production": "Live and operational.",
        "Terminated": "Removed from the active portfolio.",
    }
    for stage in stage_order:
        group = [p for p in products if p["stage"] == stage]
        if not group:
            continue
        sections.append(
            f'<section class="section" data-fsection data-collapsible id="stage-{stage.lower().replace(" ","-")}">'
            f'<button type="button" class="section-head" aria-expanded="true"><h2>{esc(stage)}</h2><span class="chev">▾</span></button>'
            f'<div class="section-body">'
            f'<p class="muted">{esc(blurbs.get(stage, ""))} <strong>{len(group)}</strong></p>'
            f'<div class="cards view-cards-target">{"".join(product_card_html(p) for p in group)}</div>'
            f'<div class="table-wrap view-table-target"><table><thead><tr>{PRODUCT_TABLE_HEAD}</tr></thead>'
            f"<tbody>{product_table_rows(group)}</tbody></table></div></div></section>"
        )
    body = f"""
      <p class="eyebrow">2026 register</p>
      <h1>Roadmap</h1>
      <p class="lede">Year-quarter targets first, then the same register by lifecycle stage.</p>
      {filter_bar(product_filter_groups(), "product", "Search the roadmap")}
""" + view_toggle_bar(count_html="Cards / table layout") + f"""
      <div class="empty" id="fempty" hidden>No products match these filters.</div>
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
            heads_o = sorted({h for p in owned for h in p["headProducts"]})
            cats_o = sorted({p["category"] for p in owned if p["category"]})
            ftext = " ".join([code, title, blurb] + [p["product"] for p in owned]).lower()
            out.append(
                f'<a class="owner-card" href="owner-{esc(code)}.html" data-fitem data-n="{kind}-{esc(code)}" '
                f'data-f-role="{"Commercial (CPO)" if kind == "cpo" else "Technical (TPO)"}" '
                f'data-f-stage="{esc("|".join(stages))}" data-f-head="{esc("|".join(heads_o))}" '
                f'data-f-category="{esc("|".join(cats_o))}" data-ftext="{esc(ftext)}">'
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
      {filter_bar([("role", "Role", ["Commercial (CPO)", "Technical (TPO)"]), ("stage", "Owns products in stage", STAGES), ("head", "Head product", HEAD_ORDER, ["Unplaced"]), ("category", "Category", [])], "owner", "Search owners or their products")}
      <div data-fsection><h2 class="section">Commercial</h2>
      <div class="cards" style="margin-top:1rem">{cards(cpos, "cpo")}</div></div>
      <div data-fsection><h2 class="section">Technical</h2>
      <div class="cards" style="margin-top:1rem">{cards(tpos, "tpo")}</div></div>
      <div class="empty" id="fempty" hidden>No owners match these filters.</div>
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
        <div class="table-wrap view-table-target"><table><thead><tr>{PRODUCT_TABLE_HEAD}</tr></thead>
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
        <div class="table-wrap view-table-target"><table><thead><tr>{PRODUCT_TABLE_HEAD}</tr></thead>
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
            '<div class="panel section note-panel"><h2>Notes</h2>'
            '<p class="muted" style="margin:.35rem 0 1rem;font-size:.875rem">From the product register.</p>'
            + "".join(tga_bits)
            + '<p style="margin-top:1rem;font-size:.85rem"><a href="status.html">All notes →</a></p></div>'
        )
    else:
        tga_panel = ""
    if p.get("description"):
        desc_panel = (
            '<div class="panel desc-panel"><h2>Description</h2>'
            f'<p class="ka">{esc(p["description"])}</p></div>'
        )
    else:
        desc_panel = (
            '<div class="panel desc-panel"><h2>Description</h2>'
            '<p class="muted" style="margin:.5rem 0 0">No description in the register yet.</p></div>'
        )
    portal_dd = f'<dt>Portal name</dt><dd><strong>{esc(p["portalName"])}</strong></dd>' if p.get("portalName") else ""
    if p.get("pddf") or p.get("pddfUrl"):
        portal_dd += f'<dt>PDDF phase</dt><dd>{pddf_html(p)}</dd>'
        if p.get("pddfUrl"):
            portal_dd += f'<dd><a class="chip active" href="{esc(p["pddfUrl"])}" target="_blank" rel="noopener">📄 Open PDF</a></dd>'

    body = f"""
      <p class="eyebrow"><a class="muted" href="products.html">Products</a> / {esc(p['product'])}</p>
      <div style="display:flex;flex-wrap:wrap;gap:1rem;align-items:flex-start;justify-content:space-between">
        <div>
          <h1>{esc(p['product'])}</h1>
          <p class="lede">{esc(' · '.join(b for b in (p.get('type'), p.get('group'), p.get('cls') or p.get('category')) if b) or '—')} · line #{p['n']}</p>{('<p class="card-portal" style="font-size:.95rem;margin:.25rem 0 0">Portal name: <strong>' + esc(p['portalName']) + '</strong></p>') if p.get('portalName') else ''}
          {feature_tags_html(p)}
        </div>
        <span class="card-badges">{type_badge_html(p)}<span class="badge {badge_class(p['stage'])}" style="font-size:.85rem;padding:.4rem .75rem">{esc(p['stage'])}</span></span>
      </div>
      {desc_panel}
      <div class="detail-grid">
        <div class="panel">
          <h2>Register details</h2>
          <dl class="dl" style="margin-top:1rem">
            <dt>Type</dt><dd>{esc(p.get('type') or '—')}</dd>
            <dt>Platform</dt><dd>{esc(', '.join(p['headProducts']))}</dd>
            <dt>Category</dt><dd>{esc(p.get('group') or '—')}</dd>
            <dt>Class</dt><dd>{esc(p.get('cls') or p.get('category') or '—')}</dd>
            <dt>Stage</dt><dd>{esc(p['stage'])}</dd>
            <dt>Due</dt><dd>{esc(p['due'] or 'Unscheduled')}</dd>
            <dt>Technology</dt><dd>{esc(p['technology'] or '—')}</dd>
            {portal_dd}
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


def status_rows_from_products(products: list[dict]) -> list[dict]:
    rows = []
    for p in products:
        if not (p.get("statusUpdate") or p.get("whereUsed")):
            continue
        rows.append({
            "n": p["n"],
            "product": p["product"],
            "statusUpdate": p.get("statusUpdate") or "",
            "whereUsed": p.get("whereUsed") or "",
            "portalName": p.get("portalName") or "",
            "headProducts": p.get("headProducts") or [],
            "category": p.get("group") or p.get("cls") or "",
            "stage": p.get("stage") or "",
            "due": p.get("due") or "",
            "cpo": p.get("cpo") or "",
            "tpo": p.get("tpo") or "",
            "matchedProductN": p["n"],
        })
    return rows


def build_status_notes(products: list[dict], tga_rows: list[dict], updated: str) -> str:
    # Prefer ALL-sourced notes; NNI rows are only used when USE_NNI is on.
    rows = list(tga_rows) if (USE_NNI and tga_rows) else status_rows_from_products(products)
    with_status = sum(1 for t in rows if t.get("statusUpdate"))
    with_where = sum(1 for t in rows if t.get("whereUsed"))
    matched = sum(1 for t in rows if t.get("matchedProductN") is not None)
    data_json = json.dumps(rows, ensure_ascii=False)
    source = "the NNI sheet" if (USE_NNI and tga_rows) else "the product register (ALL)"
    body = f"""
      <p class="eyebrow">Register notes</p>
      <h1>Status notes</h1>
      <p class="lede">Status updates and “where can it be used” from {source}. Georgian text is kept as written.</p>
      <div class="kpi-grid" style="margin-top:1.25rem">
        <div class="kpi"><div class="k">Rows</div><div class="v">{len(tga_rows)}</div></div>
        <div class="kpi"><div class="k">With status</div><div class="v">{with_status}</div></div>
        <div class="kpi"><div class="k">With where</div><div class="v">{with_where}</div></div>
        <div class="kpi"><div class="k">Matched to ALL</div><div class="v">{matched}</div></div>
      </div>
      <div class="filters" id="filters" style="margin-top:1.5rem">
        <label class="search-box" style="margin-left:0;max-width:26rem"><span aria-hidden="true">⌕</span>
          <input id="q" type="search" placeholder="Search notes (incl. Georgian)" autocomplete="off"/>
        </label>
        <div class="filter-row" id="note-filters" role="group" aria-label="Show"><span class="filter-label">Show</span></div>
        <div class="filter-row" id="stage-filters" role="group" aria-label="Stage"><span class="filter-label">Stage</span></div>
        <div class="filter-row" id="head-filters" role="group" aria-label="Head product"><span class="filter-label">Head product</span></div>
        <div class="filter-row" id="cat-filters" role="group" aria-label="Category"><span class="filter-label">Category</span></div>
      </div>
      <p class="count" id="view-count"></p>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>N</th><th>Product</th><th>Status update</th><th>Where used</th><th>Portal name</th>
              <th>Head</th><th>Category</th><th>Stage</th><th>Due</th><th>CPO</th><th>TPO</th><th>ALL link</th>
            </tr>
          </thead>
          <tbody id="tbody"></tbody>
        </table>
      </div>
      <div class="empty" id="empty" hidden>No rows match these filters.</div>
"""
    stages_json = json.dumps(STAGES)
    js = f"""
<script>
const ROWS = {data_json};
const state = {{ filter: "", stage: "", head: "", category: "", q: "" }};
function esc(s) {{
  return String(s ?? "").replace(/[&<>"']/g, c => ({{"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}}[c]));
}}
function badgeClass(stage) {{
  return "badge-" + String(stage || "").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
}}
function matches(t) {{
  if (state.filter === "status" && !t.statusUpdate) return false;
  if (state.filter === "where" && !t.whereUsed) return false;
  if (state.filter === "unmatched" && (t.matchedProductN != null || !t.product)) return false;
  if (state.stage && t.stage !== state.stage) return false;
  if (state.head && !(t.headProducts || []).includes(state.head)) return false;
  if (state.category && t.category !== state.category) return false;
  if (state.q) {{
    const hay = [t.product, t.statusUpdate, t.whereUsed, t.portalName || "", t.category, t.stage, t.due, t.cpo, t.tpo, ...(t.headProducts||[])].join(" ").toLowerCase();
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
      <td>${{esc(t.portalName || "—")}}</td>
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
function chipsHtml(key, values, extra) {{
  const one = (lab, val) => `<button type="button" class="chip ${{state[key]===val?"active":""}}" data-key="${{key}}" data-val="${{esc(val)}}" aria-pressed="${{state[key]===val}}">${{esc(lab)}}</button>`;
  return one("All", "") + values.map(v => one(v, v)).join("") + (extra || []).map(([l, v]) => one(l, v)).join("");
}}
function paint(id, key, values, extra) {{
  const el = document.getElementById(id);
  const label = el.querySelector(".filter-label");
  el.innerHTML = "";
  el.appendChild(label);
  el.insertAdjacentHTML("beforeend", chipsHtml(key, values, extra));
}}
function bind() {{
  const STAGE_ORDER = {stages_json};
  const HEAD_ORDER = ["Universal Cloud", "Public Cloud", "Private Cloud", "Colocation", "SMB", "Bare Metal", "Standalone", "Cross-portfolio"];
  const stageSet = new Set(ROWS.map(r => r.stage).filter(Boolean));
  const stages = STAGE_ORDER.filter(s => stageSet.has(s)).concat([...stageSet].filter(s => !STAGE_ORDER.includes(s)).sort());
  const headSet = new Set(ROWS.flatMap(r => r.headProducts || []).filter(Boolean));
  const heads = HEAD_ORDER.filter(h => headSet.has(h))
    .concat([...headSet].filter(h => !HEAD_ORDER.includes(h) && h !== "Unplaced").sort((a,b)=>a.localeCompare(b)))
    .concat(headSet.has("Unplaced") ? ["Unplaced"] : []);
  const cats = [...new Set(ROWS.map(r => r.category).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
  const GROUPS = {{
    filter:   () => paint("note-filters", "filter", [], [["Has status note","status"],["Has where-used","where"],["Unmatched","unmatched"]]),
    stage:    () => paint("stage-filters", "stage", stages),
    head:     () => paint("head-filters", "head", heads),
    category: () => paint("cat-filters", "category", cats),
  }};
  Object.values(GROUPS).forEach(fn => fn());
  document.getElementById("filters").addEventListener("click", e => {{
    const btn = e.target.closest(".chip[data-key]"); if (!btn) return;
    const key = btn.dataset.key;
    if (!(key in GROUPS)) return;
    state[key] = btn.dataset.val;
    GROUPS[key]();
    render();
  }});
  document.getElementById("q").addEventListener("input", e => {{ state.q = e.target.value.trim().toLowerCase(); render(); }});
  render();
}}
bind();
</script>
"""
    return shell("status", "Status notes", body, products, updated, extra_js=js)




PROJECT_TABLE_HEAD = (
    "<th>Project</th><th>Type</th><th>Description</th><th>Category</th>"
    "<th>Stage</th><th>Due</th><th>CPO</th><th>TPO</th><th>Technology</th>"
    "<th>Features</th><th>Phase</th>"
)

PROJECTS_JS = r"""
<script>
const PRODUCTS = __PRODUCTS__;
const STAGES = __STAGES__;
const state = { stage: "", due: "", type: "", group: "", cpo: "", tpo: "", pddf: "", q: "" };
const PDDF_LABELS = __PDDF_LABELS__;
function pddfLabel(code) { return (code && PDDF_LABELS[code] && PDDF_LABELS[code] !== code) ? (code + " — " + PDDF_LABELS[code]) : (PDDF_LABELS[code] || code || "—"); }
function typeBadge(p) {
  if (!p.type) return "";
  const cls = "type-badge type-" + String(p.type).toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
  return `<span class="${cls}" title="${esc(p.type)}">${esc(p.type)}</span>`;
}
function featTags(p) {
  const feats = p.features || [];
  if (!feats.length) return "";
  return `<div class="feat-tags">${feats.map(f => `<span class="feat-tag">${esc(f)}</span>`).join("")}</div>`;
}
function badgeClass(stage) {
  return "badge-" + String(stage || "").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
}
function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
}
function hrefOf(p) { return p.href || ("project-" + (p.id || p.n) + ".html"); }
function ownerChip(code, role) {
  if (!code) return `<span class="owner"><span class="avatar muted">—</span><span><div>Unassigned</div><div class="role">${role}</div></span></span>`;
  return `<a class="owner" href="owner-${esc(code)}.html"><span class="avatar">${esc(code)}</span><span><div>${esc(code)}</div><div class="role">${role}</div></span></a>`;
}
function matches(p) {
  if (state.stage && p.stage !== state.stage) return false;
  if (state.due === "Unscheduled") { if (p.due) return false; }
  else if (state.due && p.due !== state.due) return false;
  if (state.type === "Unspecified") { if (p.type) return false; }
  else if (state.type && p.type !== state.type) return false;
  if (state.group && p.group !== state.group) return false;
  if (state.cpo === "__unassigned__") { if (p.cpo) return false; }
  else if (state.cpo && p.cpo !== state.cpo) return false;
  if (state.tpo === "__unassigned__") { if (p.tpo) return false; }
  else if (state.tpo && p.tpo !== state.tpo) return false;
  if (state.pddf === "__notset__") { if (p.pddf) return false; }
  else if (state.pddf && p.pddf !== state.pddf) return false;
  if (state.q) {
    const hay = [p.product, p.type || "", p.group || "", p.stage, p.due || "", p.technology || "", p.cpo || "", p.tpo || "", p.description || "", p.pddf || "", ...(p.features || [])].join(" ").toLowerCase();
    if (!hay.includes(state.q)) return false;
  }
  return true;
}
function renderChips(id, values, key, allLabel, extra) {
  const el = document.getElementById(id);
  if (!el) return;
  let html = "";
  html += `<button type="button" class="chip ${state[key]===""?"active":""}" data-key="${key}" data-val="" aria-pressed="${state[key]===""}">${allLabel}</button>`;
  for (const v of values) html += `<button type="button" class="chip ${state[key]===v?"active":""}" data-key="${key}" data-val="${esc(v)}" aria-pressed="${state[key]===v}">${esc(v)}</button>`;
  if (extra) for (const [lab, val] of extra) html += `<button type="button" class="chip ${state[key]===val?"active":""}" data-key="${key}" data-val="${esc(val)}" aria-pressed="${state[key]===val}">${esc(lab)}</button>`;
  el.querySelectorAll("button").forEach(b => b.remove());
  el.insertAdjacentHTML("beforeend", html);
}
function applyUrl() {
  const u = new URL(location.href);
  for (const k of Object.keys(state)) {
    if (k === "q") { const q = u.searchParams.get("q"); if (q) state.q = q.toLowerCase(); continue; }
    const v = u.searchParams.get(k);
    if (v) state[k] = v;
  }
}
function syncUrl() {
  const u = new URL(location.href);
  for (const k of Object.keys(state)) {
    if (!state[k]) u.searchParams.delete(k);
    else u.searchParams.set(k, state[k]);
  }
  history.replaceState(null, "", u);
}
function render() {
  const rows = PRODUCTS.filter(matches);
  const count = document.getElementById("view-count");
  if (count) count.textContent = `${rows.length} project${rows.length===1?"":"s"} in view.`;
  const cards = document.getElementById("cards");
  const tbody = document.getElementById("tbody");
  const empty = document.getElementById("empty");
  if (!rows.length) { cards.innerHTML=""; tbody.innerHTML=""; empty.hidden=false; return; }
  empty.hidden = true;
  cards.innerHTML = rows.map(p => {
    const meta = p.group || "—";
    return `
    <article class="card">
      <a class="card-stretch" href="${hrefOf(p)}" aria-label="${esc(p.product)}"></a>
      <div class="card-top"><div><div class="card-title">${esc(p.product)}</div></div><span class="card-badges">${typeBadge(p)}<span class="badge ${badgeClass(p.stage)}">${esc(p.stage)}</span></span></div>
      <div class="meta">${esc(meta)}</div>
      <div class="meta">${esc(p.due || "Unscheduled")}${p.technology ? " · " + esc(p.technology) : ""}${p.pddf ? " · " + esc(pddfLabel(p.pddf)) : ""}</div>
      ${featTags(p)}
      ${p.description ? `<p class="card-desc ka" title="${esc(p.description)}">${esc(p.description)}</p>` : ""}
      <div class="owners">${ownerChip(p.cpo,"CPO")}${ownerChip(p.tpo,"TPO")}</div>
    </article>`;
  }).join("");
  tbody.innerHTML = rows.map(p => {
    return `
    <tr>
      <td><a href="${hrefOf(p)}"><strong>${esc(p.product)}</strong></a></td>
      <td>${typeBadge(p) || "—"}</td>
      <td class="desc-cell">${p.description ? `<span class="ka desc-trunc" title="${esc(p.description)}">${esc(p.description)}</span>` : "—"}</td>
      <td>${esc(p.group || "—")}</td>
      <td><span class="badge ${badgeClass(p.stage)}">${esc(p.stage)}</span></td>
      <td>${esc(p.due || "Unscheduled")}</td>
      <td>${esc(p.cpo || "Unassigned")}</td>
      <td>${esc(p.tpo || "Unassigned")}</td>
      <td>${esc(p.technology || "—")}</td>
      <td>${esc((p.features || []).join(", ") || "—")}</td>
      <td>${esc(pddfLabel(p.pddf))}</td>
    </tr>`;
  }).join("");
  syncUrl();
}
function bind() {
  const dues = [...new Set(PRODUCTS.map(p => p.due).filter(Boolean))].sort();
  const groups = [...new Set(PRODUCTS.map(p => p.group).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
  const types = [...new Set(PRODUCTS.map(p => p.type).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
  const unType = PRODUCTS.some(p => !p.type);
  function countBy(key) { const m = {}; PRODUCTS.forEach(p => { if (p[key]) m[p[key]] = (m[p[key]] || 0) + 1; }); return m; }
  const cpoCount = countBy("cpo");
  const cpos = Object.keys(cpoCount).sort((a,b) => (cpoCount[b]-cpoCount[a]) || a.localeCompare(b));
  const tpos = Object.keys(countBy("tpo")).sort((a,b)=>a.localeCompare(b));
  const unCpo = PRODUCTS.some(p => !p.cpo) || state.cpo === "__unassigned__";
  const unTpo = PRODUCTS.some(p => !p.tpo) || state.tpo === "__unassigned__";
  const pddfCodes = [...new Set(PRODUCTS.map(p => p.pddf).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
  function renderPddf() {
    const el = document.getElementById("pddf-filters");
    if (!el) return;
    let html = `<button type="button" class="chip ${state.pddf===""?"active":""}" data-key="pddf" data-val="" aria-pressed="${state.pddf===""}">All</button>`;
    for (const v of pddfCodes) html += `<button type="button" class="chip ${state.pddf===v?"active":""}" data-key="pddf" data-val="${esc(v)}" aria-pressed="${state.pddf===v}">${esc(pddfLabel(v))}</button>`;
    if (PRODUCTS.some(p => !p.pddf)) html += `<button type="button" class="chip ${state.pddf==="__notset__"?"active":""}" data-key="pddf" data-val="__notset__" aria-pressed="${state.pddf==="__notset__"}">Not set</button>`;
    el.querySelectorAll("button").forEach(b => b.remove());
    el.insertAdjacentHTML("beforeend", html);
  }
  function paintAll() {
    renderChips("stage-filters", stagePresent, "stage", "All");
    renderChips("type-filters", types, "type", "All", unType ? [["Unspecified", "Unspecified"]] : []);
    renderChips("due-filters", dues, "due", "All", [["Unscheduled", "Unscheduled"]]);
    renderChips("group-filters", groups, "group", "All");
    renderChips("cpo-filters", cpos, "cpo", "All", unCpo ? [["Unassigned", "__unassigned__"]] : []);
    renderChips("tpo-filters", tpos, "tpo", "All", unTpo ? [["Unassigned", "__unassigned__"]] : []);
    renderPddf();
  }
  applyUrl();
  const qEl = document.getElementById("q");
  if (qEl && state.q) qEl.value = state.q;
  const stagePresent = STAGES.filter(s => PRODUCTS.some(p => p.stage === s));
  paintAll();
  document.getElementById("filters").addEventListener("click", e => {
    const btn = e.target.closest("button[data-key]");
    if (!btn) return;
    const key = btn.getAttribute("data-key");
    state[key] = btn.getAttribute("data-val") || "";
    paintAll();
    render();
  });
  document.getElementById("q").addEventListener("input", e => { state.q = e.target.value.trim().toLowerCase(); render(); });
  render();
}
bind();
</script>
"""


def projects_catalog_js(rows: list[dict]) -> str:
    return (
        PROJECTS_JS
        .replace("__PRODUCTS__", json.dumps(rows, ensure_ascii=False))
        .replace("__STAGES__", json.dumps(STAGES))
        .replace("__PDDF_LABELS__", json.dumps(PDDF_PHASE_LABELS, ensure_ascii=False))
    )


def build_internal_projects(projects: list[dict], products: list[dict], updated: str) -> str:
    n = len(projects)
    lede = (
        "Internal tools, portal work and platform projects from the second table on the ALL sheet. "
        "These are not commercial products."
        if n else
        "No internal projects in the register yet. New rows added under the second table on the ALL sheet will appear here."
    )
    body = f"""
      <div style="display:flex;flex-wrap:wrap;gap:1rem;align-items:flex-end;justify-content:space-between">
        <div>
          <p class="eyebrow">Internal</p>
          <h1>Internal projects</h1>
          <p class="lede">{esc(lede)}</p>
        </div>
        <label class="search-box"><span aria-hidden="true">⌕</span>
          <input id="q" type="search" placeholder="Search projects, descriptions, owners" autocomplete="off"/>
        </label>
      </div>
      <div class="filters" id="filters">
        <div class="filter-row" id="stage-filters" role="group" aria-label="Stage"><span class="filter-label">Stage</span></div>
        <div class="filter-row" id="type-filters" role="group" aria-label="Type"><span class="filter-label">Type</span></div>
        <div class="filter-row" id="due-filters" role="group" aria-label="Due date"><span class="filter-label">Due date</span></div>
        <div class="filter-row" id="group-filters" role="group" aria-label="Category"><span class="filter-label">Category</span></div>
        <div class="filter-row" id="cpo-filters" role="group" aria-label="Commercial owner"><span class="filter-label">Commercial owner</span></div>
        <div class="filter-row" id="tpo-filters" role="group" aria-label="Technical owner"><span class="filter-label">Technical owner</span></div>
        <div class="filter-row" id="pddf-filters" role="group" aria-label="Phase"><span class="filter-label">Phase</span></div>
      </div>
""" + view_toggle_bar() + f"""
      <div class="cards view-cards-target" id="cards"></div>
      <div class="table-wrap view-table-target"><table>
        <thead><tr>{PROJECT_TABLE_HEAD}</tr></thead>
        <tbody id="tbody"></tbody>
      </table></div>
      <div class="empty" id="empty" hidden>No projects match these filters.</div>
"""
    return shell("internal-projects", "Internal projects", body, products, updated, extra_js=projects_catalog_js(projects))


def build_project_page(p: dict, products: list[dict], updated: str) -> str:
    if p.get("description"):
        desc_panel = (
            '<div class="panel desc-panel"><h2>Description</h2>'
            f'<p class="ka">{esc(p["description"])}</p></div>'
        )
    else:
        desc_panel = (
            '<div class="panel desc-panel"><h2>Description</h2>'
            '<p class="muted" style="margin:.5rem 0 0">No description in the register yet.</p></div>'
        )
    phase_dd = f'<dt>Phase</dt><dd>{pddf_html(p)}</dd>' if p.get("pddf") or p.get("pddfUrl") else ""
    body = f"""
      <p class="eyebrow"><a class="muted" href="internal-projects.html">Internal projects</a> / {esc(p['product'])}</p>
      <div style="display:flex;flex-wrap:wrap;gap:1rem;align-items:flex-start;justify-content:space-between">
        <div>
          <h1>{esc(p['product'])}</h1>
          <p class="lede">{esc(' · '.join(b for b in (p.get('type'), p.get('group')) if b) or 'Internal project')}</p>
          {feature_tags_html(p)}
        </div>
        <span class="card-badges">{type_badge_html(p)}<span class="badge {badge_class(p['stage'])}" style="font-size:.85rem;padding:.4rem .75rem">{esc(p['stage'])}</span></span>
      </div>
      {desc_panel}
      <div class="detail-grid">
        <div class="panel">
          <h2>Register details</h2>
          <dl class="dl" style="margin-top:1rem">
            <dt>Type</dt><dd>{esc(p.get('type') or '—')}</dd>
            <dt>Category</dt><dd>{esc(p.get('group') or '—')}</dd>
            <dt>Stage</dt><dd>{esc(p['stage'] or '—')}</dd>
            <dt>Due</dt><dd>{esc(p['due'] or 'Unscheduled')}</dd>
            <dt>Technology</dt><dd>{esc(p['technology'] or '—')}</dd>
            {phase_dd}
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
    return shell("internal-projects", p["product"], body, products, updated)


def build_glossary(products: list[dict], updated: str) -> str:
    items = "".join(f"<dt>{esc(k)}</dt><dd>{esc(v)}</dd>" for k, v in GLOSSARY)
    body = f"""
      <p class="eyebrow">2026 register</p>
      <h1>Legend</h1>
      <p class="lede">Portfolio language used in the register. Head product is the commercial environment; CPO and TPO are people.</p>
      <dl class="glossary">{items}</dl>
"""
    return shell("glossary", "Legend", body, products, updated)


BRAND = ROOT / "brand"
BRAND_FILES = ("logo-344.png", "logo-516.png", "logo-344.webp", "logo-516.webp",
               "favicon-32.png", "apple-touch-icon.png", "favicon.ico")


def copy_brand_assets() -> None:
    """Copy logo + favicon (generated once by make_brand.py) into docs/ so every build keeps them."""
    import shutil
    missing = [f for f in BRAND_FILES if not (BRAND / f).exists()]
    if missing:
        raise SystemExit(f"Missing brand assets {missing} in {BRAND}; run make_brand.py")
    for f in BRAND_FILES:
        shutil.copy2(BRAND / f, SITE / "assets" / f)
    shutil.copy2(BRAND / "favicon.ico", SITE / "favicon.ico")


def main() -> None:
    global CURRENT_PROJECTS
    products, projects = load_register(EXCEL)
    CURRENT_PROJECTS = projects
    finalize_stages(products + projects)
    tga_rows: list[dict] = []
    if USE_NNI:
        tga_rows = load_tga_nni(EXCEL)
        attach_tga_notes(products, tga_rows)
    updated = data_updated_label()
    OUT_JSON.write_text(json.dumps(products, ensure_ascii=False, indent=2) + "\n")
    OUT_PROJECTS_JSON.write_text(json.dumps(projects, ensure_ascii=False, indent=2) + "\n")
    OUT_TGA_JSON.write_text(json.dumps(tga_rows, ensure_ascii=False, indent=2) + "\n")
    SITE.mkdir(parents=True, exist_ok=True)
    (SITE / "assets").mkdir(exist_ok=True)
    (SITE / "assets" / "site.css").write_text(SHARED_CSS + stage_css() + FILTER_CSS)
    (SITE / "assets" / "filter.js").write_text(FILTER_JS)
    (SITE / "assets" / "sort.js").write_text("window.UGT_STAGE_ORDER = " + json.dumps([x for x in ["Wishlist", "Pipeline", "Pre-Development", "Development", "Deployment", "Pre-Production", "Production", "Terminated"] if x in STAGES] + [x for x in STAGES if x not in ("Wishlist", "Pipeline", "Pre-Development", "Development", "Deployment", "Pre-Production", "Production", "Terminated")]) + ";\n" + SORT_JS)
    (SITE / "assets" / "view.js").write_text(VIEW_JS)
    copy_brand_assets()
    (SITE / "assets" / "resize.js").write_text(RESIZE_JS)
    (SITE / "index.html").write_text(build_overview(products, updated), encoding="utf-8")
    (SITE / "products.html").write_text(build_products(products, updated), encoding="utf-8")
    (SITE / "platforms.html").write_text(build_platforms(products, updated), encoding="utf-8")
    wanted_plat = set()
    for plat in platform_nav_list(products):
        dest = SITE / f"{platform_slug(plat)}.html"
        dest.write_text(build_platform_page(plat, products, updated), encoding="utf-8")
        wanted_plat.add(dest.name)
    for stale in SITE.glob("platform-*.html"):
        if stale.name not in wanted_plat:
            stale.unlink()
    (SITE / "internal-projects.html").write_text(build_internal_projects(projects, products, updated), encoding="utf-8")
    (SITE / "roadmap.html").write_text(build_roadmap(products, updated), encoding="utf-8")
    (SITE / "owners.html").write_text(build_owners_index(products, updated), encoding="utf-8")
    (SITE / "status.html").write_text(build_status_notes(products, tga_rows, updated), encoding="utf-8")
    (SITE / "glossary.html").write_text(build_glossary(products, updated), encoding="utf-8")
    codes = sorted({p["cpo"] for p in products if p["cpo"]} | {p["tpo"] for p in products if p["tpo"]})
    for code in codes:
        (SITE / f"owner-{code}.html").write_text(build_owner_page(code, products, updated), encoding="utf-8")
    wanted_prod = set()
    for rec in products:
        dest = SITE / f"product-{rec['n']}.html"
        dest.write_text(build_product_page(rec, products, updated), encoding="utf-8")
        wanted_prod.add(dest.name)
    for stale in SITE.glob("product-*.html"):
        if stale.name not in wanted_prod:
            stale.unlink()
    wanted_proj = set()
    for rec in projects:
        dest = SITE / f"project-{rec['id']}.html"
        dest.write_text(build_project_page(rec, products, updated), encoding="utf-8")
        wanted_proj.add(dest.name)
    for stale in SITE.glob("project-*.html"):
        if stale.name not in wanted_proj:
            stale.unlink()
    (SITE / "assets" / "data.js").write_text(
        "window.UGT_PRODUCTS = " + json.dumps(products, ensure_ascii=False) + ";\n"
        + "window.UGT_PROJECTS = " + json.dumps(projects, ensure_ascii=False) + ";\n"
        + "window.UGT_TGA_NNI = " + json.dumps(tga_rows, ensure_ascii=False) + ";\n"
        + "window.UGT_UPDATED = " + json.dumps(updated) + ";\n",
        encoding="utf-8",
    )
    nni = f", {len(tga_rows)} NNI rows" if USE_NNI else ", NNI skipped"
    print(f"Built {len(products)} products, {len(projects)} internal projects{nni} → {SITE} (updated: {updated})")

if __name__ == "__main__":
    main()
