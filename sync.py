#!/usr/bin/env python3
"""
Diff helper + rebuild/push for the portfolio site.

OneDrive cannot be reached from this script (MCP only). A future agent run must:
  1) Call MCP get_drive_item (see SYNC_STEPS.md)
  2) Pass the metadata JSON to this script, OR write it to --meta
  3) If changed, DownloadFile the xlsx, then run this script with --apply

Usage:
  python3 sync.py --meta /tmp/drive_item.json --check-only
  python3 sync.py --meta /tmp/drive_item.json --apply
  python3 sync.py --force-apply   # rebuild+push using current xlsx (no meta check)
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
STATE = ROOT / "state.json"
PRODUCTS = ROOT / "products.json"
EXCEL = ROOT / "UGT_Product_Portfolio_2026.xlsx"
REPO = ROOT / "repo"
BUILD = ROOT / "build.py"


def load_json(path: Path):
    if not path.exists():
        return None
    return json.loads(path.read_text())


def summarise_diff(old: list[dict] | None, new: list[dict]) -> str:
    if old is None:
        return f"Initial sync: {len(new)} products"
    old_by = {p["n"]: p for p in old}
    new_by = {p["n"]: p for p in new}
    added = [new_by[n] for n in sorted(set(new_by) - set(old_by))]
    removed = [old_by[n] for n in sorted(set(old_by) - set(new_by))]
    changed = []
    for n in sorted(set(old_by) & set(new_by)):
        a, b = old_by[n], new_by[n]
        fields = []
        for k in ("product", "headProducts", "category", "technology", "stage", "due", "cpo", "tpo", "statusUpdate", "whereUsed"):
            if a.get(k) != b.get(k):
                fields.append(f"{k}: {a.get(k)!r} → {b.get(k)!r}")
        if fields:
            changed.append((n, a.get("product"), fields))
    lines = []
    if added:
        lines.append("Added: " + ", ".join(f"{p['n']}:{p['product']}" for p in added))
    if removed:
        lines.append("Removed: " + ", ".join(f"{p['n']}:{p['product']}" for p in removed))
    if changed:
        lines.append(f"Changed ({len(changed)}):")
        for n, name, fields in changed:
            lines.append(f"  #{n} {name}: " + "; ".join(fields))
    if not lines:
        return "Content unchanged (metadata differed but products identical)"
    return "\n".join(lines)


def update_state(meta: dict) -> None:
    STATE.write_text(
        json.dumps(
            {
                "itemId": meta.get("id") or "01DXQX44UQPCQGPFTZKZBJ3BMECRZ5LYJJ",
                "eTag": meta.get("eTag"),
                "cTag": meta.get("cTag"),
                "lastModifiedDateTime": meta.get("lastModifiedDateTime"),
                "size": meta.get("size"),
                "name": meta.get("name"),
            },
            indent=2,
        )
        + "\n"
    )


def run(cmd: list[str], cwd: Path | None = None) -> None:
    print("+", " ".join(cmd), flush=True)
    subprocess.check_call(cmd, cwd=str(cwd) if cwd else None)


def apply_rebuild_push(summary: str) -> None:
    old = load_json(PRODUCTS)
    run(["python3", str(BUILD)])
    new = load_json(PRODUCTS)
    detail = summarise_diff(old, new)
    print(detail)
    # copy built docs + build.py + products.json + state.json + sync files into repo
    if not REPO.exists():
        raise SystemExit(f"Repo dir missing: {REPO}")
    # docs are built at ROOT/docs; repo should contain docs/ at root for Pages
    import shutil
    docs_dst = REPO / "docs"
    if docs_dst.exists():
        shutil.rmtree(docs_dst)
    shutil.copytree(ROOT / "docs", docs_dst)
    for name in ("build.py", "sync.py", "SYNC_STEPS.md", "products.json", "tga_nni.json", "state.json", "README.md"):
        src = ROOT / name
        if src.exists():
            shutil.copy2(src, REPO / name)
    # .gitignore
    (REPO / ".gitignore").write_text(
        "*.xlsx\n*.png\n__pycache__/\n.DS_Store\ngh_login.log\nsite/\n"
    )
    run(["git", "add", "-A"], cwd=REPO)
    # commit only if changes
    st = subprocess.run(["git", "status", "--porcelain"], cwd=REPO, capture_output=True, text=True)
    if not st.stdout.strip():
        print("NO_GIT_CHANGES")
        return
    msg = "Update portfolio from Excel\n\n" + detail
    run(["git", "commit", "-m", msg], cwd=REPO)
    run(["git", "push", "origin", "HEAD"], cwd=REPO)
    print("PUSHED")
    print("SUMMARY")
    print(summary)
    print(detail)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--meta", type=Path, help="JSON from get_drive_item")
    ap.add_argument("--check-only", action="store_true")
    ap.add_argument("--apply", action="store_true", help="Update state, rebuild, commit, push")
    ap.add_argument("--force-apply", action="store_true")
    args = ap.parse_args()

    if args.force_apply:
        apply_rebuild_push("Forced rebuild")
        return

    if not args.meta:
        print("Provide --meta path with get_drive_item JSON (see SYNC_STEPS.md)", file=sys.stderr)
        sys.exit(2)
    meta = json.loads(args.meta.read_text())
    prev = load_json(STATE) or {}
    etag, ctag = meta.get("eTag"), meta.get("cTag")
    if etag == prev.get("eTag") and ctag == prev.get("cTag"):
        print("NO_CHANGE")
        return
    summary = (
        f"Excel changed: eTag {prev.get('eTag')!r} → {etag!r}; "
        f"lastModified {prev.get('lastModifiedDateTime')} → {meta.get('lastModifiedDateTime')}"
    )
    print(summary)
    if args.check_only:
        print("CHANGED")
        return
    if not args.apply:
        print("Re-run with --apply after DownloadFile refreshes the xlsx", file=sys.stderr)
        sys.exit(3)
    if not EXCEL.exists():
        raise SystemExit(f"Missing Excel at {EXCEL}; DownloadFile first")
    update_state(meta)
    try:
        apply_rebuild_push(summary)
    except BaseException:
        # roll back saved state so the next run retries instead of reporting NO_CHANGE
        STATE.write_text(json.dumps(prev, indent=2) + "\n")
        raise


if __name__ == "__main__":
    main()
