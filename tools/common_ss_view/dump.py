# -*- coding: utf-8 -*-
"""共通スプシから関係タブだけをローカルへ出す。全タブ常時読みはしない。"""
from __future__ import annotations

import argparse
import csv
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT.parent / "competitor_store"))

from registry import (  # noqa: E402
    OUT_DIR,
    files_for_menu,
    load_registry,
    resolve_file,
    resolve_intent,
    resolve_menu,
    save_registry,
    tabs_for_intent,
)


def _svc():
    from client import sheets_service

    svc = sheets_service(write=False, interactive=False)
    if not svc:
        raise SystemExit("NO_CREDS: tools/c1_hpc_packaged/secrets の token を確認")
    return svc


def _titles(svc, sid: str) -> list[str]:
    meta = svc.spreadsheets().get(spreadsheetId=sid, fields="sheets.properties.title").execute()
    return [s["properties"]["title"] for s in meta.get("sheets") or []]


def _read_values(svc, sid: str, title: str) -> list[list[str]]:
    rng = "'%s'" % title.replace("'", "''")
    return svc.spreadsheets().values().get(spreadsheetId=sid, range=rng).execute().get("values") or []


def _slice_rows(rows: list[list[str]], spec: dict, run_id: str, maker: str) -> list[list[str]]:
    if not rows:
        return rows
    role = str(spec.get("role") or "")
    data_start = int(spec.get("data_start") or 2)
    header = rows[:1]
    body = rows[1:]
    if role == "log" and run_id:
        hit = [r for r in body if any(run_id in str(c) for c in r)]
        if hit:
            return header + hit
        tail = int(spec.get("tail") or 40)
        return header + body[-tail:]
    if maker and role in ("lookup", "result"):
        hit = [r for r in body if any(maker in str(c) for c in r)]
        if hit:
            return header + hit[:80]
    if role == "log":
        tail = int(spec.get("tail") or 40)
        return header + body[-tail:]
    if role == "keepa":
        tail = int(spec.get("tail") or 20)
        return header + body[-tail:]
    head_n = int(spec.get("head") or 40)
    start = max(0, data_start - 2)
    return header + body[start : start + head_n]


def _safe_name(s: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", s)


def _write_tsv(path: Path, rows: list[list[str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f, delimiter="\t", lineterminator="\n")
        for r in rows:
            w.writerow([str(c) for c in r])


def cmd_dump(args: argparse.Namespace) -> int:
    reg = load_registry()
    menu = resolve_menu(args.menu, reg)
    intent = resolve_intent(args.intent, reg)
    file_keys = files_for_menu(menu, reg)
    if args.with_competitor and "統合情報DB260815_競合（amz/楽/Yah）" not in file_keys:
        file_keys.append("統合情報DB260815_競合（amz/楽/Yah）")
    extra_tabs = [t for t in (args.sheet or []) if t]
    svc = _svc()
    seen = ["見た場所:"]
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    for key in file_keys:
        sid = str(reg["files"][key]["id"])
        titles = _titles(svc, sid)
        if args.all_tabs:
            want = list(titles)
        else:
            want = tabs_for_intent(intent, key, reg)
            for t in extra_tabs:
                if t not in want:
                    want.append(t)
        dumped = []
        missing = []
        for title in want:
            if title not in titles:
                missing.append(title)
                continue
            spec = (reg["files"][key].get("tabs") or {}).get(title) or {"role": "result", "head": 40}
            if args.no_limit:
                rows = _read_values(svc, sid, title)
            else:
                rows = _slice_rows(_read_values(svc, sid, title), spec, args.run_id, args.maker)
            out_path = OUT_DIR / _safe_name(key) / (_safe_name(title) + ".tsv")
            _write_tsv(out_path, rows)
            dumped.append("%s (%d行) -> %s" % (title, len(rows), out_path.as_posix()))
        line = "- ファイル: %s／タブ: %s" % (key, "、".join(dumped) if dumped else "（該当タブなし）")
        if missing:
            line += "／未存在: " + "、".join(missing)
        seen.append(line)
        print("titles[%s] %s" % (key, " | ".join(titles[:30])))

    log_note = "- ログ: "
    if args.run_id:
        log_note += args.run_id
    else:
        log_note += "直近行（--run-id 未指定）"
    if args.maker:
        log_note += "／メーカー=%s" % args.maker
    seen.append(log_note)
    seen_text = "\n".join(seen) + "\n"
    (OUT_DIR / "seen.txt").write_text(seen_text, encoding="utf-8")
    print(seen_text)
    return 0


def cmd_list(_args: argparse.Namespace) -> int:
    reg = load_registry()
    svc = _svc()
    for key, rec in reg["files"].items():
        sid = rec["id"]
        titles = _titles(svc, sid)
        print("=== %s ===" % key)
        print("id", sid)
        print("tabs", " | ".join(titles))
    return 0


def cmd_add_sheet(args: argparse.Namespace) -> int:
    """このシートも見て、の恒久追加。registry.json と要件の表を揃える。"""
    reg = load_registry()
    key, _sid = resolve_file(args.file, reg)
    tabs = reg["files"][key].setdefault("tabs", {})
    tabs[args.sheet] = {"role": args.role, "head": 40, "tail": 40}
    intent = resolve_intent(args.intent, reg)
    by_file = reg["intents"].setdefault(intent, {})
    lst = by_file.setdefault(key, [])
    if args.sheet not in lst:
        lst.append(args.sheet)
    save_registry(reg)
    print("added", key, args.sheet, "intent", intent)
    print("docs/org/COMMON_SS_VIEW_REQUIREMENTS.md §5 も同じタブを追記すること")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("dump", help="関係タブだけ out/ へ")
    p.add_argument("--menu", default="仕入れ検討①")
    p.add_argument("--intent", default="judge")
    p.add_argument("--run-id", default="")
    p.add_argument("--maker", default="")
    p.add_argument("--sheet", action="append", default=[], help="追加タブ（今回だけ）")
    p.add_argument("--all-tabs", action="store_true")
    p.add_argument("--no-limit", action="store_true")
    p.add_argument("--with-competitor", action="store_true", help="出品メニューに競合DBを足す")
    p.set_defaults(func=cmd_dump)

    p = sub.add_parser("list", help="登録ファイルのタブ名一覧")
    p.set_defaults(func=cmd_list)

    p = sub.add_parser("add-sheet", help="見るタブを registry に恒久追加")
    p.add_argument("--file", required=True)
    p.add_argument("--sheet", required=True)
    p.add_argument("--intent", default="judge")
    p.add_argument("--role", default="result")
    p.set_defaults(func=cmd_add_sheet)

    args = ap.parse_args()
    return int(args.func(args))


if __name__ == "__main__":
    raise SystemExit(main())
