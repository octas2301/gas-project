# -*- coding: utf-8 -*-
"""Yahoo 追加画像 uploadLibImage のローカル試験（Pro手UP回避）。

認証: 出品スプシ ▼設定(Yahooマッピング) B12–B15（Sheets API読取）
画像: tools/yahoo_page_mock/out/lib_ready/

  python tools/yahoo_page_mock/upload_lib_images_api.py
  python tools/yahoo_page_mock/upload_lib_images_api.py --p6c
  python tools/yahoo_page_mock/upload_lib_images_api.py --dry-run
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

import requests

ROOT = Path(__file__).resolve().parents[2]
LIB_READY = Path(__file__).resolve().parent / "out" / "lib_ready"
CAPTION_THEMES = Path(__file__).resolve().parent / "out" / "caption_themes"
SS_VIEW = ROOT / "tools" / "common_ss_view"
LISTING_SSID = "1LIWp0qjgvPaZtjsIBmCGqCEEB7AA00nLmBA7iE1MI28"
MAP_RANGE = "'▼設定(Yahooマッピング)'!B12:B15"
UPLOAD_URL = "https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/uploadLibImage"
TOKEN_URL = "https://auth.login.yahoo.co.jp/yconnect/v2/token"
LIB_PUBLIC = "https://shopping.c.yimg.jp/lib/octas/"

# P6b 固定3枠（互換）
FILES_P6B = [
    "maguro_body_t03_sizzle.jpg",
    "maguro_body_t04_aji.jpg",
    "maguro_body_t11_arrange.jpg",
    "buri_body_t03_sizzle.jpg",
    "buri_body_t04_aji.jpg",
    "buri_body_t11_arrange.jpg",
]


def _files_from_p6c_pack() -> list[str]:
    names: list[str] = []
    seen = set()
    if not CAPTION_THEMES.is_dir():
        return names
    for p in sorted(CAPTION_THEMES.glob("*.json")):
        try:
            doc = json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            continue
        for r in doc.get("ranked") or []:
            fn = str((r or {}).get("libFile") or "").strip()
            if not fn or fn in seen:
                continue
            seen.add(fn)
            names.append(fn)
    return names


def _creds() -> tuple[str, str, str, str]:
    sys.path.insert(0, str(SS_VIEW))
    sys.path.insert(0, str(ROOT / "tools" / "competitor_store"))
    from client import sheets_service  # noqa: E402

    svc = sheets_service(write=False, interactive=False)
    if not svc:
        raise SystemExit("NO_CREDS")
    vals = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=LISTING_SSID, range=MAP_RANGE)
        .execute()
        .get("values")
        or []
    )
    flat = [(r[0] if r else "").strip() for r in vals]
    while len(flat) < 4:
        flat.append("")
    client_id, client_secret, refresh, seller = flat[:4]
    if not all([client_id, client_secret, refresh, seller]):
        raise SystemExit("MAP_CREDS_INCOMPLETE")
    return client_id, client_secret, refresh, seller


def _access_token(client_id: str, client_secret: str, refresh: str) -> str:
    r = requests.post(
        TOKEN_URL,
        data={
            "grant_type": "refresh_token",
            "client_id": client_id,
            "client_secret": client_secret,
            "refresh_token": refresh,
        },
        timeout=60,
    )
    if r.status_code != 200:
        raise SystemExit("TOKEN_FAIL http=%s body=%s" % (r.status_code, r.text[:200]))
    tok = r.json().get("access_token")
    if not tok:
        raise SystemExit("TOKEN_EMPTY")
    return tok


def upload_one(token: str, seller_id: str, path: Path, *, directory: str = "") -> tuple[int, str]:
    url = "%s?seller_id=%s" % (UPLOAD_URL, seller_id)
    headers = {"Authorization": "Bearer %s" % token}
    data = {}
    if directory:
        data["directory"] = directory
    with path.open("rb") as f:
        files = {"file": (path.name, f, "image/jpeg")}
        r = requests.post(url, headers=headers, data=data, files=files, timeout=120)
    return r.status_code, r.text[:500]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument(
        "--p6c",
        action="store_true",
        help="caption_themes の ranked.libFile をUP（P6c本番）",
    )
    ap.add_argument("--directory", default="", help="追加画像ディレクトリ名（空=既定パーツ）")
    ap.add_argument("--sleep", type=float, default=0.8)
    args = ap.parse_args()

    files = _files_from_p6c_pack() if args.p6c else list(FILES_P6B)
    if args.p6c and not files:
        print("NO_P6C_FILES Run: python tools/yahoo_page_mock/build_yahoo_caption_themes.py --prep-lib")
        return 1

    missing = [n for n in files if not (LIB_READY / n).is_file()]
    if missing:
        print("MISSING_FILES", missing)
        print("Run: python tools/yahoo_page_mock/build_yahoo_caption_themes.py --prep-lib")
        return 1

    client_id, client_secret, refresh, seller = _creds()
    print("seller_id=%s mode=%s files=%s" % (seller, "p6c" if args.p6c else "p6b", len(files)))
    if args.dry_run:
        for n in files:
            p = LIB_READY / n
            print("DRY", n, "kb=%.1f" % (p.stat().st_size / 1024))
        return 0

    token = _access_token(client_id, client_secret, refresh)
    print("token_ok")
    ok = 0
    ng = 0
    for name in files:
        path = LIB_READY / name
        code, body = upload_one(token, seller, path, directory=args.directory)
        status = "OK" if code == 200 and "OK" in body else "NG"
        if status == "OK":
            ok += 1
        else:
            ng += 1
        print("%s http=%s file=%s body=%s" % (status, code, name, body.replace("\n", " ")[:160]))
        time.sleep(args.sleep)

    print("--- verify public ---")
    for name in files:
        url = LIB_PUBLIC + name
        try:
            hr = requests.head(url, timeout=30, allow_redirects=True)
            # some CDNs disallow HEAD
            if hr.status_code >= 400:
                hr = requests.get(url, timeout=30, stream=True)
            print("pub http=%s %s" % (hr.status_code, name))
        except Exception as e:
            print("pub ERR %s %s" % (name, e))
        time.sleep(0.3)

    print("DONE ok=%s ng=%s" % (ok, ng))
    return 0 if ng == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
