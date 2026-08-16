# -*- coding: utf-8 -*-
"""S4b: セラー構成比を Amazon カテゴリ名列＋％（分析②）。メイン1セルも書く。seller GETしない。"""
from __future__ import annotations

import gzip
import json
import re
import sys
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import Request, urlopen

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT.parent / "purchase_research_path3"))

from apply_keepa_full import COMPETITOR_SS, append_log, as_dicts, read_all, sheets_service
from dry_offer_count import col_letter
from init_store import ensure_seller_google
from keepa_csv_vs_api import keepa_key
from schema import SHEET_SELLER

CAT_API = "https://api.keepa.com/category"
COL_PREFIX = "構成%"  # 列名衝突回避。中身は数値％


def parse_mix(raw: str) -> list[tuple[str, float]]:
    out = []
    for part in str(raw or "").split("|"):
        part = part.strip()
        if not part:
            continue
        bits = part.split(":")
        if len(bits) < 3:
            continue
        cid = bits[0].strip()
        pct_s = bits[2].replace("%", "").strip()
        try:
            pct = float(pct_s)
        except ValueError:
            continue
        if cid:
            out.append((cid, pct))
    return out


def cat_url(ids: list[str], key: str) -> str:
    q = urlencode({"key": key, "domain": "5", "category": ",".join(ids)})
    return CAT_API + "?" + q


def decode_body(raw: bytes) -> dict:
    if raw[:2] == b"\x1f\x8b":
        raw = gzip.decompress(raw)
    return json.loads(raw.decode("utf-8"))


def names_from_body(body: dict) -> dict[str, str]:
    out = {}
    cats = body.get("categories") or body.get("category") or {}
    if isinstance(cats, dict):
        for k, v in cats.items():
            if isinstance(v, dict):
                name = str(v.get("name") or v.get("contextFreeName") or "").strip()
                if name:
                    out[str(k)] = name
            elif v:
                out[str(k)] = str(v).strip()
    if isinstance(cats, list):
        for v in cats:
            if not isinstance(v, dict):
                continue
            cid = str(v.get("catId") or v.get("categoryId") or v.get("id") or "").strip()
            name = str(v.get("name") or v.get("contextFreeName") or "").strip()
            if cid and name:
                out[cid] = name
    return out


def col_title(name: str, cid: str) -> str:
    n = re.sub(r"[\r\n]+", " ", name or "").strip() or cid
    if len(n) > 40:
        n = n[:40]
    return "%s %s" % (COL_PREFIX, n)


def main() -> int:
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    args = ap.parse_args()
    svc = sheets_service(write=True)
    h, rows = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_SELLER))
    rec = rows[0] if rows else {}
    mix = parse_mix(rec.get("カテゴリ構成") or "")
    ids = [c for c, _p in mix]
    ok = len(mix) >= 1 and all(re.fullmatch(r"\d+", c) for c, _p in mix)
    line = "runId=pr_20260816_s4bdry n_cat=%d sellerGETなし %s" % (len(mix), "PASS" if ok else "FAIL")
    print(line)
    print("ids", ids)
    if not args.apply:
        append_log(svc, "S4b", line)
        return 0 if ok else 1
    if not ok:
        return 1
    key = keepa_key()
    if not key:
        append_log(svc, "S4b", "runId=pr_20260816_s4bcol FAIL no_key")
        return 2
    req = Request(cat_url(ids, key), method="GET")
    with urlopen(req, timeout=60) as resp:
        body = decode_body(resp.read())
    names = names_from_body(body)
    consumed = body.get("tokensConsumed")
    left = body.get("tokensLeft")
    named = sum(1 for c, _p in mix if names.get(c))
    print("named", named, "/", len(mix), "consumed", consumed, "left", left)
    if named < 1:
        print("keys", list(body.keys())[:12])
        append_log(svc, "S4b", "runId=pr_20260816_s4bcol FAIL no_names consumed=%s" % consumed)
        return 1
    main_cid, main_pct = mix[0]
    main_name = names.get(main_cid) or main_cid
    main_label = "%s %.1f%%" % (main_name, main_pct)
    print("ensure", ensure_seller_google(COMPETITOR_SS))
    h2, _ = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_SELLER))
    want_cols = []
    for cid, pct in mix:
        title = col_title(names.get(cid) or cid, cid)
        want_cols.append((title, pct))
    missing = [t for t, _p in want_cols if t not in h2]
    if "メインカテゴリ" not in h2:
        missing = ["メインカテゴリ"] + missing
    if missing:
        start = len(h2) + 1
        rng = "'%s'!%s1" % (SHEET_SELLER.replace("'", "''"), col_letter(start))
        svc.spreadsheets().values().update(
            spreadsheetId=COMPETITOR_SS,
            range=rng,
            valueInputOption="RAW",
            body={"values": [missing]},
        ).execute()
        h2, _ = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_SELLER))
    data = []
    if "メインカテゴリ" in h2:
        data.append(
            {
                "range": "'%s'!%s2" % (SHEET_SELLER.replace("'", "''"), col_letter(h2.index("メインカテゴリ") + 1)),
                "values": [[main_label]],
            }
        )
    for title, pct in want_cols:
        if title not in h2:
            continue
        data.append(
            {
                "range": "'%s'!%s2" % (SHEET_SELLER.replace("'", "''"), col_letter(h2.index(title) + 1)),
                "values": [[pct]],
            }
        )
    svc.spreadsheets().values().batchUpdate(
        spreadsheetId=COMPETITOR_SS,
        body={"valueInputOption": "USER_ENTERED", "data": data},
    ).execute()
    h3, rows3 = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_SELLER))
    r = rows3[0]
    n_pct = sum(1 for t, _p in want_cols if str(r.get(t) or "").strip())
    vok = str(r.get("メインカテゴリ") or "").strip() and n_pct == len(want_cols)
    line2 = "runId=pr_20260816_s4bcol named=%d pct_cols=%d consumed=%s sellerGETなし %s" % (
        named,
        n_pct,
        consumed,
        "PASS" if vok else "FAIL",
    )
    append_log(svc, "S4b", line2)
    print(line2)
    print("main", main_label)
    print("VERIFY", "PASS" if vok else "FAIL")
    return 0 if vok else 1


if __name__ == "__main__":
    raise SystemExit(main())
