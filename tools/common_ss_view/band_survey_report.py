# -*- coding: utf-8 -*-
"""規模帯・調査カテゴリ落としを①ログから集計する（手元調査用）。

使い方:
  python tools/common_ss_view/band_survey_report.py
  python tools/common_ss_view/band_survey_report.py --tail 4000

出力: tools/common_ss_view/out/band_survey_report.txt
"""
from __future__ import annotations

import argparse
import re
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT.parent / "competitor_store"))

from client import sheets_service  # noqa: E402

SID_R = "1tf7gvkD88yyNz7JWXfNysBcIqSZDlI9dC-l6gOPyLjE"
SID_C = "1UrdWDBw8NcuOf71Bi-2m8WNQDW2onIkA-zl6mLE7AHs"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--tail", type=int, default=4000)
    args = ap.parse_args()

    svc = sheets_service(write=False, interactive=False)
    meta = svc.spreadsheets().get(spreadsheetId=SID_R, fields="sheets.properties").execute()
    rc = 1000
    for s in meta.get("sheets") or []:
        if s["properties"]["title"] == "①ログ":
            rc = int((s["properties"].get("gridProperties") or {}).get("rowCount") or 1000)
            break
    start = max(1, rc - args.tail)
    vals = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=SID_R, range="'①ログ'!A%d:H%d" % (start, rc))
        .execute()
        .get("values")
        or []
    )

    survey_n = 0
    allow_c = Counter()
    drop_c = Counter()
    keep_sum = drop_sum = miss_sum = 0
    keep_pos = 0
    samples = []
    why = Counter()
    sm_ok = 0

    for r in vals:
        if len(r) < 2:
            continue
        tag = str(r[1] or "")
        detail = str(r[7] if len(r) > 7 else "")
        at = str(r[3] if len(r) > 3 else "")
        parent = str(r[5] if len(r) > 5 else "")
        if tag == "survey_diag":
            survey_n += 1
            am = re.search(r"allow=([^\s]+)", detail)
            if am:
                allow_c[am.group(1)] += 1
            km = re.search(r"keep=(\d+)", detail)
            dm = re.search(r"drop=(\d+)", detail)
            mm = re.search(r"miss=(\d+)", detail)
            k = int(km.group(1)) if km else 0
            d = int(dm.group(1)) if dm else 0
            m = int(mm.group(1)) if mm else 0
            keep_sum += k
            drop_sum += d
            miss_sum += m
            if k > 0:
                keep_pos += 1
            dt = re.search(r"dropTop=([^\s]+)", detail)
            if dt:
                for part in dt.group(1).split("|"):
                    if "=" in part:
                        key, n = part.rsplit("=", 1)
                        try:
                            drop_c[key] += int(n)
                        except ValueError:
                            drop_c[key] += 1
            samples.append("%s | %s | %s" % (at, parent, detail[:200]))
        elif tag == "enum_diag_sum":
            wm = re.search(r"why=([^\s]+)", detail)
            if wm:
                for part in wm.group(1).split("|"):
                    if "=" in part:
                        k2, v = part.split("=", 1)
                        try:
                            why[k2] += int(v)
                        except ValueError:
                            pass
        elif tag == "enum_diag":
            if "why=sm_queue_ok" in detail or re.search(r"queue=([1-9]\d*)", detail):
                if "why=sm_queue_ok" in detail:
                    sm_ok += 1

    # メーカーサマリ規模帯
    body = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=SID_C, range="'メーカーサマリ'")
        .execute()
        .get("values")
        or []
    )
    hdr = [str(x) for x in (body[0] if body else [])]
    idx_band = next((i for i, h in enumerate(hdr) if "規模帯" in h), -1)
    idx_est = next((i for i, h in enumerate(hdr) if "Catalog見積" in h), -1)
    band_c = Counter()
    est_blank = 0
    for row in body[1:]:
        b = str(row[idx_band] if idx_band >= 0 and idx_band < len(row) else "").strip() or "(空)"
        band_c[b] += 1
        if idx_est >= 0:
            ev = row[idx_est] if idx_est < len(row) else ""
            if ev is None or str(ev).strip() == "":
                est_blank += 1

    lines = [
        "# 規模帯・調査カテゴリ報告",
        "",
        "## カテゴリ一致（survey_diag）",
        "件数=%s" % survey_n,
        "ロジック許可(allow)の種類=%s" % dict(allow_c.most_common(5)),
        "keep合計=%s / drop合計=%s / miss合計=%s" % (keep_sum, drop_sum, miss_sum),
        "keep>0 だった社ログ=%s" % keep_pos,
        "",
        "### Amazon側で落ちた大分類トップ（dropTop集計）",
    ]
    for k, n in drop_c.most_common(20):
        lines.append("- %s … %s件" % (k, n))
    lines.extend(
        [
            "",
            "### 直近 survey_diag 例",
            *(samples[-12:] or ["(なし)"]),
            "",
            "## 列挙 why（enum_diag_sum）",
            str(dict(why)),
            "sm_queue_ok行=%s" % sm_ok,
            "",
            "## メーカーサマリ規模帯",
            str(dict(band_c.most_common(10))),
            "Catalog見積空白=%s / 行=%s" % (est_blank, max(0, len(body) - 1)),
            "",
            "## 読み方",
            "- allow= … ①プロファイルで調べる大分類（いまは食品ルート57239051が多い）",
            "- dropTop= … 検索には出たが allow と一致せず落ちた Amazon 大分類",
            "- keep=0 が続く＝規模帯の材料（キュー）が付かない",
            "",
            "## 仮想合否",
            "ログ調査可能=%s" % ("YES" if survey_n > 0 else "NO"),
            "カテゴリ一致で残る=%s" % ("YES" if keep_pos > 0 else "NO（allow外ばかり）"),
            "洗える小中キュー=%s" % ("YES" if sm_ok > 0 else "NO"),
        ]
    )
    out = ROOT / "out" / "band_survey_report.txt"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text("\n".join(lines), encoding="utf-8")
    print("\n".join(lines))
    print("wrote", out.as_posix())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
