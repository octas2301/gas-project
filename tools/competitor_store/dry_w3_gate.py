# -*- coding: utf-8 -*-
"""W3 dry: 門は Keepaフル優先。シートの門結果と比較。既定は非書（--log で①ログのみ）。"""
from __future__ import annotations

import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
from apply_keepa_full import RESEARCH_SS, COMPETITOR_SS, T_CAND, T_LOG, as_dicts, append_log, cand_asins, read_all  # noqa: E402
from client import sheets_service  # noqa: E402
from keepa_full import keepa_get_needed, latest_row_for_asin, warehouse_get_needed  # noqa: E402
from schema import PURPOSE_RESEARCH, SHEET_KEEPA_FULL  # noqa: E402

FROZEN = re.compile(r"冷凍|冷蔵|生鮮")
T_PROF = "①プロファイル"


def utc_now() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def stats_slot(p: dict, idx: int):
    st = p.get("stats") if isinstance(p.get("stats"), dict) else {}
    for key in ("avg90", "avg30", "avg"):
        arr = st.get(key)
        if not isinstance(arr, list) or idx >= len(arr):
            continue
        v = arr[idx]
        if v is None or v == -1:
            continue
        try:
            n = float(v)
        except (TypeError, ValueError):
            continue
        if n >= 0:
            return n
    return None


def product_usable(p: dict) -> bool:
    if not p:
        return False
    return stats_slot(p, 18) is not None or stats_slot(p, 1) is not None or stats_slot(p, 3) is not None


def gate(p: dict, price_min: float, rank_max: float) -> tuple[str, str]:
    if not product_usable(p):
        return "落ち", "stats空"
    title = str(p.get("title") or "")
    price = stats_slot(p, 18)
    if price is None:
        price = stats_slot(p, 1)
    rank = stats_slot(p, 3)
    if price is not None and price < price_min:
        return "落ち", "価格<%s" % int(price_min)
    if rank is not None and rank > rank_max:
        return "落ち", "順位>%s" % int(rank_max)
    if FROZEN.search(title):
        return "落ち", "冷凍冷蔵生鮮"
    return "通過", "門通過"


def read_profile(svc) -> tuple[float, float]:
    raw = read_all(svc, RESEARCH_SS, T_PROF)
    price_min, rank_max = 2000.0, 150000.0
    if len(raw) < 2:
        return price_min, rank_max
    h = [str(x) for x in raw[0]]
    row = raw[1]
    m = {h[i]: (row[i] if i < len(row) else "") for i in range(len(h))}
    try:
        p = float(m.get("価格下限") or 0)
        if p > 0:
            price_min = p
    except (TypeError, ValueError):
        pass
    try:
        r = float(m.get("順位段階1") or 0)
        if r > 0:
            rank_max = r
    except (TypeError, ValueError):
        pass
    return price_min, rank_max


def col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def main() -> int:
    log = "--log" in sys.argv
    apply_tr = "--apply-transcript" in sys.argv
    svc = sheets_service(write=True, interactive=False)
    if not svc:
        print("NO_CREDS")
        return 2
    price_min, rank_max = read_profile(svc)
    headers, cand = as_dicts(read_all(svc, RESEARCH_SS, T_CAND))
    _, full = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    asins = cand_asins(svc)
    need_get = 0
    no_row = 0
    unusable = 0
    csv_hit = 0
    match = 0
    mismatch = 0
    empty_old = 0
    pass_n = 0
    drop_n = 0
    samples = []
    changes = []
    for i, c in enumerate(cand):
        a = str(c.get("ASIN") or "").strip().upper()
        if not a:
            continue
        row = latest_row_for_asin(full, a)
        if not row:
            no_row += 1
            need_get += 1
            continue
        if warehouse_get_needed(row):
            need_get += 1
        try:
            p = json.loads(row.get("生JSON") or "{}")
        except json.JSONDecodeError:
            unusable += 1
            need_get += 1
            continue
        if "csv" in p:
            csv_hit += 1
        if not product_usable(p):
            unusable += 1
        st, why = gate(p, price_min, rank_max)
        if st == "通過":
            pass_n += 1
        else:
            drop_n += 1
        old = str(c.get("門結果") or "")
        if not old:
            empty_old += 1
        elif old == st:
            match += 1
        else:
            mismatch += 1
            changes.append({"row": i + 2, "asin": a, "old": old, "st": st, "why": why})
            if len(samples) < 8:
                samples.append("%s old=%s new=%s %s" % (a, old, st, why))
    print(
        "cand=%d full_join=%d need_get=%d no_row=%d unusable=%d csv=%d"
        % (len(asins), len(asins) - no_row, need_get, no_row, unusable, csv_hit)
    )
    print(
        "gate pass=%d drop=%d vs①候補 match=%d mismatch=%d empty_old=%d priceMin=%s rankMax=%s"
        % (pass_n, drop_n, match, mismatch, empty_old, int(price_min), int(rank_max))
    )
    for s in samples:
        print("mismatch", s)
    # W3a/b 合否: 倉庫が読める・csvなし・90日GET不要が主。門の新旧差は転記前の観測。
    ok_store = no_row == 0 and csv_hit == 0
    ok_skip = no_row == 0  # 日付上の GET。unusable は別カウント
    print("W3a_store", "PASS" if ok_store else "FAIL")
    print("W3b_fresh_rows", "PASS" if ok_skip else "FAIL")
    print("transcript_n", len(changes))
    if apply_tr and changes:
        idx = {h: i for i, h in enumerate(headers)}
        data = []
        for ch in changes:
            r = ch["row"]
            for name, val in (("門結果", ch["st"]), ("門理由", ch["why"]), ("runId", "pr_20260815_w3tr")):
                col = idx.get(name)
                if col is None:
                    continue
                data.append(
                    {
                        "range": "'%s'!%s%d" % (T_CAND.replace("'", "''"), col_letter(col + 1), r),
                        "values": [[val]],
                    }
                )
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=RESEARCH_SS,
            body={"valueInputOption": "RAW", "data": data},
        ).execute()
        print("transcript_wrote", len(changes))
    if log or apply_tr:
        append_log(
            svc,
            "W3",
            "runId=pr_20260815_w3tr match=%d mismatch=%d pass=%d drop=%d unusable=%d wrote=%s"
            % (match, mismatch, pass_n, drop_n, unusable, len(changes) if apply_tr else 0),
        )
        print("logged")
    return 0 if ok_store else 1


if __name__ == "__main__":
    raise SystemExit(main())
