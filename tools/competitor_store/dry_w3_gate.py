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
from keepa_full import _cat_root_tree, keepa_get_needed, latest_row_for_asin, warehouse_get_needed  # noqa: E402
from schema import AMAZON_SELLER_CAT_COLS, PURPOSE_RESEARCH, SHEET_KEEPA_FULL  # noqa: E402

FROZEN = re.compile(r"冷凍|冷蔵|生鮮")
T_PROF = "①プロファイル"

PROF_HEADERS = [
    "カテゴリ",
    "マッチ語",
    "価格下限",
    "順位段階1",
    "順位段階2",
    "順位段階3",
    "門に使う順位段階",
    "本体ルール",
    "本体不在%",
    "冷凍生",
    "大型",
    "出品者1人",
    "粗利",
    "有効",
    "メモ",
]


def _num(v, default=0.0) -> float:
    try:
        n = float(v)
        return n if n > 0 else default
    except (TypeError, ValueError):
        return default


def default_profile_rows() -> list[list[str]]:
    common = [
        "共通",
        "",
        "2000",
        "150000",
        "",
        "",
        "順位段階1",
        "6ヶ月不在なら通過",
        "80",
        "落とす",
        "落とさない",
        "フラグのみ",
        "使わない",
        "はい",
        "未設定カテゴリはここ。コード定数に固定しない",
    ]
    food = [
        "食品・飲料・お酒",
        "",
        "2000",
        "150000",
        "60000",
        "30000",
        "順位段階1",
        "6ヶ月不在なら通過",
        "80",
        "落とす",
        "落とさない",
        "フラグのみ",
        "使わない",
        "はい",
        "段階2/3は賞味期限向け。門の落ちは「門に使う順位段階」列",
    ]
    food_alias = [
        "食品",
        "",
        "2000",
        "150000",
        "60000",
        "30000",
        "順位段階1",
        "6ヶ月不在なら通過",
        "80",
        "落とす",
        "落とさない",
        "フラグのみ",
        "使わない",
        "はい",
        "旧行名。ルート名は食品・飲料・お酒",
    ]
    snack = [
        "おかし",
        "菓子|スイーツ|チョコ",
        "1000",
        "60000",
        "",
        "",
        "順位段階1",
        "6ヶ月不在なら通過",
        "80",
        "落とす",
        "落とさない",
        "フラグのみ",
        "使わない",
        "はい",
        "DOMAIN1 仮置き。人が直してよい",
    ]
    rows = [PROF_HEADERS, common, food, food_alias, snack]
    skip = {"食品・飲料・お酒"}
    for name in AMAZON_SELLER_CAT_COLS:
        if name in skip:
            continue
        rows.append(
            [
                name,
                "",
                "2000",
                "150000",
                "",
                "",
                "順位段階1",
                "6ヶ月不在なら通過",
                "80",
                "落とす",
                "落とさない",
                "フラグのみ",
                "使わない",
                "はい",
                "初期は共通と同じ。カテゴリごとに人が変える",
            ]
        )
    return rows


def load_profiles(svc) -> list[dict]:
    raw = read_all(svc, RESEARCH_SS, T_PROF)
    if len(raw) < 2:
        return []
    h = [str(x) for x in raw[0]]
    out = []
    for row in raw[1:]:
        m = {h[i]: (row[i] if i < len(row) else "") for i in range(len(h))}
        if str(m.get("有効") or "はい").strip() in ("いいえ", "no", "NO", "0"):
            continue
        out.append(m)
    return out


def profile_for_product(p: dict, profs: list[dict]) -> dict:
    root, joined = _cat_root_tree(p or {})
    blob = (root + " " + joined).strip()
    common = None
    for m in profs:
        cat = str(m.get("カテゴリ") or "").strip()
        if cat == "共通":
            common = m
            continue
        needles = [x for x in str(m.get("マッチ語") or "").split("|") if x.strip()]
        if needles and any(n in blob for n in needles):
            return m
        if cat and cat == root:
            return m
        if cat == "食品" and root.startswith("食品"):
            return m
    return common or (profs[0] if profs else {})


def stage_to_col(raw) -> str:
    s = str(raw if raw is not None else "1").strip()
    if s.endswith(".0") and s[:-2].replace("-", "", 1).isdigit():
        s = s[:-2]
    return {
        "1": "順位段階1",
        "2": "順位段階2",
        "3": "順位段階3",
        "順位段階1": "順位段階1",
        "順位段階2": "順位段階2",
        "順位段階3": "順位段階3",
    }.get(s, "順位段階1")


def rank_limit(prof: dict) -> float:
    key = stage_to_col(prof.get("門に使う順位段階"))
    n = _num(prof.get(key), 0)
    if n <= 0:
        n = _num(prof.get("順位段階1"), 150000)
    return n if n > 0 else 150000.0


def amazon_oos180(p: dict):
    st = p.get("stats") if isinstance(p.get("stats"), dict) else {}
    arr = st.get("outOfStockPercentage180")
    if not isinstance(arr, list) or not arr:
        return None
    v = arr[0]
    if v in (None, "", -1):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def gate_product(p: dict, profs: list[dict]) -> tuple[str, str]:
    prof = profile_for_product(p, profs)
    price_min = _num(prof.get("価格下限"), 2000)
    rank_max = rank_limit(prof)
    return gate_with_prof(p, prof, price_min, rank_max)


def gate_with_prof(p: dict, prof: dict, price_min: float, rank_max: float) -> tuple[str, str]:
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
    frozen_on = str(prof.get("冷凍生") or "落とす").strip() != "落とさない"
    if frozen_on and FROZEN.search(title):
        return "落ち", "冷凍冷蔵生鮮"
    rule = str(prof.get("本体ルール") or "").strip()
    if rule == "6ヶ月不在なら通過":
        av = p.get("availabilityAmazon")
        if av in (0, "0"):
            oos = amazon_oos180(p)
            thresh = _num(prof.get("本体不在%"), 80)
            if oos is None:
                pass
            elif oos >= thresh:
                pass
            else:
                return "落ち", "本体いる(6ヶ月在庫あり)"
    return "通過", "門通過"


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
    return gate_with_prof(p, {"冷凍生": "落とす", "本体ルール": "見ない"}, price_min, rank_max)


def read_profile(svc) -> tuple[float, float]:
    profs = load_profiles(svc)
    pick = {}
    for m in profs:
        if str(m.get("カテゴリ") or "") == "共通":
            pick = m
            break
        if str(m.get("カテゴリ") or "") in ("食品・飲料・お酒", "食品") and not pick:
            pick = m
    if not pick and profs:
        pick = profs[0]
    price_min = _num(pick.get("価格下限"), 2000)
    rank_max = rank_limit(pick) if pick else 150000.0
    return price_min, rank_max


def col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def apply_stage_dropdown(svc, apply: bool) -> int:
    from dry_t1_list_paste import ORIG_SS

    if RESEARCH_SS == ORIG_SS:
        print("refuse original")
        return 2
    raw = read_all(svc, RESEARCH_SS, T_PROF)
    if len(raw) < 2 or "門に使う順位段階" not in [str(x) for x in raw[0]]:
        print("FAIL no 門に使う順位段階")
        return 1
    h = [str(x) for x in raw[0]]
    gi = h.index("門に使う順位段階")
    mapped = []
    for r in raw[1:]:
        old = r[gi] if gi < len(r) else ""
        neu = stage_to_col(old)
        mapped.append((old, neu))
    print("g_col", gi + 1, "n", len(mapped), "sample", mapped[:3])
    meta = svc.spreadsheets().get(spreadsheetId=RESEARCH_SS, fields="sheets.properties").execute()
    sid = None
    for sh in meta.get("sheets") or []:
        if sh["properties"]["title"] == T_PROF:
            sid = sh["properties"]["sheetId"]
            break
    if sid is None:
        print("FAIL no sheet")
        return 1
    if not apply:
        print("dry dropdown 順位段階1/2/3")
        return 0
    values = [[m[1]] for m in mapped]
    rng = "'%s'!%s2" % (T_PROF.replace("'", "''"), col_letter(gi + 1))
    svc.spreadsheets().values().update(
        spreadsheetId=RESEARCH_SS,
        range=rng,
        valueInputOption="RAW",
        body={"values": values},
    ).execute()
    svc.spreadsheets().batchUpdate(
        spreadsheetId=RESEARCH_SS,
        body={
            "requests": [
                {
                    "setDataValidation": {
                        "range": {
                            "sheetId": sid,
                            "startRowIndex": 1,
                            "endRowIndex": 1 + max(len(mapped), 40),
                            "startColumnIndex": gi,
                            "endColumnIndex": gi + 1,
                        },
                        "rule": {
                            "condition": {
                                "type": "ONE_OF_LIST",
                                "values": [
                                    {"userEnteredValue": "順位段階1"},
                                    {"userEnteredValue": "順位段階2"},
                                    {"userEnteredValue": "順位段階3"},
                                ],
                            },
                            "showCustomUi": True,
                            "strict": True,
                        },
                    }
                }
            ]
        },
    ).execute()
    append_log(svc, "PROF", "runId=pr_20260816_profdrop n=%s" % len(mapped))
    print("wrote dropdown n", len(mapped))
    return 0


def main() -> int:
    log = "--log" in sys.argv
    apply_tr = "--apply-transcript" in sys.argv
    init_master = "--init-master" in sys.argv
    apply_init = "--apply" in sys.argv
    svc = sheets_service(write=True, interactive=False)
    if not svc:
        print("NO_CREDS")
        return 2
    if "--dropdown" in sys.argv:
        return apply_stage_dropdown(svc, apply_init)
    if init_master:
        from dry_t1_list_paste import ORIG_SS

        rows = default_profile_rows()
        print("init_master headers", rows[0])
        print("init_master n_rows", len(rows) - 1, "sample", [r[0] for r in rows[1:6]])
        if RESEARCH_SS == ORIG_SS:
            print("refuse original")
            return 2
        if not apply_init:
            print("dry no write")
            return 0
        rng = "'%s'!A1" % T_PROF.replace("'", "''")
        svc.spreadsheets().values().update(
            spreadsheetId=RESEARCH_SS,
            range=rng,
            valueInputOption="RAW",
            body={"values": rows},
        ).execute()
        append_log(svc, "PROF", "runId=pr_20260816_prof n=%s" % (len(rows) - 1))
        print("wrote", len(rows) - 1)
        return 0
    profs = load_profiles(svc)
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
        st, why = gate_product(p, profs) if profs else gate(p, 2000, 150000)
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
        "gate pass=%d drop=%d vs①候補 match=%d mismatch=%d empty_old=%d profs=%s"
        % (pass_n, drop_n, match, mismatch, empty_old, len(profs))
    )
    for s in samples:
        print("mismatch", s)
    ok_store = no_row == 0 and csv_hit == 0
    print("W3a_store", "PASS" if ok_store else "FAIL")
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
