# -*- coding: utf-8 -*-
"""スケジュール分類: 名付き公式Bを先に固定→空きにA（最大14日）。"""
from __future__ import annotations

from datetime import date, datetime, timedelta
from typing import List, Optional, Tuple

# 要件§0 大型イベント語＋公式SaleとしてBに残す
OFFICIAL_B_TOKENS = (
    "プライム感謝",
    "プライムデー",
    "プライム",
    "ブラックフライデー",
    "スマイル",
    "smile",
    "Smile",
)

# 要件§10.6 / §10.7（JP）
COOLDOWN_OSUSUME_DAYS = 14  # おすすめ同士
COOLDOWN_LIGHTNING_DAYS = 7  # 数量限定同士（将来）
A_MAX_DAYS = 14  # レーンAも運用上1本最大14日
AB_GAP_DAYS_DEFAULT = 0  # A↔Bの空き（未確証。検証は 0→5→10→14）
B_HORIZON_DAYS_DEFAULT = 90  # 通常の取り込み水平線
# 早期申請割引（要件§10.9）。締切前なら horizon 外でもBに含める
EARLY_FEE_DEADLINES_DEFAULT = {
    "プライム感謝": "2026-08-05",
    "ブラックフライデー": "2026-09-30",
    "サイバー": "2026-09-30",
}


def is_official_b_schedule(schedule: str) -> bool:
    s = str(schedule or "")
    if not s.strip():
        return False
    low = s.lower()
    for tok in OFFICIAL_B_TOKENS:
        if tok.lower() in low or tok in s:
            return True
    return False


def is_a_equivalent_schedule(schedule: str) -> bool:
    """月枠・名なしカスタムなど → 自動化（A）相当。バルク本線にしない。"""
    s = str(schedule or "").strip()
    if not s:
        return False
    if is_official_b_schedule(s):
        return False
    if s.startswith("月") or "月 (" in s or s.startswith("カスタム"):
        return True
    if str(s).startswith("A期間"):
        return True
    return True


def classify_schedule(schedule: str) -> str:
    from sheet_schema import LANE_A, LANE_B

    if is_official_b_schedule(schedule):
        return LANE_B
    return LANE_A


def parse_ymd(s: Optional[object]) -> Optional[date]:
    """YYYY-MM-DD または Sheets/Excel シリアル日付を date に。"""
    if s is None:
        return None
    if isinstance(s, datetime):
        return s.date()
    if isinstance(s, date):
        return s
    raw = str(s).strip()
    if not raw or raw.upper() in ("NA",):
        return None
    # Sheets が日付をシリアルで返す場合（例: 46262 = 2026-08-28）
    try:
        if isinstance(s, (int, float)) or (
            raw.replace(".", "", 1).isdigit() and "e" not in raw.lower()
        ):
            n = float(raw)
            if 20000.0 <= n <= 80000.0:
                return date(1899, 12, 30) + timedelta(days=int(n))
    except (ValueError, OverflowError):
        pass
    try:
        return datetime.strptime(raw[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def format_ymd(s: Optional[object]) -> str:
    """シート書き込み用に YYYY-MM-DD 文字列へ。解析不能なら空。"""
    d = parse_ymd(s)
    return d.isoformat() if d else ""


def ranges_overlap(
    a0: Optional[date], a1: Optional[date], b0: Optional[date], b1: Optional[date]
) -> bool:
    if not a0 or not a1 or not b0 or not b1:
        return False
    return a0 <= b1 and b0 <= a1


def cooldown_ok(
    prev_end: Optional[date],
    next_start: Optional[date],
    *,
    days: int = COOLDOWN_OSUSUME_DAYS,
) -> bool:
    """
    おすすめ再実施: 次の開始日 - 前回終了日 >= days なら可。
    日付欠落時は判定不能→True（呼び側で要確認）。
    """
    if not prev_end or not next_start:
        return True
    return (next_start - prev_end).days >= days


def earliest_after_cooldown(
    prev_end: Optional[date], *, days: int = COOLDOWN_OSUSUME_DAYS
) -> Optional[date]:
    if not prev_end:
        return None
    return prev_end + timedelta(days=days)


def a_end_before_b(b_start: date, gap_days: int) -> date:
    """空き gap_days 日のあと B 開始。gap=0 なら A終了は B開始の前日。"""
    return b_start - timedelta(days=1 + max(0, int(gap_days)))


def a_start_after_b(b_end: date, gap_days: int) -> date:
    """B終了のあと空き gap_days 日。gap=0 なら A開始は B終了の翌日。"""
    return b_end + timedelta(days=1 + max(0, int(gap_days)))


def _place_chunk(
    g0: date, g1: date, *, a_max_days: int, prefer_end: bool
) -> Optional[Tuple[date, date]]:
    if g1 < g0:
        return None
    span = (g1 - g0).days + 1
    if span <= 0:
        return None
    use = min(int(a_max_days), span)
    if prefer_end:
        a1 = g1
        a0 = a1 - timedelta(days=use - 1)
        if a0 < g0:
            a0 = g0
            a1 = min(g1, a0 + timedelta(days=use - 1))
    else:
        a0 = g0
        a1 = a0 + timedelta(days=use - 1)
        if a1 > g1:
            a1 = g1
            a0 = max(g0, a1 - timedelta(days=use - 1))
    if a1 < a0:
        return None
    return a0, a1


def fill_a_windows(
    b_picked: List[dict],
    *,
    today: date,
    limit_a: int = 2,
    a_max_days: int = A_MAX_DAYS,
    gap_days: int = AB_GAP_DAYS_DEFAULT,
    horizon_end: Optional[date] = None,
) -> List[dict]:
    """
    B占有を避け、空きに最大 a_max_days のAを最大 limit_a 本。
    B直前の空きは終端寄せ（間隔検証向き）、B直後は始端寄せ。
    """
    if horizon_end is None:
        horizon_end = today + timedelta(days=120)

    blockers: List[Tuple[date, date]] = []
    for b in b_picked:
        s, e = parse_ymd(b.get("start")), parse_ymd(b.get("end"))
        if s and e and e >= today:
            blockers.append((s, e))
    blockers.sort(key=lambda x: x[0])

    # (g0, g1, prefer_end)
    raw_gaps: List[Tuple[date, date, bool]] = []
    cursor = today
    for i, (bs, be) in enumerate(blockers):
        free_end = a_end_before_b(bs, gap_days)
        if free_end >= cursor:
            # 最初のB直前だけ終端寄せ（間隔検証）。BとBの間は直後から埋める
            prefer_end = i == 0
            raw_gaps.append((cursor, free_end, prefer_end))
        cursor = max(cursor, a_start_after_b(be, gap_days))
    if cursor <= horizon_end:
        raw_gaps.append((cursor, horizon_end, False))

    out: List[dict] = []
    for g0, g1, prefer_end in raw_gaps:
        if len(out) >= limit_a:
            break
        placed = _place_chunk(g0, g1, a_max_days=a_max_days, prefer_end=prefer_end)
        if not placed:
            continue
        a0, a1 = placed
        out.append(
            {
                "schedule": f"A期間 ({a0.isoformat()} - {a1.isoformat()})",
                "start": a0.isoformat(),
                "end": a1.isoformat(),
                "synthetic": True,
                "ab_gap_days": int(gap_days),
            }
        )
    return out


def _pick_non_overlapping(
    items: List[dict],
    *,
    limit: int,
    today: date,
    allow_undated: bool,
    cooldown_days: Optional[int] = None,
) -> List[dict]:
    """日付順に貪欲。期間非重なり。cooldown_days指定時は終了→次開始の間隔も見る。"""
    out: List[dict] = []
    undated_used = False
    for x in items:
        end = parse_ymd(x.get("end"))
        start = parse_ymd(x.get("start"))
        if end and end < today:
            continue
        if start and end is None and start < today:
            continue
        if start and end:
            if any(
                ranges_overlap(
                    start,
                    end,
                    parse_ymd(p.get("start")),
                    parse_ymd(p.get("end")),
                )
                for p in out
            ):
                continue
            if cooldown_days is not None:
                bad = False
                for p in out:
                    pe = parse_ymd(p.get("end"))
                    ps = parse_ymd(p.get("start"))
                    if pe and start and start >= pe:
                        if not cooldown_ok(pe, start, days=cooldown_days):
                            bad = True
                            break
                    elif end and ps and end <= ps:
                        if not cooldown_ok(end, ps, days=cooldown_days):
                            bad = True
                            break
                if bad:
                    continue
        else:
            if not allow_undated:
                continue
            if undated_used:
                continue
            undated_used = True
        out.append(x)
        if len(out) >= limit:
            break
    return out


def within_early_fee_window(
    schedule: str,
    *,
    today: date,
    deadlines: Optional[dict] = None,
) -> bool:
    """早期申請割引の申請締切前なら True（horizon外でもB候補に残す）。"""
    dlmap = deadlines if deadlines is not None else EARLY_FEE_DEADLINES_DEFAULT
    name = str(schedule or "")
    low = name.lower()
    for token, ymd in dlmap.items():
        hit = False
        if token in name:
            hit = True
        elif token == "ブラックフライデー" and (
            "ブラック" in name or "black" in low or "フライデー" in name
        ):
            hit = True
        elif token == "サイバー" and ("サイバー" in name or "cyber" in low):
            hit = True
        elif str(token).startswith("プライム") and "プライム感謝" in name:
            hit = True
        if not hit:
            continue
        dl = parse_ymd(str(ymd))
        if dl and today <= dl:
            return True
    return False


def pick_schedules_split(
    candidates: List[dict],
    *,
    limit_b: int = 2,
    limit_a: int = 2,
    today: Optional[date] = None,
    ab_gap_days: int = AB_GAP_DAYS_DEFAULT,
    a_max_days: int = A_MAX_DAYS,
    b_horizon_days: int = B_HORIZON_DAYS_DEFAULT,
    early_fee_deadlines: Optional[dict] = None,
) -> Tuple[List[dict], List[dict]]:
    """
    1) 名付き公式Bを固定（非重なり＋おすすめ14日。horizon外でも早期割引締切前は残す）
    2) 空きに合成A（最大 a_max_days × limit_a）。月枠候補は使わない
    """
    today = today or date.today()
    horizon_cut = today + timedelta(days=max(0, int(b_horizon_days)))
    dlmap = early_fee_deadlines if early_fee_deadlines is not None else EARLY_FEE_DEADLINES_DEFAULT

    b_list: List[dict] = []
    for c in candidates:
        name = c.get("schedule") or ""
        if not is_official_b_schedule(name):
            continue
        st = parse_ymd(c.get("start"))
        early = within_early_fee_window(str(name), today=today, deadlines=dlmap)
        if st and st > horizon_cut and not early:
            continue
        b_list.append(c)

    def b_key(x):
        has = 0 if parse_ymd(x.get("start")) else 1
        return has, parse_ymd(x.get("start")) or date.max, str(x.get("schedule") or "")

    b_list.sort(key=b_key)
    b_picked = _pick_non_overlapping(
        b_list,
        limit=limit_b,
        today=today,
        allow_undated=False,
        cooldown_days=COOLDOWN_OSUSUME_DAYS,
    )
    if not b_picked:
        b_picked = _pick_non_overlapping(
            b_list,
            limit=limit_b,
            today=today,
            allow_undated=True,
            cooldown_days=COOLDOWN_OSUSUME_DAYS,
        )
    a_picked = fill_a_windows(
        b_picked,
        today=today,
        limit_a=limit_a,
        a_max_days=a_max_days,
        gap_days=ab_gap_days,
    )
    return b_picked, a_picked


def is_mainline_osusume_type(deal_type: object) -> bool:
    s = str(deal_type or "").strip()
    if not s:
        return True
    if "数量限定" in s or "LIGHTNING" in s.upper():
        return False
    if "おすすめ" in s or "BEST_DEAL" in s.upper() or "BEST DEAL" in s.upper():
        return True
    return False


def shrink_range_avoiding_blockers(
    start: date, end: date, blockers: List[Tuple[date, date]]
) -> Optional[Tuple[date, date]]:
    """公式B等と重なる部分を避け、残る最長区間を返す。無ければ None。"""
    segments: List[Tuple[date, date]] = [(start, end)]
    for b0, b1 in blockers:
        nxt: List[Tuple[date, date]] = []
        for s0, s1 in segments:
            if not ranges_overlap(s0, s1, b0, b1):
                nxt.append((s0, s1))
                continue
            if s0 < b0:
                left_end = min(s1, b0 - timedelta(days=1))
                if left_end >= s0:
                    nxt.append((s0, left_end))
            if s1 > b1:
                right_start = max(s0, b1 + timedelta(days=1))
                if s1 >= right_start:
                    nxt.append((right_start, s1))
        segments = nxt
        if not segments:
            return None
    best = max(segments, key=lambda se: (se[1] - se[0]).days)
    return best


def pick_named_within_horizon(
    candidates: List[dict],
    *,
    today: Optional[date] = None,
    b_horizon_days: int = B_HORIZON_DAYS_DEFAULT,
    limit_b: int = 2,
) -> List[dict]:
    """
    カタログ名付きのうち開始が horizon 以内のみ（早期割引で遠方BFを引き込まない）。
    台帳への新規追加用。
    """
    today = today or date.today()
    horizon_cut = today + timedelta(days=max(0, int(b_horizon_days)))
    named = []
    for c in candidates:
        name = str(c.get("schedule") or "")
        if not is_official_b_schedule(name):
            continue
        st = parse_ymd(c.get("start"))
        en = parse_ymd(c.get("end"))
        if en and en < today:
            continue
        if st and st > horizon_cut:
            continue
        if not st:
            continue
        named.append(c)
    named.sort(key=lambda x: parse_ymd(x.get("start")) or date.max)
    return _pick_non_overlapping(
        named,
        limit=limit_b,
        today=today,
        allow_undated=False,
        cooldown_days=COOLDOWN_OSUSUME_DAYS,
    )


def pick_schedules_sku_local(
    local: List[dict],
    *,
    limit_b: int = 2,
    limit_a: int = 2,
    limit_bulk_custom: int = 2,
    today: Optional[date] = None,
    ab_gap_days: int = AB_GAP_DAYS_DEFAULT,
    a_max_days: int = A_MAX_DAYS,
    b_horizon_days: int = B_HORIZON_DAYS_DEFAULT,
    early_fee_deadlines: Optional[dict] = None,
    skip_schedule_names: Optional[set] = None,
) -> Tuple[List[dict], List[dict], List[dict]]:
    """
    当該SKUの②行だけを正（カタログBF等は入れない）。

    戻り値: (名付きB, バルク提出するカスタム/月おすすめ, 合成A)
    - 名付き／カスタムの日付は②の値を正（勝手に伸ばさない）
    - カスタムが公式Bと重なるときは短縮（公式優先）
    - skip_schedule_names: 既登録（UL済等）で再提出しない枠
    """
    today = today or date.today()
    skip = {str(x).strip() for x in (skip_schedule_names or set()) if str(x).strip()}
    named_src = []
    for c in local:
        name = str(c.get("schedule") or "").strip()
        if not is_official_b_schedule(name):
            continue
        if name in skip:
            continue
        if not is_mainline_osusume_type(c.get("deal_type")):
            continue
        named_src.append(c)
    b_picked, _ = pick_schedules_split(
        named_src,
        limit_b=limit_b,
        limit_a=0,
        today=today,
        ab_gap_days=ab_gap_days,
        a_max_days=a_max_days,
        b_horizon_days=b_horizon_days,
        early_fee_deadlines=early_fee_deadlines,
    )

    blockers: List[Tuple[date, date]] = []
    for b in b_picked:
        s, e = parse_ymd(b.get("start")), parse_ymd(b.get("end"))
        if s and e:
            blockers.append((s, e))

    customs: List[dict] = []
    for c in local:
        name = str(c.get("schedule") or "").strip()
        if not name or is_official_b_schedule(name):
            continue
        if name in skip:
            continue
        if not is_mainline_osusume_type(c.get("deal_type")):
            continue
        # 合成A期間は②に無い
        if str(name).startswith("A期間"):
            continue
        s, e = parse_ymd(c.get("start")), parse_ymd(c.get("end"))
        if not s or not e or e < today:
            continue
        clipped = shrink_range_avoiding_blockers(s, e, blockers)
        if not clipped:
            continue
        cs, ce = clipped
        # Amazon提示の枠を超えない（短縮のみ）
        entry = dict(c)
        entry["start"] = cs.isoformat()
        entry["end"] = ce.isoformat()
        if (cs, ce) != (s, e):
            entry["clipped"] = True
            entry["schedule_original"] = name
        # スケジュール表示名は②のまま（SCドロップダウン一致）。日付列だけ短縮可
        customs.append(entry)

    def c_key(x):
        return parse_ymd(x.get("start")) or date.max, str(x.get("schedule") or "")

    customs.sort(key=c_key)
    bulk_custom = _pick_non_overlapping(
        customs,
        limit=limit_bulk_custom,
        today=today,
        allow_undated=False,
        cooldown_days=COOLDOWN_OSUSUME_DAYS,
    )
    for bc in bulk_custom:
        s, e = parse_ymd(bc.get("start")), parse_ymd(bc.get("end"))
        if s and e:
            blockers.append((s, e))

    blockers_as_b = []
    for s, e in blockers:
        blockers_as_b.append({"start": s.isoformat(), "end": e.isoformat()})
    a_picked = fill_a_windows(
        blockers_as_b,
        today=today,
        limit_a=limit_a,
        a_max_days=a_max_days,
        gap_days=ab_gap_days,
    )
    return b_picked, bulk_custom, a_picked
