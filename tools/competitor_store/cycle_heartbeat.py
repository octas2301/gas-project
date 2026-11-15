# -*- coding: utf-8 -*-
"""心拍の次アクション選定（副作用なし）。正本: CYCLE §7 / §4.4。"""
from __future__ import annotations

from typing import Any

from cycle_c3_select import (
    BAND_L,
    BAND_M,
    BAND_S,
    BAND_XL,
    K_SMALL_MED,
    catalog_band,
    catalog_estimate_for_maker,
    count_washable_small_med,
    pick_cycle_maker_select,
    pick_enum_unknown,
    pick_wash_continue_target,
)

ACT_CONTINUE = "wash_continue"
ACT_ENUM = "enum_unknown"
ACT_WASH_S = "wash_small"
ACT_WASH_M = "wash_med"
ACT_ALIGHT = "alight"
ACT_C0 = "c0"
ACT_AUTO_CK = "auto_ck"
ACT_JITTA = "jitta_resolve"
ACT_WASH_L = "wash_large"
ACT_WASH_XL = "wash_xl"
ACT_IDLE = "idle"
ACT_PAUSED = "paused"
ACT_BUSY = "other_busy"

JITTA_BACKLOG_GATE_DEFAULT = 30


def _effective_auto_ck(
    auto_ck_candidates: int,
    *,
    jitta_pending: int,
    jitta_hb_lane: bool,
    jitta_backlog_gate: int,
) -> int:
    n = max(0, int(auto_ck_candidates))
    if not jitta_hb_lane:
        return n
    if int(jitta_pending) >= int(jitta_backlog_gate):
        return 0
    return n


def _count_jitta_pending(maker_sum: list[dict]) -> int:
    """サマリの未調査（信頼度空＋未調査／レガシー要確認）。"""
    n = 0
    for row in maker_sum or []:
        conf = str(row.get("実態信頼度") or "").strip()
        if conf:
            continue
        j = str(row.get("実態") or "").strip()
        if not j or j == "未調査" or j == "要確認":
            n += 1
    return n

AMAZON_SELLER_BLOCK = frozenset(
    {
        "amazon.co.jp",
        "amazon",
        "アマゾン",
        "amazon.jp",
    }
)


def _band_for_maker(
    maker: str,
    *,
    maker_sum: list[dict],
    catalog_estimate: dict[str, int] | None,
    wash_queue: list[dict] | None,
) -> str:
    row = next((r for r in maker_sum if str(r.get("メーカー") or "").strip() == maker), {})
    est = catalog_estimate_for_maker(
        maker,
        catalog_estimate=catalog_estimate,
        wash_queue=wash_queue,
        wash_progress=str(row.get("洗い進捗") or ""),
        row=row,
    )
    return str(row.get("規模帯") or "").strip() or catalog_band(est)


def _shop_before_large_action(
    *,
    unchecked_sellers: int,
    auto_ck_candidates: int,
    untagged_gate_asins: int,
    shop_rounds_active: bool,
) -> dict[str, Any] | None:
    """小中が閾値未満の店ラウンド1手。§4.4: 自動レ点→A-light→C0。大を選定してから defer しない。"""
    # 自動レ点は店ラウンド終了フラグより優先
    if auto_ck_candidates > 0:
        return {
            "action": ACT_AUTO_CK,
            "detail": f"shop_before_large auto_ck={auto_ck_candidates}",
        }
    if not shop_rounds_active:
        return None
    if unchecked_sellers > 0:
        return {
            "action": ACT_ALIGHT,
            "detail": f"shop_before_large shops_ck_on={unchecked_sellers}",
        }
    if untagged_gate_asins > 0:
        return {
            "action": ACT_C0,
            "detail": f"shop_before_large untagged={untagged_gate_asins}",
        }
    return None


def _token_band(token_room: bool | None, token_band: str) -> str:
    b = str(token_band or "").strip()
    if b in ("none", "mid", "ample", "unknown"):
        return b
    if token_room is True:
        return "ample"
    if token_room is False:
        return "none"
    return "unknown"


def next_heartbeat_action(
    *,
    paused: bool = False,
    other_running: bool = False,
    maker_sum: list[dict] | None = None,
    sellers: list[dict] | None = None,
    wash_split: list[dict] | None = None,
    wash_queue: list[dict] | None = None,
    catalog_estimate: dict[str, int] | None = None,
    unchecked_sellers: int = 0,
    untagged_gate_asins: int = 0,
    auto_ck_candidates: int = 0,
    k: int = K_SMALL_MED,
    now_ms: int | None = None,
    shop_after_sm: str = "",
    xl_weekly_used: int = 0,
    xl_weekly_cap: int = 2,
    shop_rounds_active: bool = True,
    jitta_pending: int | None = None,
    jitta_hb_lane: bool = True,
    jitta_backlog_gate: int = JITTA_BACKLOG_GATE_DEFAULT,
    token_room: bool | None = None,
    token_band: str = "",
    wash_before_next_shop: bool = False,
    one_shop_yield: bool = True,
    restock_after_wash: bool = False,
    skip_enum_once: bool = False,
) -> dict[str, Any]:
    """戻り: {action, maker?, detail}"""
    if other_running:
        return {"action": ACT_BUSY, "detail": "other_menu_or_trigger"}
    if paused:
        return {"action": ACT_PAUSED, "detail": "await_01_or_manual"}

    maker_sum = maker_sum or []
    sellers = sellers or []

    if jitta_pending is None:
        jitta_pend = _count_jitta_pending(maker_sum) if jitta_hb_lane else 0
    else:
        jitta_pend = max(0, int(jitta_pending)) if jitta_hb_lane else 0

    auto_ck_eff = _effective_auto_ck(
        auto_ck_candidates,
        jitta_pending=jitta_pend,
        jitta_hb_lane=jitta_hb_lane,
        jitta_backlog_gate=jitta_backlog_gate,
    )

    cont = pick_wash_continue_target(
        maker_sum,
        wash_split=wash_split,
        wash_queue=wash_queue,
        skip_queue=bool(one_shop_yield),
    )
    if cont:
        return {"action": ACT_CONTINUE, "maker": cont, "detail": "in_progress"}

    shop = str(shop_after_sm or "").strip()
    if shop:
        shop_band = _band_for_maker(
            shop,
            maker_sum=maker_sum,
            catalog_estimate=catalog_estimate,
            wash_queue=wash_queue,
        )
        shop_large = shop_band in (BAND_L, BAND_XL)
        if shop_large:
            # 大・特大完走後: 門通過→セラー追加 → レ点 → A-light
            if untagged_gate_asins > 0:
                return {
                    "action": ACT_C0,
                    "detail": f"shop_after_large parent={shop} untagged={untagged_gate_asins}",
                }
            if auto_ck_eff > 0:
                return {
                    "action": ACT_AUTO_CK,
                    "detail": f"shop_after_large parent={shop} auto_ck={auto_ck_eff}",
                }
            if unchecked_sellers > 0:
                return {
                    "action": ACT_ALIGHT,
                    "detail": f"shop_after_large parent={shop} shops={unchecked_sellers}",
                }
        else:
            if auto_ck_eff > 0:
                return {
                    "action": ACT_AUTO_CK,
                    "detail": f"shop_after_sm parent={shop} auto_ck={auto_ck_eff}",
                }
            if unchecked_sellers > 0:
                return {
                    "action": ACT_ALIGHT,
                    "detail": f"shop_after_sm parent={shop} shops={unchecked_sellers}",
                }
            sm_after = count_washable_small_med(
                maker_sum,
                sellers=sellers,
                wash_queue=wash_queue,
                catalog_estimate=catalog_estimate,
                now_ms=now_ms,
            )
            if untagged_gate_asins > 0 and sm_after == 0:
                return {
                    "action": ACT_C0,
                    "detail": f"shop_after_sm parent={shop} untagged={untagged_gate_asins}",
                }

    sm_n = count_washable_small_med(
        maker_sum,
        sellers=sellers,
        wash_queue=wash_queue,
        catalog_estimate=catalog_estimate,
        now_ms=now_ms,
    )
    tok = _token_band(token_room, token_band)
    can_wash = tok in ("mid", "ample", "unknown")
    can_shop = tok == "ample"
    need_restock = sm_n < int(k)

    if not skip_enum_once:
        unk = pick_enum_unknown(
            maker_sum,
            sellers=sellers,
            wash_queue=wash_queue,
            catalog_estimate=catalog_estimate,
            cap=1,
            now_ms=now_ms,
        )
        if unk:
            return {"action": ACT_ENUM, "maker": unk[0], "detail": "band_unknown"}

    picked = pick_cycle_maker_select(
        maker_sum,
        sellers=sellers,
        wash_split=wash_split,
        wash_queue=wash_queue,
        catalog_estimate=catalog_estimate,
        cap=1,
        now_ms=now_ms,
    )

    xl_room = int(xl_weekly_used) < int(xl_weekly_cap)
    picked_sm_band = None
    picked_sm_maker = None
    if picked:
        _pm = picked[0]
        _pb = _band_for_maker(
            _pm,
            maker_sum=maker_sum,
            catalog_estimate=catalog_estimate,
            wash_queue=wash_queue,
        )
        if _pb in (BAND_S, BAND_M):
            picked_sm_band = _pb
            picked_sm_maker = _pm

    if restock_after_wash and sm_n > 0:
        if jitta_pend > 0:
            return {
                "action": ACT_JITTA,
                "detail": f"restock_after_wash jitta_pending={jitta_pend} sm_pool={sm_n}",
            }
        if can_shop:
            shop_restock = _shop_before_large_action(
                unchecked_sellers=unchecked_sellers,
                auto_ck_candidates=auto_ck_eff,
                untagged_gate_asins=untagged_gate_asins,
                shop_rounds_active=True,
            )
            if shop_restock and shop_restock.get("action") != ACT_C0:
                shop_restock["detail"] = (
                    str(shop_restock.get("detail") or "")
                    + f" restock_after_wash sm_pool={sm_n}"
                )
                return shop_restock

    if sm_n > 0 and not can_wash:
        return {"action": ACT_IDLE, "detail": f"token_reserve sm_pool={sm_n}"}

    if picked_sm_band and can_wash:
        act = ACT_WASH_S if picked_sm_band == BAND_S else ACT_WASH_M
        suffix = " restock_wash" if need_restock else ""
        return {
            "action": act,
            "maker": picked_sm_maker,
            "detail": f"sm_pool={sm_n}{suffix}",
        }

    if sm_n == 0 and jitta_pend > 0:
        return {
            "action": ACT_JITTA,
            "detail": f"sm_empty jitta_pending={jitta_pend} sm_pool=0",
        }
    if sm_n == 0:
        shop_empty = _shop_before_large_action(
            unchecked_sellers=unchecked_sellers,
            auto_ck_candidates=auto_ck_eff,
            untagged_gate_asins=untagged_gate_asins,
            shop_rounds_active=shop_rounds_active,
        )
        if shop_empty:
            return shop_empty

    if sm_n == 0 and picked:
        m = picked[0]
        band = _band_for_maker(
            m,
            maker_sum=maker_sum,
            catalog_estimate=catalog_estimate,
            wash_queue=wash_queue,
        )
        if band in (BAND_L, BAND_XL):
            if not can_wash:
                return {"action": ACT_IDLE, "detail": "token_reserve xl_hold sm_empty"}
            why = "shop_rounds_exhausted" if not shop_rounds_active else "shop_material_empty"
            if not xl_room:
                if jitta_pend > 0:
                    return {
                        "action": ACT_JITTA,
                        "detail": f"xl_week_cap jitta_pending={jitta_pend} {why}",
                    }
                return {
                    "action": ACT_IDLE,
                    "detail": (
                        f"xl_week_cap count={xl_weekly_used}/{xl_weekly_cap} {why}"
                    ),
                }
            if band == BAND_L:
                return {
                    "action": ACT_WASH_L,
                    "maker": m,
                    "detail": f"{why} sm_empty xl_week={xl_weekly_used}",
                }
            return {
                "action": ACT_WASH_XL,
                "maker": m,
                "detail": f"{why} sm_l_empty xl_week={xl_weekly_used}",
            }

    if sm_n == 0 and auto_ck_eff > 0:
        return {"action": ACT_AUTO_CK, "detail": f"sm_pool=0 auto_ck={auto_ck_eff}"}
    if sm_n == 0 and unchecked_sellers > 0:
        return {"action": ACT_ALIGHT, "detail": f"sm_pool=0 shops={unchecked_sellers}"}
    if sm_n == 0 and untagged_gate_asins > 0:
        return {"action": ACT_C0, "detail": f"untagged={untagged_gate_asins}"}
    if jitta_pend > 0:
        return {
            "action": ACT_JITTA,
            "detail": f"drain_pending={jitta_pend}",
        }
    return {"action": ACT_IDLE, "detail": "no_work"}


def should_arm_01_resume(*, paused: bool, manual_resumed: bool) -> bool:
    if manual_resumed:
        return False
    return paused


def on_01_fire(*, paused: bool, manual_resumed: bool, other_running: bool) -> str:
    if manual_resumed or not paused:
        return "noop_manual" if manual_resumed else "noop_not_paused"
    if other_running:
        return "defer_busy"
    return "resume"


def seller_is_blocked_for_auto_ck(row: dict) -> bool:
    name = str(row.get("店名") or row.get("name") or "").strip().lower()
    sid = str(row.get("sellerId") or row.get("SellerId") or "").strip().lower()
    if not name and not sid:
        return False
    for b in AMAZON_SELLER_BLOCK:
        if b in name or b == sid:
            return True
    if str(row.get("hold") or "").upper() in ("TRUE", "1", "Y", "HOLD"):
        return True
    if str(row.get("保留") or "").strip():
        return True
    return False


def pick_auto_ck_sellers(
    sellers: list[dict],
    *,
    daily_cap: int = 3,
    already_cked_today: int = 0,
) -> list[str]:
    """
    調査CK=OFF のうち店規模 小→未知→中→大→特大、同帯は小規模出現が多い店を先に。Amazon／hold 除外。
    戻り: sellerId リスト。
    """
    left = max(0, int(daily_cap) - int(already_cked_today))
    if left <= 0:
        return []
    pool: list[tuple[int, int, int, str]] = []
    for i, row in enumerate(sellers):
        sid = str(row.get("sellerId") or row.get("SellerId") or "").strip()
        if not sid:
            continue
        ck = str(row.get("調査CK") or "").upper()
        if ck in ("TRUE", "1"):
            continue
        if seller_is_blocked_for_auto_ck(row):
            continue
        hit = int(row.get("小規模出現") or row.get("small_hit") or 0)
        band = str(row.get("規模帯") or row.get("band") or "").strip()
        pool.append((_auto_ck_band_rank(band), -hit, i, sid))
    pool.sort()
    return [sid for _, _, _, sid in pool[:left]]


def _auto_ck_band_rank(band: str) -> int:
    b = str(band or "").strip()
    if b == BAND_S:
        return 1
    if b == BAND_M:
        return 3
    if b == BAND_L:
        return 4
    if b == BAND_XL:
        return 5
    return 2


def recompute_small_maker_hits(
    *,
    gate_asins: list[dict],
    maker_bands: dict[str, str],
) -> dict[str, int]:
    """
    門通過 ASIN 行 {sellerId, メーカー} と メーカー→規模帯 から sellerId→小規模出現。
    小・中の門通過のみカウント。
    """
    out: dict[str, int] = {}
    for row in gate_asins:
        maker = str(row.get("メーカー") or "").strip()
        band = maker_bands.get(maker) or ""
        if band not in (BAND_S, BAND_M):
            continue
        sid = str(row.get("sellerId") or row.get("BuyBoxセラー") or "").strip()
        if not sid:
            continue
        out[sid] = out.get(sid, 0) + 1
    return out


def build_daily_mail_body(
    *,
    day: str,
    paused: bool,
    completed_makers: list[str],
    top_categories: list[str],
    gate_pass_n: int,
    enum_n: int = 0,
    alight_shops: int = 0,
    c0_new: int = 0,
    sm_pool: int = 0,
) -> str:
    state = "一時停止" if paused else "稼働または待機"
    lines = [
        "自動リサーチ日次（生存確認）",
        f"日付: {day}",
        f"心拍: {state}",
        f"完走メーカー({len(completed_makers)}): " + ("、".join(completed_makers) if completed_makers else "（なし）"),
        f"主要カテゴリ: " + ("、".join(top_categories[:8]) if top_categories else "不明"),
        f"門通過数: {gate_pass_n}",
        f"参考: 列挙={enum_n} A-light店={alight_shops} C0新規={c0_new} 小中プール={sm_pool}",
    ]
    return "\n".join(lines) + "\n"
