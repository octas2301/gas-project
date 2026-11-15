# -*- coding: utf-8 -*-
"""チェック済み ∩ 競合写真1枚以上・1テーマ1枚・想像で埋めない。"""
from sub_image_lp_themes import THEME_BY_ID, select_lp_slots, select_lp_slots_forced

hits = {
    1: {
        "themeId": 1,
        "count": 3,
        "score": 2.0,
        "paths": ["a.jpg"],
        "ocrSnippets": [],
        "phaseOrder": 1,
        "contentCluster": "emotion",
    },
    3: {
        "themeId": 3,
        "count": 1,
        "score": 1.0,
        "paths": ["b.jpg"],
        "ocrSnippets": [],
        "phaseOrder": 1,
        "contentCluster": "sizzle",
    },
}

# チェック 3,12,20 のうち写真があるのは 3 だけ
jobs = select_lp_slots(hits, forced_theme_ids=[3, 12, 20], proposals_per_slot=1)
ids = [int(j["themeId"]) for j in jobs]
assert ids == [3], ids
assert all(j.get("proposal") == "a" for j in jobs)
src = {int(j["themeId"]): j["source"] for j in jobs}
assert src[3] == "competitor_forced", src
assert 12 not in ids
assert 20 not in ids
assert 1 not in ids

# チェック指定が空リスト＝0枚（自動埋めしない）
empty_forced = select_lp_slots(hits, forced_theme_ids=[], proposals_per_slot=1)
assert empty_forced == []

# 不正IDのみ＝0枚（自動へ落とさない）
bad = select_lp_slots(hits, forced_theme_ids=[99, 0], proposals_per_slot=1)
assert bad == []

# チェック無し（None）＝写真があるテーマだけ
auto = select_lp_slots(hits, target_slots=10, proposals_per_slot=1, allow_invented=False)
auto_ids = [int(j["themeId"]) for j in auto]
assert set(auto_ids) == {1, 3}, auto_ids
assert all(j.get("proposal") == "a" for j in auto)

# 想像ONの旧挙動は残す
invented = select_lp_slots(
    hits, target_slots=5, proposals_per_slot=1, allow_invented=True
)
assert len(invented) >= 2
assert all(int(j["themeId"]) in THEME_BY_ID for j in invented)

direct = select_lp_slots_forced(hits, [3, 12], proposals_per_slot=1)
assert [int(j["themeId"]) for j in direct] == [3]

print("ALL_PASS", "forced", ids, "auto", auto_ids, "inventedN", len(invented))
