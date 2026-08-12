# -*- coding: utf-8 -*-
"""シートヘッダ定義（要件§2）。"""
from __future__ import annotations

# 対象SKU名簿（やる／やらない）
MASTER_SHEET = "タイムセール_マスタ"
# 施策（1行=1施策・レーン列・新着は上）
SALE_SHEET = "タイムセール"
ANALYSIS_SHEET = "タイムセール_分析"

LEGACY_EXEC = "タイムセール_実行"
LEGACY_LANE_A = "タイムセール_期間値下げ"
OLD_LOG_SHEET = "タイムセール_ログ"

LANE_A = "A_期間値下げ"
LANE_B = "B_公式"

# 人入力セル（行2以降）の背景色
MASTER_HUMAN_INPUT_BG = "#FFF2CC"

# 旧ヘッダ名 → 現行名（並替・移行時）
MASTER_HEADER_ALIASES = {
    "最終売価円": "目標売価円",
    "ポイント目標%": "期間中ポイント%",
    "ポイント現在%": "出品者ポイント現在%",
    # 旧「戻し*」（売価段上げ時代の名）→ ポイント減衰運用名
    "戻し期間": "減衰期間",
    "戻し価格円": "減衰段%",  # 旧名のまま円と誤認されやすい。実体は段%
    "戻し間隔": "減衰間隔",
    "戻し進捗": "減衰進捗",
    "戻し状態": "減衰状態",
    "次回戻し日": "次回減衰日",
}

# 削除する列（移行時に捨てる。データは移行しない）
MASTER_HEADERS_DROPPED = frozenset(
    {
        "販促売価円",  # 非推奨・使わない
    }
)

# ---------------------------------------------------------------------------
# マスタ列順（目的グループ。各グループ内は人入力→表示／システム）
# ---------------------------------------------------------------------------
# 実質戻し: 人=目標売価・販促%・減衰期間/段%/間隔 → 円・実質（数式）→ 進捗系
MASTER_HEADERS = [
    # 基本
    "SKU",
    "ASIN",
    "親ASIN",
    "商品名",
    "画像URL",
    "marketplace",
    "通貨",
    "有効",
    # SC取込
    "出品者価格_SC",
    "タイムセール価格_SC",
    "販売商品数_SC",
    "V30",
    "Q_fba",
    "原価U",
    # ポイント（人入力先頭）
    "期間中ポイント%",
    "ポイントメモ",
    "期間中ポイント円",
    "セール前ポイント%",
    "セール前ポイント円",
    "出品者ポイント現在%",
    "出品者ポイント現在円",
    "ポイント状態",
    # 実質戻し＝ポイント減衰（人入力先頭）
    "目標売価円",  # 最終売価（our_price・原則固定）人必須
    "販促ポイント%",  # 人必須: 減衰開始の出品者付与%
    "減衰期間",  # 人／提案: 1〜6か月
    "減衰段%",  # 人／提案: 1回に下げるポイント%ポイント
    "減衰間隔",  # 人／提案: 週・月
    "減衰開始日",  # 人: カレンダー1段目の日付（例: 現行TS終了翌日）
    "減衰実行依頼",  # E: TRUE でポーラー／今すぐ対象
    "販促ポイント円",  # 数式表示
    "実質価格円",  # 数式表示
    "減衰中ポイント%",  # 運用目標（提案時=販促%。taper成功後更新）
    "次回減衰後%",  # 数式: MAX(終着, 減衰中−段%)
    "減衰進捗",
    "減衰状態",
    "次回減衰日",
    "最終減衰実行日時",
    "現在売価円",
    # レーンA実績
    "A実施",
    "A最終送付日時",
    "A期間",
    "A価格円",
    "Aログ参照",
    # メモ
    "メモ",
]

# 行2以降を黄色にする列（人入力）
MASTER_HUMAN_INPUT_COLS = (
    "有効",
    "期間中ポイント%",
    "ポイントメモ",
    "目標売価円",
    "販促ポイント%",
    "減衰期間",
    "減衰段%",
    "減衰間隔",
    "減衰開始日",
    "減衰実行依頼",
    "メモ",
)

# 実質戻し（売価固定＋ポイント減衰）。人必須=目標売価円＋販促ポイント%。提案はGAS/Python。
PRICE_TARGET_COL = "目標売価円"  # 最終売価
PRICE_FINAL_COL = PRICE_TARGET_COL
PROMO_POINT_PCT_COL = "販促ポイント%"
PROMO_POINT_YEN_COL = "販促ポイント円"
EFFECTIVE_PRICE_COL = "実質価格円"
# 表示列はシート数式で常時計算
TAPER_REQUEST_COL = "減衰実行依頼"
TAPER_ACTIVE_COL = "減衰中ポイント%"
TAPER_NEXT_PCT_COL = "次回減衰後%"
TAPER_LAST_RUN_COL = "最終減衰実行日時"
MASTER_DISPLAY_FORMULA_COLS = (
    PROMO_POINT_YEN_COL,
    EFFECTIVE_PRICE_COL,
    TAPER_NEXT_PCT_COL,
)
# 減衰スケジュール（マスタで管理。提案メニュー／price_recovery_*／taper_send が読む）
PRICE_RECOVERY_PERIOD_COL = "減衰期間"
PRICE_RECOVERY_STEP_COL = "減衰段%"
PRICE_RECOVERY_INTERVAL_COL = "減衰間隔"
TAPER_START_COL = "減衰開始日"
PRICE_RECOVERY_PROGRESS_COL = "減衰進捗"
PRICE_RECOVERY_STATUS_COL = "減衰状態"
PRICE_RECOVERY_NEXT_COL = "次回減衰日"
PRICE_CURRENT_SELL_COL = "現在売価円"
PRICE_RECOVERY_COLS = (
    PRICE_TARGET_COL,
    PROMO_POINT_PCT_COL,
    PRICE_RECOVERY_PERIOD_COL,
    PRICE_RECOVERY_STEP_COL,
    PRICE_RECOVERY_INTERVAL_COL,
    TAPER_START_COL,
    TAPER_REQUEST_COL,
    PROMO_POINT_YEN_COL,
    EFFECTIVE_PRICE_COL,
    TAPER_ACTIVE_COL,
    TAPER_NEXT_PCT_COL,
    PRICE_RECOVERY_PROGRESS_COL,
    PRICE_RECOVERY_STATUS_COL,
    PRICE_RECOVERY_NEXT_COL,
    TAPER_LAST_RUN_COL,
    PRICE_CURRENT_SELL_COL,
)
# 互換（削除済み列名。コード参照の残骸用）
PRICE_PROMO_COL = "販促売価円"

# マスタ1行目ヘッダのグルーピング色（#RRGGBB）
MASTER_HEADER_COLOR_GROUPS = (
    (
        "基本",
        "#E8EAED",
        ("SKU", "ASIN", "親ASIN", "商品名", "画像URL", "marketplace", "通貨", "有効"),
    ),
    (
        "SC取込",
        "#D2E3FC",
        (
            "出品者価格_SC",
            "タイムセール価格_SC",
            "販売商品数_SC",
            "V30",
            "Q_fba",
            "原価U",
        ),
    ),
    (
        "ポイント",
        "#E8DEF8",
        (
            "期間中ポイント%",
            "ポイントメモ",
            "期間中ポイント円",
            "セール前ポイント%",
            "セール前ポイント円",
            "出品者ポイント現在%",
            "出品者ポイント現在円",
            "ポイント状態",
        ),
    ),
    (
        "実質戻し",
        "#FCE8C3",
        PRICE_RECOVERY_COLS,
    ),
    (
        "レーンA実績",
        "#C8E6C9",
        ("A実施", "A最終送付日時", "A期間", "A価格円", "Aログ参照"),
    ),
    ("メモ", "#FFF9C4", ("メモ",)),
)

# Phase0: タイムセール_マスタ上のポイント管理（§10.10）
POINT_PERIOD_COL = "期間中ポイント%"
POINT_PERIOD_YEN_COL = "期間中ポイント円"
POINT_BEFORE_COL = "セール前ポイント%"  # 運用上は最終終着%（減衰フロア）。fetchで上書きしない
POINT_END_COL = POINT_BEFORE_COL
POINT_BEFORE_YEN_COL = "セール前ポイント円"
POINT_CURRENT_COL = "出品者ポイント現在%"
POINT_CURRENT_YEN_COL = "出品者ポイント現在円"
POINT_STATUS_COL = "ポイント状態"
POINT_MEMO_COL = "ポイントメモ"
POINT_COLS = (
    POINT_PERIOD_COL,
    POINT_MEMO_COL,
    POINT_PERIOD_YEN_COL,
    POINT_BEFORE_COL,
    POINT_BEFORE_YEN_COL,
    POINT_CURRENT_COL,
    POINT_CURRENT_YEN_COL,
    POINT_STATUS_COL,
)

A_DONE_COL = "A実施"
A_SENT_AT_COL = "A最終送付日時"
A_PERIOD_COL = "A期間"
A_PRICE_COL = "A価格円"
A_LOG_COL = "Aログ参照"
A_TRACK_COLS = (
    A_DONE_COL,
    A_SENT_AT_COL,
    A_PERIOD_COL,
    A_PRICE_COL,
    A_LOG_COL,
)
SALE_A_TRACK_COLS = (A_DONE_COL, A_SENT_AT_COL, A_LOG_COL)

LEGACY_POINT_TARGET = "ポイント目標%"
LEGACY_POINT_CURRENT = "ポイント現在%"

SALE_HEADERS = [
    "sale_id",
    "レーン",
    "SKU",
    "ASIN",
    "親ASIN",
    "商品名",
    "画像",
    "marketplace",
    "通貨",
    "有効",
    "承認済",
    "種別",
    "スケジュール",
    "開始日",
    "終了日",
    "出品者価格_SC",
    "タイムセール価格_SC",
    "タイムセール価格_確定",
    "販売商品数_SC",
    "V30",
    "販売商品数_確定",
    "通常価格",
    "セール価格",
    "目標販売",
    "想定利益",
    "原価U",
    "提出対象",
    "状態",
    "更新日時",
    "runId",
    "メッセージ",
    "A実施",
    "A最終送付日時",
    "Aログ参照",
    "メモ",
]

EXEC_SHEET = SALE_SHEET
LANE_A_SHEET = SALE_SHEET
EXEC_HEADERS = SALE_HEADERS
LANE_A_HEADERS = SALE_HEADERS
LEGACY_MASTER = MASTER_SHEET
POINT_TARGET_COL = POINT_PERIOD_COL


def normalize_master_header_name(name: str) -> str:
    s = str(name or "").strip()
    return MASTER_HEADER_ALIASES.get(s, s)


def a_track_fields_from_row(r: dict, *, sale: bool = False) -> dict:
    """マスタ／施策のA実績列を保全するための辞書。"""
    cols = SALE_A_TRACK_COLS if sale else A_TRACK_COLS
    return {c: (r.get(c) if r.get(c) is not None else "") for c in cols}
