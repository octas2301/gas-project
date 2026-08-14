#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""▼設定(Amazonマッピング) を Yahooマッピングの右隣に作成する（初期投入・試行用）。"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any, List

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCRIPT_DIR = Path(__file__).resolve().parent
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
TOKEN_RW = SCRIPT_DIR / "secrets" / "token_sheets_rw.json"
SHEET_TITLE = "▼設定(Amazonマッピング)"
YAHOO_TITLE = "▼設定(Yahooマッピング)"


def load_cfg() -> dict:
    return json.loads((SCRIPT_DIR / "config.local.json").read_text(encoding="utf-8"))


def get_creds() -> Credentials:
    cred_path = SCRIPT_DIR / "secrets" / "credentials.json"
    creds = None
    if TOKEN_RW.is_file():
        creds = Credentials.from_authorized_user_file(str(TOKEN_RW), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(cred_path), SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_RW.parent.mkdir(parents=True, exist_ok=True)
        TOKEN_RW.write_text(creds.to_json(), encoding="utf-8")
    return creds


def build_map_rows() -> List[List[Any]]:
    """FOOD（缶飯初期）MAP。継承=子非空→子、空→親。試行用。"""
    header = [
        "productType",
        "attrKey",
        "scHeaderJa",
        "scHeaderAlias",
        "required",
        "masterColPrimary",
        "masterColFallback",
        "inherit",
        "transform",
        "defaultValue",
        "doNotUse",
        "sourceNote",
        "notes",
        "enabled",
        "sampleFilledThisRun",
    ]
    # tuple fields aligned to header
    rows: List[List[Any]] = [
        # identity / structure
        ["FOOD", "sku", "SKU", "SKU|contribution_sku#1.value", "MUST", "子SKU", "親SKU", "CHILD_THEN_PARENT", "", "", "", "GENERATED sellerSku", "親は親SKU、子は子SKU", "TRUE", "YES"],
        ["FOOD", "product_type", "商品タイプ", "商品タイプ|product_type#1.value", "MUST", "Amazon Product Type", "", "CHILD_THEN_PARENT", "", "", "", "マスタ必須（空・既定禁止）", "C1: FOOD許可。空なら親除外", "TRUE", "YES"],
        ["FOOD", "action", "出品情報アクション", "出品情報アクション|::record_action", "MUST", "", "", "NO_INHERIT", "FIXED", "作成または置換 (完全更新)", "", "固定", "", "TRUE", "YES"],
        ["FOOD", "parentage", "親子レベル", "親子レベル", "MUST", "", "", "NO_INHERIT", "ROLE", "", "", "GENERATED variationRole", "親/子供", "TRUE", "YES"],
        ["FOOD", "parent_sku", "親SKU", "親SKU", "MUST", "親SKU", "", "CHILD_THEN_PARENT", "", "", "", "マスタ", "", "TRUE", "YES"],
        ["FOOD", "var_theme", "バリエーション テーマ", "バリエーション テーマ|variation_theme#1.name", "MUST", "", "", "NO_INHERIT", "FIXED", "サイズ", "", "固定（GROCERY純正プルダウン）", "缶飯サイズバリエ", "TRUE", "YES"],
        # names / copy
        ["FOOD", "title", "商品名", "商品名", "MUST", "最終商品名amazon", "商品名amazon|商品名案(Amazon)", "CHILD_THEN_PARENT", "TITLE_DEDUP_MAX75", "", "", "マスタ", "重複語禁止・ハイライト時75文字以内", "TRUE", "YES"],
        ["FOOD", "highlight", "商品のハイライト", "商品のハイライト", "SHOULD", "楽天キャッチコピーAIから取得", "Yahoo!キャッチコピーAIから取得|商品説明の箇条書き①", "CHILD_THEN_PARENT", "HIGHLIGHT_IF_TITLE_LE75", "", "", "マスタ優先B", "タイトル≤75: GV→Yahoo→箇条書き①。超なら空", "TRUE", "YES"],
        ["FOOD", "brand", "ブランド名", "ブランド名", "MUST", "", "", "NO_INHERIT", "FIXED", "ノーブランド品", "", "固定", "GTIN免除ノーブランド", "TRUE", "YES"],
        ["FOOD", "id_type", "商品IDの種類", "商品IDの種類", "MUST", "", "", "NO_INHERIT", "FIXED", "GTIN免除", "", "固定", "", "TRUE", "YES"],
        ["FOOD", "browse", "推奨されるブラウズノード", "推奨されるブラウズノード", "MUST", "Amazon Browse Node", "", "CHILD_THEN_PARENT", "", "", "", "マスタ必須（空・既定禁止）", "唐辛子既定埋込禁止。P4b取得後に必須", "TRUE", "YES"],
        ["FOOD", "mfr_name", "メーカー名", "メーカー名", "MUST", "メーカー名", "", "CHILD_THEN_PARENT", "", "", "", "マスタ", "", "TRUE", "YES"],
        # images
        ["FOOD", "main_image_url", "メイン画像のURL", "メイン画像のURL", "MUST", "Amazon MAIN URL", "", "CHILD_THEN_PARENT", "", "", "", "マスタ", "空フォールバック禁止", "TRUE", "YES"],
        ["FOOD", "other_image1", "その他の画像のURL1", "その他の画像のURL", "SHOULD", "Amazon PT URL", "", "CHILD_THEN_PARENT", "SPLIT_PT_URL_1", "", "", "マスタ PT URL |分割", "", "TRUE", "YES"],
        ["FOOD", "other_image2", "その他の画像のURL2", "その他の画像のURL", "OPT", "Amazon PT URL", "", "CHILD_THEN_PARENT", "SPLIT_PT_URL_2", "", "", "マスタ", "", "TRUE", "YES"],
        ["FOOD", "other_image3", "その他の画像のURL3", "その他の画像のURL", "OPT", "Amazon PT URL", "", "CHILD_THEN_PARENT", "SPLIT_PT_URL_3", "", "", "マスタ", "", "TRUE", "YES"],
        ["FOOD", "other_image4", "その他の画像のURL4", "その他の画像のURL", "OPT", "Amazon PT URL", "", "CHILD_THEN_PARENT", "SPLIT_PT_URL_4", "", "", "マスタ", "", "TRUE", "YES"],
        ["FOOD", "other_image5", "その他の画像のURL5", "その他の画像のURL", "OPT", "Amazon PT URL", "", "CHILD_THEN_PARENT", "SPLIT_PT_URL_5", "", "", "マスタ", "", "TRUE", "YES"],
        ["FOOD", "other_image6", "その他の画像のURL6", "その他の画像のURL", "OPT", "Amazon PT URL", "", "CHILD_THEN_PARENT", "SPLIT_PT_URL_6", "", "", "マスタ", "", "TRUE", "YES"],
        ["FOOD", "other_image7", "その他の画像のURL7", "その他の画像のURL", "OPT", "Amazon PT URL", "", "CHILD_THEN_PARENT", "SPLIT_PT_URL_7", "", "", "マスタ", "", "TRUE", "YES"],
        ["FOOD", "other_image8", "その他の画像のURL8", "その他の画像のURL", "OPT", "Amazon PT URL", "", "CHILD_THEN_PARENT", "SPLIT_PT_URL_8", "", "", "マスタ", "", "TRUE", "YES"],
        # description
        ["FOOD", "desc", "製品の説明", "製品の説明", "SHOULD", "商品説明の箇条書き①", "", "CHILD_THEN_PARENT", "", "", "", "マスタ", "", "TRUE", "YES"],
        ["FOOD", "spec1", "箇条書き1", "箇条書き", "MUST", "商品説明の箇条書き①", "", "CHILD_THEN_PARENT", "", "", "", "マスタ", "", "TRUE", "YES"],
        ["FOOD", "spec2", "箇条書き2", "箇条書き", "SHOULD", "商品説明の箇条書き②", "商品説明の箇条書き①", "CHILD_THEN_PARENT", "", "", "", "マスタ", "空なら①流用（現行C1）", "TRUE", "YES"],
        ["FOOD", "spec3", "箇条書き3", "箇条書き", "SHOULD", "商品説明の箇条書き③", "商品説明の箇条書き①", "CHILD_THEN_PARENT", "", "", "", "マスタ", "", "TRUE", "YES"],
        ["FOOD", "spec4", "箇条書き4", "箇条書き", "SHOULD", "商品説明の箇条書き④", "商品説明の箇条書き①", "CHILD_THEN_PARENT", "", "", "", "マスタ", "", "TRUE", "YES"],
        ["FOOD", "spec5", "箇条書き5", "箇条書き", "SHOULD", "商品説明の箇条書き⑤", "商品説明の箇条書き①", "CHILD_THEN_PARENT", "", "", "", "マスタ", "", "TRUE", "YES"],
        ["FOOD", "keyword1", "検索用キーワード", "検索用キーワード", "SHOULD", "検索キーワード", "★Amazon検索KW(150字)|▼マスタ(★Amazon検索KW(150字))", "CHILD_THEN_PARENT", "KW_JOIN_1SLOT", "", "", "マスタ", "1枠のみ（99016対策）", "TRUE", "YES"],
        # quantity / identity attrs (問題の4項目)
        ["FOOD", "number_items", "商品の入数", "商品の入数|number_of_items", "MUST", "A.セット商品数", "setCount", "CHILD_ONLY", "PARSE_SET_COUNT", "", "▼マスタ(総個数)|商品の入数(親総量)", "マスタ／GENERATED", "親総個数を使わない。子のセット数", "TRUE", "YES"],
        ["FOOD", "color", "色", "色|カラー", "OPT", "カラー", "色", "CHILD_THEN_PARENT", "", "", "", "マスタ", "GROCERYテーマ=サイズ時は未出力（その他固定禁止）", "TRUE", "YES"],
        ["FOOD", "mfr_part", "メーカー型番", "メーカー型番", "MUST", "メーカー型番", "メーカー品番|型番", "CHILD_THEN_PARENT", "FALLBACK_CHILD_SKU", "", "", "マスタHJ→GENERATED", "HJ優先。空なら品番→GENERATED→子SKU", "TRUE", "YES"],
        ["FOOD", "import_type", "輸入種別", "輸入種別", "SHOULD", "", "", "NO_INHERIT", "FIXED", "正規品", "", "固定", "", "TRUE", "YES"],
        ["FOOD", "exclusive", "Amazon.co.jp限定商品ですか？", "この商品はAmazon.co.jp限定商品ですか？", "SHOULD", "", "", "NO_INHERIT", "FIXED", "いいえ", "", "固定", "", "TRUE", "YES"],
        ["FOOD", "heat", "商品は感熱性ですか？", "商品は感熱性ですか？|▼マスタ(Amz:感熱性)", "SHOULD", "商品は感熱性ですか？", "▼マスタ(Amz:感熱性)", "CHILD_THEN_PARENT", "YES_NO_JP", "いいえ", "", "マスタ", "", "TRUE", "YES"],
        ["FOOD", "temperature_rating", "温度の定格", "温度の定格", "OPT", "温度の定格", "", "NO_INHERIT", "", "常温：室温", "保存方法(食品)", "マスタ／既定", "GROCERY黒セル時は未出力。保存方法長文を定格に載せない", "TRUE", "YES"],
        ["FOOD", "ingredients", "原料", "原料|特記すべき原材料／原料|原材料(食品)|▼マスタ(原材料(食品))", "MUST", "特記すべき原材料／原料", "原材料(食品)|(原材料(食品))|▼マスタ(原材料(食品))", "CHILD_THEN_PARENT", "", "", "", "マスタ", "", "TRUE", "YES"],
        ["FOOD", "item_form", "商品の形式", "商品の形式|商品形態", "OPT", "商品の形式", "", "NO_INHERIT", "", "", "", "マスタ", "GROCERY黒セル時は未出力（ホール固定禁止）", "TRUE", "YES"],
        ["FOOD", "unit_count", "ユニット数", "ユニット数|Amazonユニット数量", "MUST", "ユニット数", "Amazonユニット数量", "CHILD_ONLY", "USE_SET_COUNT", "", "一人分の数量|▼マスタ(総個数)", "マスタ", "缶数。一人分の数量(160)を使わない", "TRUE", "YES"],
        ["FOOD", "unit_uom", "商品のユニット数の単位", "商品のユニット数の単位|Amazon一人分単位|▼マスタ(Amz:ユニット数単位)", "MUST", "商品のユニット数の単位", "▼マスタ(Amz:ユニット数単位)", "CHILD_THEN_PARENT", "", "缶", "", "マスタ", "", "TRUE", "YES"],
        ["FOOD", "size", "サイズ", "サイズ|size#1.value", "MUST", "バリエーション値", "A.セット商品数", "CHILD_ONLY", "", "", "パッケージサイズ名", "マスタ", "テーマ=サイズ時はATのsize属性。package_size_nameではない", "TRUE", "YES"],
        ["FOOD", "item_weight", "商品の重量", "商品の重量", "MUST", "商品の重量", "", "CHILD_ONLY", "PARSE_WEIGHT_FROM_SIZE", "", "▼マスタ(総重量)", "マスタ／サイズから", "親総重量960を継承しない", "TRUE", "YES"],
        ["FOOD", "item_weight_unit", "商品の重量の単位", "商品の重量の単位", "MUST", "商品の重量の単位", "", "CHILD_THEN_PARENT", "", "グラム", "", "マスタ／既定", "", "TRUE", "YES"],
        # offer
        ["FOOD", "condition", "商品の状態", "商品の状態", "MUST", "", "", "NO_INHERIT", "FIXED", "新品", "", "固定", "", "TRUE", "YES"],
        ["FOOD", "list_price", "税込みの参考価格", "税込みの参考価格|定価、市場価格|定価", "MUST", "定価、市場価格", "定価|市場価格", "CHILD_THEN_PARENT", "DIGITS_ONLY", "", "", "マスタ", "親の非数字は継承しない", "TRUE", "YES"],
        ["FOOD", "tax_code", "商品タックスコード", "商品タックスコード", "MUST", "商品タックスコード", "", "CHILD_THEN_PARENT", "", "", "", "マスタ", "", "TRUE", "YES"],
        ["FOOD", "fulfillment", "フルフィルメントチャネルコード", "フルフィルメントチャネルコード (JP)", "MUST", "", "", "NO_INHERIT", "FIXED", "出品者出荷（デフォルト）", "", "固定", "", "TRUE", "YES"],
        ["FOOD", "inventory", "在庫数", "在庫数 (JP)", "MUST", "在庫数", "", "CHILD_THEN_PARENT", "", "0", "", "GENERATED／試験ZERO", "試験は0", "TRUE", "YES"],
        ["FOOD", "price", "商品の販売価格", "商品の販売価格 JPY (Amazonで販売, JP)|販売価格amazon", "MUST", "販売価格amazon", "", "CHILD_THEN_PARENT", "", "", "", "GENERATED／マスタ", "", "TRUE", "YES"],
        ["FOOD", "shipping", "配送テンプレート", "配送テンプレート (JP)", "MUST", "", "shippingTemplate", "CHILD_THEN_PARENT", "", "送料無料パターン", "", "GENERATED／既定", "", "TRUE", "YES"],
        ["FOOD", "origin", "原産国", "原産国|原産国/地域|▼マスタ(原産国)", "MUST", "原産国/地域", "原産国|▼マスタ(原産国)", "CHILD_THEN_PARENT", "", "", "", "マスタ", "", "TRUE", "YES"],
        ["FOOD", "hazmat", "危険物規制の種類", "商品に適用される危険物規制の種類|▼マスタ(Amz:危険物規制)", "SHOULD", "▼マスタ(Amz:危険物規制)", "", "CHILD_THEN_PARENT", "", "該当なし", "", "マスタ／既定", "", "TRUE", "YES"],
        ["FOOD", "liquid", "液体物は含まれていますか？", "液体物は含まれていますか？|液体物含有|▼マスタ(Amz:液体物含有)", "SHOULD", "液体物含有", "▼マスタ(Amz:液体物含有)", "CHILD_THEN_PARENT", "YES_NO_JP", "いいえ", "", "マスタ", "", "TRUE", "YES"],
        # optional master basics not filled this run but keep for binding
        ["FOOD", "flavor", "味", "味", "OPT", "味", "", "CHILD_THEN_PARENT", "", "", "", "マスタ", "今回未出力", "TRUE", "NO"],
        ["FOOD", "expiration_dated", "商品に有効期限はありますか?", "商品に有効期限はありますか?", "OPT", "商品に有効期限はありますか?", "", "CHILD_THEN_PARENT", "", "", "", "マスタ", "今回未出力", "TRUE", "NO"],
        ["FOOD", "expiration_type", "商品の有効期限タイプ", "商品の有効期限タイプ", "OPT", "商品の有効期限タイプ", "", "CHILD_THEN_PARENT", "", "", "", "マスタ", "今回未出力", "TRUE", "NO"],
        ["FOOD", "shelf_life", "倉庫の保存可能期間", "倉庫の保存可能期間", "OPT", "倉庫の保存可能期間", "", "CHILD_THEN_PARENT", "", "", "", "マスタ", "今回未出力", "TRUE", "NO"],
        ["FOOD", "shelf_life_unit", "FC納品時残存賞味期限", "FC納品時残存賞味期限", "OPT", "FC納品時残存賞味期限", "", "CHILD_THEN_PARENT", "", "", "", "マスタ", "今回未出力", "TRUE", "NO"],
        ["FOOD", "grind_type", "グラインドタイプ", "グラインドタイプ", "OPT", "グラインドタイプ", "", "CHILD_THEN_PARENT", "", "", "", "マスタ", "調味料向け・缶飯は通常空", "TRUE", "NO"],
        ["FOOD", "package_length", "梱包:奥(cm)", "梱包:奥(cm)|▼マスタ(梱包:奥(cm))", "OPT", "梱包:奥(cm)", "▼マスタ(梱包:奥(cm))", "CHILD_THEN_PARENT", "", "", "", "マスタ", "今回未出力", "TRUE", "NO"],
        ["FOOD", "package_width", "梱包:幅(cm)", "梱包:幅(cm)|▼マスタ(梱包:幅(cm))", "OPT", "梱包:幅(cm)", "▼マスタ(梱包:幅(cm))", "CHILD_THEN_PARENT", "", "", "", "マスタ", "今回未出力", "TRUE", "NO"],
        ["FOOD", "package_height", "梱包:高(cm)", "梱包:高(cm)|▼マスタ(梱包:高(cm))", "OPT", "梱包:高(cm)", "▼マスタ(梱包:高(cm))", "CHILD_THEN_PARENT", "", "", "", "マスタ", "今回未出力", "TRUE", "NO"],
        ["FOOD", "package_weight", "梱包:重量(g)", "梱包:重量(g)|▼マスタ(梱包:重量(g))", "OPT", "梱包:重量(g)", "▼マスタ(梱包:重量(g))", "CHILD_THEN_PARENT", "", "", "", "マスタ", "今回未出力", "TRUE", "NO"],
        ["FOOD", "set_count_ref", "（参考）A.セット商品数", "A.セット商品数", "OPT", "A.セット商品数", "", "CHILD_ONLY", "PARSE_SET_COUNT", "", "", "マスタ", "入数・ユニットの元データ", "TRUE", "REF"],
        ["FOOD", "master_total_qty_ref", "（参考・禁止）総個数", "総個数|▼マスタ(総個数)", "OPT", "▼マスタ(総個数)", "総個数", "PARENT_ONLY", "", "", "", "マスタ親", "Amazon入数に使わない", "TRUE", "REF"],
        ["FOOD", "master_total_weight_ref", "（参考・禁止）総重量", "総重量|▼マスタ(総重量)", "OPT", "▼マスタ(総重量)", "総重量", "PARENT_ONLY", "", "", "", "マスタ親", "Amazon重量に使わない", "TRUE", "REF"],
        ["FOOD", "serving_qty_ref", "（参考・禁止）一人分の数量", "一人分の数量", "OPT", "一人分の数量", "", "CHILD_THEN_PARENT", "", "", "", "マスタ", "ユニット数に使わない（160g誤用）", "TRUE", "REF"],
    ]
    return [header] + rows


def build_rules_block() -> List[List[Any]]:
    return [
        [],
        ["=== RULES（継承・変換の正本・試行） ==="],
        ["ruleId", "意味"],
        ["CHILD_THEN_PARENT", "子SKU行に値あり→子。空→親。両方空→defaultValue。なお必須でdefaultも空ならエラー"],
        ["CHILD_ONLY", "子のみ参照。親の総量などをバリエ子へ流さない（入数・重量・サイズ）"],
        ["PARENT_ONLY", "親のみ（参考・禁止列の説明用）"],
        ["NO_INHERIT", "固定／GENERATED役割。マスタ継承しない"],
        ["PARSE_SET_COUNT", "A.セット商品数「3個で1セット」→3。なければGENERATED setCount"],
        ["PARSE_WEIGHT_FROM_SIZE", "バリエーション値「3缶/480g」→480"],
        ["USE_SET_COUNT", "ユニット数＝缶数（PARSE_SET_COUNTと同値）"],
        ["TITLE_DEDUP_MAX75", "商品名の重複語削除。ハイライト使用時は75文字以内"],
        ["HIGHLIGHT_IF_TITLE_LE75", "タイトルが75超ならハイライトを出さない（100476）"],
        ["KW_JOIN_1SLOT", "検索KWは1枠に空白結合（99016）"],
        ["FALLBACK_CHILD_SKU", "メーカー型番／品番空→GENERATED→子SKU"],
        ["HIGHLIGHT_PRIORITY_B", "タイトル≤75: 楽天キャッチ→Yahooキャッチ→箇条書き①。超なら空"],
        ["YES_NO_JP", "はい／いいえ正規化"],
        ["FIXED", "defaultValueをそのまま書く"],
        ["ROLE", "親行=親、子行=子供"],
        [],
        ["=== 共通認識 ==="],
        ["1", "マスタで子に値がある項目は子を使う。子が空なら親を使う（inherit=CHILD_THEN_PARENT）"],
        ["2", "バリエ数量・重量・サイズはCHILD_ONLY（親総個数／総重量を流用しない）"],
        ["3", "本シートは試行用。C1実行エンジンはまだjson優先。シート確定後にコード接続予定"],
        ["4", "列番号は持たない。scHeaderAliasで項目名解決する"],
        ["5", "Yahooマッピングと違い縦型・PT単位・UL必要項目中心"],
    ]


def build_errors_block() -> List[List[Any]]:
    return [
        [],
        ["=== ERRORS（今回缶飯 CK_5beb0cbf67ea_B1 初期） ==="],
        [
            "errorCode",
            "productType",
            "count",
            "symptom",
            "rootCause",
            "fixMaster",
            "fixMap",
            "status",
            "sampleFile",
        ],
        [
            "18367",
            "FOOD",
            "13",
            "PTがSEASONING→FOODに更新された",
            "缶飯なのにSEASONINGで送った",
            "Amazon Product Type=FOOD",
            "product_type default=FOOD／browse見直し",
            "OPEN",
            "CK_5beb0cbf67ea_B1_…-processing-summary.xlsm",
        ],
        [
            "13013",
            "FOOD",
            "13",
            "カタログ未作成で出品情報を付けられない",
            "親エラー波及",
            "親の商品名・PT・数量を先に直す",
            "親MUST行を優先修正",
            "OPEN",
            "同上",
        ],
        [
            "100730",
            "FOOD",
            "11",
            "同一詳細の重複投稿",
            "処理中の再送／並行処理",
            "20分以上空けて再送",
            "",
            "WAIT",
            "同上",
        ],
        [
            "100470",
            "FOOD",
            "2",
            "商品名で保存・缶詰が二重",
            "最終商品名amazonの重複語",
            "商品名から重複削除",
            "TITLE_DEDUP",
            "OPEN",
            "同上",
        ],
        [
            "100476",
            "FOOD",
            "2",
            "ハイライト利用は商品名75文字以内",
            "タイトル86文字＋ハイライトあり",
            "タイトル短縮 or ハイライト空",
            "HIGHLIGHT_IF_TITLE_LE75",
            "OPEN",
            "同上",
        ],
        [
            "20014",
            "FOOD",
            "1",
            "メディア無効／破損（96s14）",
            "画像URL問題の疑い",
            "Amazon MAIN URL確認",
            "",
            "OPEN",
            "同上",
        ],
        [
            "8007",
            "FOOD",
            "1",
            "親SKUの問題で子を処理できない",
            "親エラー",
            "親を先に修正",
            "",
            "OPEN",
            "同上",
        ],
        [
            "(human)",
            "FOOD",
            "-",
            "入数全6／ユニット160／重量960／ATサイズ空",
            "親総量＋一人分誤マップ＋size列取り違え",
            "c1_quantity_policy＋size=AT属性",
            "c1_packaged/food_fish_grocery_column_map",
            "FIXED",
            "2026-08-02 実装",
        ],
    ]


def main() -> int:
    cfg = load_cfg()
    sid = cfg["spreadsheet_id"]
    creds = get_creds()
    svc = build("sheets", "v4", credentials=creds)

    meta = (
        svc.spreadsheets()
        .get(spreadsheetId=sid, fields="sheets(properties(sheetId,title,index))")
        .execute()
    )
    yahoo_index = None
    existing_id = None
    for sh in meta.get("sheets") or []:
        p = sh.get("properties") or {}
        title = p.get("title") or ""
        if title == YAHOO_TITLE:
            yahoo_index = int(p.get("index"))
        if title == SHEET_TITLE:
            existing_id = int(p.get("sheetId"))

    if existing_id is not None:
        print("既存シートあり sheetId=%s → 内容をクリアして再投入" % existing_id)
        # clear
        svc.spreadsheets().values().clear(
            spreadsheetId=sid, range="'%s'" % SHEET_TITLE
        ).execute()
        sheet_id = existing_id
        # move next to yahoo if needed
        if yahoo_index is not None:
            svc.spreadsheets().batchUpdate(
                spreadsheetId=sid,
                body={
                    "requests": [
                        {
                            "updateSheetProperties": {
                                "properties": {
                                    "sheetId": sheet_id,
                                    "index": yahoo_index + 1,
                                },
                                "fields": "index",
                            }
                        }
                    ]
                },
            ).execute()
    else:
        insert_at = (yahoo_index + 1) if yahoo_index is not None else 0
        print("新規作成 index=%s (Yahoo右隣)" % insert_at)
        resp = (
            svc.spreadsheets()
            .batchUpdate(
                spreadsheetId=sid,
                body={
                    "requests": [
                        {
                            "addSheet": {
                                "properties": {
                                    "title": SHEET_TITLE,
                                    "index": insert_at,
                                    "gridProperties": {
                                        "rowCount": 200,
                                        "columnCount": 16,
                                    },
                                }
                            }
                        }
                    ]
                },
            )
            .execute()
        )
        sheet_id = resp["replies"][0]["addSheet"]["properties"]["sheetId"]

    values: List[List[Any]] = []
    values.append(["▼設定(Amazonマッピング) — 試行版", "缶飯 FOOD 初期投入", "subBatchId=CK_5beb0cbf67ea_B1", "継承: 子→親", "C1未接続"])
    values.append(
        [
            "使い方",
            "MAP表を直しながら試す。enabled=FALSEで無効。C1は当面 food_seasoning_column_map.json が正。本シート確定後に接続。",
        ]
    )
    values.append([])
    values.append(["=== MAP ==="])
    values.extend(build_map_rows())
    values.extend(build_rules_block())
    values.extend(build_errors_block())

    svc.spreadsheets().values().update(
        spreadsheetId=sid,
        range="'%s'!A1" % SHEET_TITLE,
        valueInputOption="RAW",
        body={"values": values},
    ).execute()

    # freeze header-ish: freeze first 5 rows, bold via basic formatting optional skip
    svc.spreadsheets().batchUpdate(
        spreadsheetId=sid,
        body={
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": sheet_id,
                            "gridProperties": {"frozenRowCount": 5},
                        },
                        "fields": "gridProperties.frozenRowCount",
                    }
                }
            ]
        },
    ).execute()

    print("DONE sheet=%s sheetId=%s rows=%d" % (SHEET_TITLE, sheet_id, len(values)))
    print(
        "URL: https://docs.google.com/spreadsheets/d/%s/edit#gid=%s" % (sid, sheet_id)
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
