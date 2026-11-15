# プロジェクト全体の位置づけと現在の開発フォーカス

**最終更新**: 2026-11-15（バックアップ Phase0: G1＋G3 タスク登録／盛り共通核は並行）
**読み方**: 次の Agent は `docs/AGENT_HANDOVER.md` の **§1.5・§2** に従い、**本ファイルを最初に読み**、続けて §2 の必読一覧でプロジェクト全体をインプットする。

---

## 0backup. 全PJバックアップ（2026-11-15）

**方針**: スプシ＝Drive日次。コード＝**G1（毎朝8:05 main push）＋G3（全claspローカル）**。次段スプシ順＝広告→在庫→競合Jitta→税務。[要件](org/PROJECT_DAILY_BACKUP_REQUIREMENTS.md)／[HUMAN_RUN](org/PROJECT_DAILY_BACKUP_HUMAN_RUN.md)

**Phase0 済**: `.gitignore` 強化。G1/G3 タスク登録。G3初回OK。GitHub baseline push 済。**Drive目視OK（2026-11-15）**: `出品_backup_2026-11-15`／`競合DB_backup_2026-11-15` あり（①調査は11/14が最新＝更新なしスキップ想定）。

**次**: 次段スプシ日次は**広告**から横展開。

---

## 0plating_generic. 盛り共通核（2026-11-15）

**採用**: 一片見本・MAIN盛り・SUB盛りで共通レイヤ（危険文オフ・面クロップ・スロット分離・qualityはできる限り low）。正本 [org/PLATING_AND_PIECE_GENERIC_LOGIC.md](org/PLATING_AND_PIECE_GENERIC_LOGIC.md)。コード `b5_generic_plating.py`。切戻し `B5_GENERIC_PLATING_CORE=0`。  
**次**: ぶり等で MAIN/SUB 検収。high は既定オフ。

---

## 0session. 全体要件確認セッション終了（2026-10-30）

**このセッションでやったこと**: プロジェクト全体像の再確認、進捗整理、エージェント引き継ぎ方針の確認。プロジェクト全体の要件定義・前提は **gas-project の docs が正本**。他エージェントはフローごとに実装要件を詰め、新要件・前提は docs に戻す（[AGENT_HANDOVER.md](AGENT_HANDOVER.md)）。

**いまの位置づけ**

| 項目 | 状態 |
|------|------|
| M1 基盤 | ほぼ済 |
| M2 出品（領域2） | 進行中・主戦場（下記 0yahoo_caption / C→D→E 等） |
| M3 領域1 | 進行中・別本線 |
| M4〜M6 | 後続 |

**次にやること（新Agent）**: **フローごとの実装要件定義**を詰める。起動時は AGENT_HANDOVER §2 をすべてインプットし、担当フローを明示する。新たに決めた要件・前提は該当 docs に追記する。

**このセッションの区切り**: 全体要件の確認はここまで。実装の要件定義は**別エージェント**で進める。

---

## 0yahoo_caption. Yahoo商品説明テンプレ＋LOCK（2026-09-14）

**セッション入口（新Agent）**: [SESSION_HANDOVER_YAHOO_CAPTION_20260915.md](SESSION_HANDOVER_YAHOO_CAPTION_20260915.md)

**本線は案A'（自動化）**: P6b lib API UP済。本採用候補へ差し替え再UP済（`sub_image_ai_compose`）。Proで魚種画像目視待ち。

**完成形（P6c-1 当選限定・紹介組立）**: P5固定文は仮本線。本文＝マスタ **`楽天サブ画像1〜10`** の当選サブ。紹介は catch＋salesTalk／subs＋魚種・主原料から機械組立。内容量の数字は紹介に使わない（大容量は残す）。写真用overviewと商品説明文AIは種にしない。画と文は同一 themeId。上位3。PC・スマホは同じ仕様で枠内。ONは `YAHOO_CAPTION_RICH_P6C_THEMES_ENABLED=true`。[要件 §0.1](org/YAHOO_CAPTION_TEMPLATE_LOCK_REQUIREMENTS.md)／[P6c](../tools/yahoo_page_mock/paste/HUMAN_RUN_YAHOO_CAPTION_P6C_THEMES.md)。

**今の人間手順**: マッピング `point-immediate`＝1 済。かつお・まぐろ・ぶりは説明API反映を **1回**（昨夜は 0 のまま `updateItem`）。**駐車**: 出品済みの任意項目だけ直す器（`getItem`→現行値同送→`editItem`）。大きいので未着手。[SESSION §5 YAHOO-PARTIAL-UPDATE](SESSION_HANDOVER_YAHOO_CAPTION_20260915.md)。

**AI-DESC-FISH-MIX（2026-09-17 コード済）**: Yahoo商品情報（explanation）の複合魚種は caption テンプレではなく AI→マスタJOIN。一意ASIN→一意商品名ベース＝AI商品名→一意JAN。メーカー両方非空なら一致必須。ブランドはJOINしない。生成プロンプトは1SKU固定。未設定=ON。切戻し `AI_DESC_FISH_MIX_JOIN_ENABLED=false`。要出品 clasp→おだしAI再生成→マスタ同期→説明API。[要件 §6.1](org/YAHOO_CAPTION_TEMPLATE_LOCK_REQUIREMENTS.md)。

---

**C**＝楽天画像マッチング（02→シート／金丸やり直しまれ）。**D**＝楽天・Yahoo出品のみ（Amazonラジオ削除）。**E**＝Amazon（07マッチング→R2→出品ファイル）。**E①**は楽天未整理を読まない（一致=楽天欄保持／レ点差分=重なりSKUだけ引継ぎ）。切戻し `AMAZON_E1_KEEP_RAKUTEN_ENABLED=false`。**E②/U4**は未充足SKUを全件キューし、E②開始から22分で切って機械再開（人が同じメニューを件数分割で押さない）。**D補助.旧E**はトップ廃止→Z「E段階・互換」。日常: `C→D→E`。正本: [org/D_MENU_CDE_MENU_SPLIT_REQUIREMENTS.md](org/D_MENU_CDE_MENU_SPLIT_REQUIREMENTS.md)。要 clasp。

**Cレイアウト拡張（同日）**: メイン2〜10＝他セット昇順、サブ＝全子共通、親＝最少セット子。青枠へ直接 IMAGE。[org/D_MENU_RAKUTEN_MATRIX_LAYOUT_REQUIREMENTS.md](org/D_MENU_RAKUTEN_MATRIX_LAYOUT_REQUIREMENTS.md)。戻し `RAKUTEN_MATRIX_LAYOUT_EXPAND_ENABLED=false`。

---

## 0b45. MAIN／サブ整合（2026-09-05）

**表示**: **B-④＝楽天MAIN基礎**／**B-⑤＝サブ画像**／**B-⑥＝Amazon MAIN**（関数名互換）。推奨順: A→B（**B一括＋B-③は起動確認・未設定OFF・案1コード済**）→正本→競合→**個品見本**→**B-④／B-⑤順不同**（文字色・正本・見本は親共有）→待ちに **B-⑥**→**C→D→E**（楽天マッチング→楽天Yahooアップ→Amazon）。正本: [ALIGN §1／§1.2](org/B45_MAIN_SUB_ALIGN_REQUIREMENTS.md)／[CDE分割](org/D_MENU_CDE_MENU_SPLIT_REQUIREMENTS.md)／[B6要件](org/D_MENU_B6_AMAZON_MAIN_REQUIREMENTS.md)。**§7.14** 個品共有継続。**文字色共有コード済**（Prop＋Drive JSON・要 clasp）。

**MAIN→サブ寄せ（2026-11-02）**: [ALIGN §4.5](org/B45_MAIN_SUB_ALIGN_REQUIREMENTS.md)。**P1人間検収OK**。**P2閉じた**。**P3第1配線**（§4.5.10 G7/G6/G1→`dry_sub`。切戻し `B4_SUB_PICTURE_GATE=0`）。**P3-W第1**（§4.5.11 WRAP→個包装訴求 themeId=21 +1。hasWrapは親SKU→商品名ベース→JAN。切戻し `B4_WRAP_APPEAL_THEME=false`。要 clasp）。**P3-W第2**（手持ち候補UI。切戻し `B4_WRAP_ASSET_LANE=0`。clasp不要）。次＝④MAIN左下Pillow。

**薄い一片あり（2026-11-10）**: [ALIGN §4.5.10bis-T](org/B45_MAIN_SUB_ALIGN_REQUIREMENTS.md)／[T2](org/B45_MAIN_SUB_ALIGN_REQUIREMENTS.md)／[T3](org/B45_MAIN_SUB_ALIGN_REQUIREMENTS.md)／[T4](org/B45_MAIN_SUB_ALIGN_REQUIREMENTS.md)。F0〜F2済＋IDENTITY≠一片。**T4配線済**（深堀UI・スタック・形状ドナー自動。仮想B/C合格）。切戻し --piece-set-max 0／B5_PIECE_TRUTH_GATE=0／B5_PIECE_GATE_UI=0／B5_SURFACE_SHAPE_SPLIT=0。**いま＝F3ぶり人間検収**（P3 MAIN2+SUB1生成済・3_review.html）。ミックスはF4。
---

## 0b5rak. 楽天MAIN基礎（表示B-④／旧B-⑤・2026-08-28）

**P0・案F本線ロック**（2026-08-29）。**P1〜P3b コード済**。**§0.2 役割定義ロック**（2026-10-10。入口ゲート廃止・大＋小焼成）。**案A親キュー＋公式URLキャッシュ**（2026-10-11。1窓・親ごと完了→次・先読み。TTL30日。要出品 clasp）。**S15〜S19 virt揃え GAS済**（ペルソナ詳細・場面Gemini・公式料理。要出品 clasp）。**後載せは不合格で既定オフ**（2026-09-20。袋はAI配置・縦横比固定。個品は見本を貼らない。短冊許可は削除）。ブラウザで複数採用＋レ点子SKU割当→`_単体__{adoptKey}.png`＋`set_base_map.json`（**textColor**）。Cは子ごとに土台＋親ごと文字色。Yahoo同エンジン／Amazon対象外。既存再利用（`--list-latest`）。拡大・再生成世代保持。要 clasp＋人間検収。正本: [org/D_MENU_B5_RAKUTEN_MAIN_BASE_REQUIREMENTS.md](org/D_MENU_B5_RAKUTEN_MAIN_BASE_REQUIREMENTS.md) §0.2・§3.6.1〜3.6.3／[HUMAN_RUN](org/D_MENU_B5_RAKUTEN_MAIN_BASE_HUMAN_RUN.md)／[ALIGN](org/B45_MAIN_SUB_ALIGN_REQUIREMENTS.md)。

---

## 0b5. A/B短縮 改修連番（2026-08-16）

| # | 内容 | 状態 |
|---|------|------|
| R0 | 行高さ Step8 廃止 | 済 |
| R1 | A.準備 SEOスペース | コード済。人間: 3 dry→4 書込（B前） |
| R2 | A Gemini画像照合OFF | Property済 `KEEPA_SKIP_REF_IMAGE_AND_MATCH=true` |
| R3 | Step5 OpenAI は Gemini失敗時のみ | 済。未設定=ON。切戻し `B_STEP5_SKIP_OPENAI_IF_GEMINI_OK=false` |
| R4 | Step5 ジャンルAI停止 | 既存スキップ維持。`B_STEP5_SKIP_GENRE_YAHOO_AI_ENABLED` を false にしない |
| R5 | Step5 AI梱包・配送・根拠停止 | 済。未設定=ON。切戻し `B_STEP5_SKIP_PACK_AI=false`。正は Keepa／6.55 |
| R6 | KW／説明 vs 7 vs 7.6 | 現状維持 |
| R7 | B 2.1＝12-⑭ | 済 |
| R8 | A.準備と 12-⑮⑯ | 改修しない |
| R9 | キュー3→4 | **しない**（R3/R5の実測後に別承認） |
| R10 | 楽天CSV／Yahoo.js／B順 | 触らない |
| R11 | 処理時間 | 済。シート `▼実行時間`。配線は B Step と Step5 の geminiSearch／geminiVision／openaiVision のみ。共通 `logOpTiming_`。切戻し `OP_TIMING_ENABLED=false`。**全関数 `console.time` 禁止**（再指示時は `.cursor/rules/op-timing-scope.mdc` で止める） |
| R12 | 異種組み合わせセット | Keepa◎がある袋はAmazonその値段。寸法は長辺・短辺最大、最短のみ数量倍。ミックス親FBAは1袋子ASIN＋Catalog梱包空ならBC〜BE。[org/COMBO_SET_BOM_REQUIREMENTS.md](org/COMBO_SET_BOM_REQUIREMENTS.md) §6–§7。要 clasp |

---

## 0props. Script Properties 上限（2026-09-17）

1件9KB・合計500KB。未使用フラグより **B5正本／文字色／Amazon PT指紋の巨大JSON** が主因。秘密（APIキー）はシートに出さない・消さない。

**人間**: 出品 clasp → Z「15. 共通・検証」→「15f. Script Properties」→ ㉓使用量 → ㉔巨大JSON退避。任意で㉕個別フラグ削除。シート名 `▼ScriptPropertiesメモ`。Drive フォルダ `ScriptProps退避`。A〜Eの読み取りは退避後も Drive を見る。

---

## 0ss. 共通スプシ参照（2026-08-19）

人間がログ・スプシ・実行結果の確認や合否を求めたときの正本: [org/COMMON_SS_VIEW_REQUIREMENTS.md](org/COMMON_SS_VIEW_REQUIREMENTS.md)。手元は `tools/common_ss_view/`（`dump.py` → `out/`）。IDは全PJ固定。メニュー単位の最小ファイル＋関係タブだけ。回答先頭に見た場所。ルール `.cursor/rules/common-ss-view.mdc`。

---

## 0comp. 競合フィールド正本（2026-08-15）

出品A/B・領域1リサーチ・定時の競合チェック／価格対抗で **同じ項目辞書**を使う。  
正本: [org/COMPETITOR_FIELDS.md](org/COMPETITOR_FIELDS.md)（CSV は `docs/org/competitor_fields/`）。組み立て時は §1 のフロー正本も開く。1件目＝競合価格は禁止。  
**専用ストア**: H1合格。**12-⑭** を B Step2.1 に組み込み（テスト）。`COMPETITOR_STORE_ENABLED=true` かつ `COMPETITOR_SS_ID` 必須。切戻し `B_COMPETITOR_STORE_APPLY_ENABLED=false`（旧横断シート直書き）。要 clasp。15〜23はA／人手／dryのまま。

---

## 0ab. A→B検収メモ（2026-08-22・実データ確認済）

runId `314f7d23-…`／7.6 `AI8_20260822_005631_717319`。AL「販売価格」は未使用（正は `販売価格amazon`）。

| 状態 | 内容 |
|------|------|
| 許容 | チーズ系ミックスの寸法3辺・FBA空／楽天Yahooセット欠け（競合未出品） |
| **保留（G7）** | **ミックス楽天 7207／同一URL複数親**（例 `lifucoco-shop/0764-012590`）— **楽天に当該セットの正しい競合が無いとき**に広い検索が汎用ヒットを拾う程度。**クリティカルではない**。問題が大きくならない限りコード改修しない。再検討条件＝正しいセット競合があるのに誤URLが載る／複数親汚染が価格提案を壊す。詳細 [B_RERUN_CLEAR_AND_JUDGE](org/B_RERUN_CLEAR_AND_JUDGE_HUMAN_RUN.md) §4.3 |
| 正常 | `timeBudgetStop`（約24.6分・Step7後→7.6手前）→再開で完了 |
| コード済・要clasp | **水色／sanky**: Step1後 `[SKU監査] ANOMALY`（接頭辞・親SKUずれ）で原因追跡可 |
| コード済・要clasp | **ミックスYahoo空**: `[ミックス競合] COMPARE` で楽天∥Yahoo並列＋Gemini byMall。`DUP_URL`で同一URL複数親 |
| コード済・要clasp | **Yahooカテゴリ**: 複数qマージ＋-score＋門点数制。仮想9PASS。要 clasp |
| **スプシ数式済** | **GB〜GD赤**: 原因＝最終名の「・」を重複語扱い。数式を `SUBSTITUTE(・)` 後にSPLITするよう全行更新済（29/58は「重複なし出品OK」確認） |
| 単品CQ空 | かつお／まぐろ／ぶり親の寸法なし（未着手） |
| **次（人間）** | **マスタ消去→再B**: 列チェックリスト＋合否判定 → [org/B_RERUN_CLEAR_AND_JUDGE_HUMAN_RUN.md](org/B_RERUN_CLEAR_AND_JUDGE_HUMAN_RUN.md) |

---

## 0amz. Amazon 競合貼付（2026-08-15・**保留**）

**A.準備**: Catalog空欄（1–2）＋商品名SEOスペース（3 dry／4 書込。AI情報取得data「商品名」のみ）。並べ替えはA末尾P5。要 clasp。シートを開き直す。

**残課題（着手しない）**: キャッシュ時 Keepaフル「目的=出品」／CAP／定時トリガー／空ASIN Catalog。`KEEPA_FULL_WRITE_ENABLED` は **false**。

**保留メモ（2026-08-21）**: キャッシュ時 Keepaフル「目的=出品」書込（`KEEPA_FULL_WRITE_ENABLED`）は、**リサーチPJが完了またはほぼ完了したタイミングで再提案**。書き先は競合DBの `Keepaフル`。いまは OFF のまま触らない。12-⑩定時対象JANログは在庫／競合価格リサーチ側で開発予定のため出品チャットでは保留。AのGemini画像照合OFFは**コストが上がり過ぎたとき再検討**（今は現行のまま）。

**後続・検証（保留）**: AのGemini画像照合を外すか。仮説=Catalog→Keepa商品名で足りる。切替 `KEEPA_SKIP_REF_IMAGE_AND_MATCH=true`。見るもの=◎の過不足・P5評点・袋数。**コスト過大時に再検討**（2026-08-21）。商品名掃除 `KEEPA_CLEAN_TITLE_BY_AI` は別。  
**済**: 12-⑭楽天Yahoo、P3 Amazon◎、P0–P5、W4読、P5 A末尾並べ替え。

**コードはフェーズ順・別承認。** 正本: [org/B_AMAZON_COMPETITOR_PASTE_REQUIREMENTS.md](org/B_AMAZON_COMPETITOR_PASTE_REQUIREMENTS.md)

- 領域1メーカー洗いではない。その商品の Catalog 階段（A′→B′→D′→E′。C′は診断）
- **Z併存**: `Keepaスナップショット` は薄い運用のまま。倉庫は `Keepaフル`＋生JSON
- **W4（予定）**: ①が書いた `Keepaフル` を90日内読む。洗い Catalog を貼付に流用しない。[TASK_STRUCTURE](org/B_PURCHASE_RESEARCH_TASK_STRUCTURE.md)
- 準備メニュー 12-⑮⑯ Catalog。**P2は12-⑰⑱**。**P3は12-⑲⑳**
- ④マスタは 12-⑭と同じ。Amazon 確定は人間◎のみ。今値 GET は後回し
- `KEEPA_PASTE_NEW_ASIN_CAP` 未設定=20／ブロック。token実測後に上限。トリガー未設置

**P3**: 12-⑳ 合格（`pasteP3-6dacd0c2` `written=20`）。人間は `AMAZON_PASTE_P3_WRITE_ENABLED` を **false に戻す**。

**Keepaフル（出品）**: W4 READ合格。P4 Python 1ASIN GET（非書）consumed=1 left=299。Aは offers=20 で2トークン想定。CAP=20据え置き未配線。WRITE OFF。

**WRITE計画 dry 合格（非書）**: 貼付60、フル200（目的リサーチ200・出品0）。貼付のうちフル既存39／未載21。現行フックは A の Keepa **raw があるときだけ**追記。キャッシュのみの A（api=0）は Property ON でも出品0行。ON は別承認。

**WRITE 実機（案A）合格 2026-08-15 22:09**: A20件 `[KeepaFullW4] take=20 hydrate=0 api=0 tokens=skip`。`[KeepaFull]` なし（raw無し）。`Keepaフル` 220行すべて目的=リサーチ、**出品0**。Property を **false に戻す**。

**P5 実機合格 2026-08-15**: `pasteP5-028e317c` DONE ok=true。P2 `pasteP2-554ce90d` write=true。シート tagged 全ASIN・stack_ok。A api=0。WRITE は false のまま。

**貼付P0–P5完了。** 案Bしない。次の出品コードはキャッシュ時フル出品（**リサーチ完了後に再提案**／別承認）または定時12-⑩（**在庫／競合価格側・保留**）。①の次L1はリサーチ agent。

**P1計画 dry 合格（Catalog GETなし）**: t18 ok。貼付10ブロックはすべて `skip_no_empty`（空ASINの埋め対象0）。12-⑮はクォータ不要。実Catalogは空欄が出てから。

**P2 12-⑱合格**: ログ `pasteP2-c5fe6e20` `write=true` DONE（21:12 JST）。直前⑰ `pasteP2-dc58ccaa` と同数。シートは全ASINに`[機械]`、上段候補。

---

## 0re. 領域1 リサーチ・見積もり（2026-08-16 完成形フロー）

**全体フォーカス（§0 出品B）とは別スレッド。** ①の実装担当は **リサーチ見積もり agent**。出品チャットは①のコードを書かない。  
**完成形・状況報告**: [org/B_PURCHASE_RESEARCH_FLOW.md](org/B_PURCHASE_RESEARCH_FLOW.md)（A–J / A′。聞かれたら §3。契約は §5。3者タイミングは §6）  
**構図**: [DOMAIN1 §1](DOMAIN1_RESEARCH_PURCHASING.md)　**予定の正（ツリー）**: [org/B_PURCHASE_RESEARCH_TASK_STRUCTURE.md](org/B_PURCHASE_RESEARCH_TASK_STRUCTURE.md)　**数字**: [org/B_PURCHASE_RESEARCH_NOW.md](org/B_PURCHASE_RESEARCH_NOW.md)

```text
領域1
├─ A 問屋                         外
├─ B 商品リサーチ
│  ├─ ① 仕入れ検討 ★リサーチ見積もり agent
│  │  ├─ 進め方
│  │  │    Keepaの正=競合DB Keepaフル  ロック（90日。出品と共用）
│  │  │    ①候補=門の作業面（転記）   ロック（倉庫ではない）
│  │  ├─ 実行基盤
│  │  │    調査複製スプシ              済（①候補 181行。競合タブを増やさない）
│  │  │    専用 GAS purchase_research/ 済（出品 clasp 禁止。Catalog は未搭載）
│  │  ├─ 門検収 C1–C8                 済
│  │  ├─ 倉庫本線
│  │  │    W1 Keepaフルへ①書き       スライス5合格
│  │  │    W2 ①候補ASINバックフィル  済（181/181。skip5＋GET176）
│  │  │    W3 門はフル優先→①候補転記 済（メニュー20:20）
│  │  │    W4 貼付90日後読み          出品PJ（dry: paste60 skip39 need21）
│  │  ├─ 前段①–⑧
│  │  │    ① セラー貯め→ストア正     済（タブ「セラー」モリタ1行。フルから貯めない）
│  │  │    ② ピック                   済（食品root・モリタ。S4構成比）
│  │  │    ③ 巡回                     済（第1版1セラー。query total=30）
│  │  │    ④ /query GET               済（モリタ30）
│  │  │    ⑤⑥⑦ 分類・食品・メーカー  済・仮想（モリタ30）
│  │  │    ⑧ 週次メーカー決定         済（今週=永谷園。Catalog GETなし）
│  │  ├─ 品番リスト転記
│  │  │    O1–O3 出品者数 COUNT_NEW  済（QS一致。空61＝今オファーなし）
│  │  │    F0 直販・新品現在・現行順位 済（GETなし。いる34/いない186）
│  │  │    F1 Keepa列 flatten       済（cat217/梱包157/FBA164。BB価格はJSON空）
│  │  │    F2／T1–T2                 済
│  │  │    課題 K1–K8               止／流用／不要（NOW表）
│  │  └─ 見積書ルート                 後
│  └─ ② 出品用                       外（貼付 Catalog 階段と混ぜない）
└─ C 見積                             外
```

**他PJ**: [TASK_STRUCTURE](org/B_PURCHASE_RESEARCH_TASK_STRUCTURE.md)。Catalog 生は `モールヒット` 禁止。**Keepaフルへ①が書く（W1）**。90日内は貼付が読む（W4）。出品 clasp 非差し替え。

### 次にやること
1. 完成形は [FLOW](org/B_PURCHASE_RESEARCH_FLOW.md)。心拍 A–D は継続。
2. **見積の今（2026-10-03）**: 本線は **メールから単価→利益が正しいか**（FLOW §5.3.2・チャット通し可）。G 常駐コードは3者のあと。**FBA最安／FBM最安は見積付与後だけ追加取得**（洗い・門では取らない）。利益列の nn 左レイアウト／自己発送は §5.3.2bis **仮**（本線より後）。
3. W4 本線は出品（`KEEPA_FULL_READ_ENABLED`）。楽天CSVは触らない

原本（不可侵）: https://docs.google.com/spreadsheets/d/1Uk-DEfoSTVqb5N8XsfSqciaZcIjzbbPpGvg4hQYgDl4/edit?gid=1060574934#gid=1060574934  
自動化用複製: https://docs.google.com/spreadsheets/d/1tf7gvkD88yyNz7JWXfNysBcIqSZDlI9dC-l6gOPyLjE/edit?gid=685859874#gid=685859874  
マニュアル: https://docs.google.com/presentation/d/1QbC8v_kwLS3W9sBMejecWAvukgL7_cY3KfbuLtorJ4I/edit?slide=id.g370bbd44c29_1_348#slide=id.g370bbd44c29_1_348

---

## 0rh. 行高さ（2026-08-16・**廃止**）

WRAP＋セル内改行で約60pxに伸びる。折り返しは運用上必要なため、CLIP＋固定高さは捨てた。B Step8・`setCheckedParentRowHeights_` 削除。人間の手調整は次Bで上書きされない（Stepが無い）。

---

## 0. セッション引き継ぎ（2026-08-14）

**場所**: **B完走フォロー** — [org/B_RUN_20260814_FOLLOWUP.md](org/B_RUN_20260814_FOLLOWUP.md)（コード改修済。⑥Nはコード非改修）。ハード死対策は継続。  
**正本**: [org/B_HARD_DEATH_SCOPE_REQUIREMENTS.md](org/B_HARD_DEATH_SCOPE_REQUIREMENTS.md)／[HUMAN_RUN](org/B_HARD_DEATH_SCOPE_HUMAN_RUN.md)

### 次にやること（人間）
1. `clasp push`
2. **メニューA（Keepa取得）→ B**（⑥Nの再テスト。同じなら原因再調査）
3. B完走後: 親行が子の約3倍で DF が行を伸ばしていないか、CV選定不可行に式が残るか、LJ改行、楽天Yahooヒット、FBA 6.55（N列＝競合ASIN）
4. Script Properties: **`B_WATCHDOG_ENABLED=false`** は維持

**ハード死対策（継続）**  
**場所**: B統合ハード死対策 — `B済`キュー自動パック（最大3・シリーズ温存）・Step1後切断・当B集合・番犬非設置。  

### Property
| Key | 推奨 |
|-----|------|
| `B_WATCHDOG_ENABLED` | **`false`（本対策中必須）** |
| `B_SCOPE_INSERTED_PRODUCTS_ONLY` | 未設定=ON |
| `B_STEP1_FORCE_SLICE_AFTER` | 未設定=ON |
| `B_STEP1_TOP_INSERT_ENABLED` | 上挿入するなら **`true`**／戻すなら `false` または削除 |

---

## 0t. セッション引き継ぎ（2026-08-14・上挿入）※履歴

**場所**: **B Step1 上挿入** — スプシ AK/IB 値固定＋テンプレ1–2＋GAS（Property 未設定=旧末尾）。  
**復元**: `_local_backup/pre_B_STEP1_TOP_INSERT_20260814_001210/`

### 済（当時）
- [多数決](org/B_STEP1_TOP_INSERT_THREE_REVIEW_MAJORITY.md)／[要件](org/B_STEP1_TOP_INSERT_REQUIREMENTS.md)／[HUMAN_RUN](org/B_STEP1_TOP_INSERT_HUMAN_RUN.md)

### Property（当時）
| Key | 推奨 |
|-----|------|
| `B_STEP1_TOP_INSERT_ENABLED` | 上挿入するなら **`true`**／戻すなら `false` または削除 |
| `B_PARENT_ROW_HEIGHT_ENABLED` | 未設定=ON／緊急 `false` |
| `B_DF_STRATEGY_FORMULA_COPY_ENABLED` | 未設定=ON |
| `B_WATCHDOG_ENABLED` | 本対策中は **false**（下記0節） |

---

## 0f. セッション引き継ぎ（2026-08-13・⑧f）※履歴

**場所**: **⑧f 列ロック確定実装**（基本ロジック＋例外列は要件 §11〜§13）。  

### 済
- **⑧f**: [要件](org/B_MASTER_CELL_COLOR_RULES_REQUIREMENTS.md) §11〜§13／[凡例](org/B_MASTER_CELL_COLOR_RULES_HUMAN_RUN.md)
  - `masterColorApplyLockedRow_`／`MASTER_COLOR_*_LETTERS_`／Step1適用
  - Z **15-㉑** 選択行／**15-㉒** 11〜775
  - スプシ: レ点行＋**11〜775** をロック色で塗色（API）
  - 親赤白 `#f4cccc`（白字なし）／エラー濃赤は別
- **⑧a＋⑧c** ほか（オレンジ未導入／条件付き書式非変更）

### Property
| Key | 推奨 |
|-----|------|
| `B_PARENT_ROW_HEIGHT_ENABLED` | 未設定=ON／緊急 `false` |
| `B_DF_STRATEGY_FORMULA_COPY_ENABLED` | 未設定=ON |
| `B_WATCHDOG_ENABLED` | 本対策中は **false**（§0） |

---

## 0s. セッション引き継ぎ（2026-08-12・シリーズ推定）※履歴

**場所**: **マスタシリーズ 不明禁止＋商品名推定**（7.6合格後）。  
**コミット**: 未（指示時）。  

### 次にやること（当時）
1. 人間: `clasp push` → シート再読込
2. 「不明」の親: **Step6 同期** → `マスタシリーズ` が推定名 or 空
3. 緊急停止: `SERIES_INFER_FROM_NAME_ENABLED=false`

### 正本
- [org/SERIES_INFER_FROM_NAME_REQUIREMENTS.md](org/SERIES_INFER_FROM_NAME_REQUIREMENTS.md)

---

## 0y. セッション引き継ぎ（2026-08-12・7.5／7.6）※履歴

**場所**: **メニュー8 Yahoo 7.5／7.6**。  
**コミット**: 未（指示時）。  

### 次にやること（当時）
1. clasp push → Z 7.6／B Step7.6 試験 → **7.6合格**

---

## 0z. セッション引き継ぎ（2026-08-12・E本線化）※履歴

**場所**: **E本線化**（楽天ジャンル／Yahoo cat＝メニュー8都度API）＋⑥ N列ASIN。  
**コミット**: 未（指示時）。  

### 次にやること（当時）
1. 人間: `clasp push` → シート再読込
2. 旧 false キー削除
3. レ点 3〜10親で B または手動7.5

### 正本
- [org/D_MENU_E_GENRE_YAHOO_REQUIREMENTS_CONFIRM.md](org/D_MENU_E_GENRE_YAHOO_REQUIREMENTS_CONFIRM.md)

---

## 0a. セッション引き継ぎ（2026-08-12・⑥N列）※履歴

**場所**: **⑥ N列ASIN自動**（◎×ブランド＝メーカー → N列＋黄セル）。  
**コミット**: 未（指示時）。  

### 次にやること（当時）
1. 人間: `clasp push` → シート再読込
2. 試験: Z→**15-⑮**
3. E: 要件確認 → **本線化済**（§0へ）

---

## 0b. セッション引き継ぎ（2026-08-12・メニュー8）※履歴

**場所**: メニュー8 **v1.13**（ログで判明: 列J=商品名ベースまで削っていた＋FO二重計上）。  
**コミット**: 未（指示時）。  

### 次にやること（当時）
1. 人間: 版履歴で **7.5前**に戻す（商品名ベース回復）
2. 人間: `clasp push`（v1.13）
3. 7.5を **1回** → GD≒70–75・商品名ベース残存を確認 → Property false

### 正本
- [org/D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md](org/D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md)（v1.13）
- [org/D_MENU_AMAZON_AI_ADOPT_HUMAN_RUN.md](org/D_MENU_AMAZON_AI_ADOPT_HUMAN_RUN.md)

### Property
| Key | 推奨 |
|-----|------|
| `AMAZON_AI_AUTO_ADOPT_ENABLED` | **false**（手動7.5のときだけ一時 true） |
| `B_INTEGRATED_MENU8_ENABLED` | 未設定／true |

---

## 0b. セッション引き継ぎ（2026-08-10）※履歴

**場所**: Amazon出品は **本番運用可**（缶飯・相乗り等 SC出品中確認済。A3サブ画像も合格）。開発フォーカスはサブ画像楽天出口の品質改善など。  
**コミット**: 未（指示時）。  

### 次にやること
1. 人間: `clasp push`（常時ON既定のコード反映）
2. Script Properties: 下表の常時ONキーが **false のまま残っていれば削除 or true**（1回だけ）
3. （任意）サブ画像楽天の写真ルール改善は `photo_realism_rules.py`

### Property チェックリスト（人間・本番常時ONセット 2026-08-10）

**考え方**: 毎日の D 出品で Script Properties を切り替えない。  
**未設定＝ON**（コード既定）。既に `false` が入っているキーは **削除**するか **true に1回**書き換える。  
緊急停止だけ明示 `false`。

**本番で常時ON（未設定で可／明示 false のみ緊急停止）**
| Property | 推奨 |
|----------|------|
| `APPROVAL_AMAZON_SPAPI_PUT_ENABLED` | **true／未設定** |
| `APPROVAL_AMAZON_SPAPI_PUT_ALLOW_PROD` | **true／未設定**（開始前確認ダイアログは残る） |
| `APPROVAL_AMAZON_LV4_ENABLED` | **true／未設定** |
| `APPROVAL_AMAZON_LV4_SC_SUMMARY_ENABLED` | **true／未設定** |
| `APPROVAL_AMAZON_LV4_SC_SUMMARY_FOLDER_ID` | 監視フォルダID（必須） |
| `APPROVAL_AMAZON_LV4_TRACK` | **B**（新規GENERATED） |
| `APPROVAL_AMAZON_SPAPI_PUT_FBA_COMPLIANCE_ATTRS` | **true／未設定** |
| ~~`APPROVAL_AMAZON_SPAPI_PUT_ALLOW_MASTER_QTY`~~ | **廃止**（2026-10-24。マスタ在庫は Property 不要。キーが残っていても読まない） |
| `AMAZON_SHELF_REGISTRY_FILE_ID` | Drive の `shelf_registry.json` ファイルID |
| `AMAZON_IMAGE_CANDIDATE_FOLDER_ID` | Drive 07（C本線） |

**常時OFFのまま（毎回ONにしない）**
| Property | 推奨 | 理由 |
|----------|------|------|
| `APPROVAL_AMAZON_P4B_PT_WRITE_ENABLED` | **false／未設定** | PT空の新規だけ一時 true |
| `AMAZON_IMAGE_U2_ENABLED` | **false／未設定** | Cコースが実行時だけ一時ON |
| `AMAZON_U4_URL_EMBED_ENABLED` | **false／未設定** | 同上（C-2／自動U4経路） |

**任意**
| Property | メモ |
|----------|------|
| `APPROVAL_AMAZON_LV4_SC_SUMMARY_INTERVAL_MIN` | 既定15。疎にするなら60 |
| `YAHOO_OAUTH_WARN_DAYS` | 既定7 |
| `RAKUTEN_IMAGE_MAIN_AUTOBIND_ENABLED` | 未設定＝**true**。旧運用に戻すときだけ `false` |
| `AMAZON_IMAGE_MAIN_AUTOBIND_ENABLED` | 未設定＝**true** |
| `YAHOO_OAUTH_REDIRECT_URI` | **店の正=`http://localhost/`**（Yahooアプリコールバックと一致。B14トークンと混ぜない）。[再認証HUMAN_RUN §0](org/YAHOO_OAUTH_REAUTH_HUMAN_RUN.md) |

### SEASONINGで確定した必須・落とし穴（再発防止）
| 項目 | 対応 |
|------|------|
| 商品の形式 | 既定 `粉末`（列134） |
| 商品の重量の単位 | 重量ありなら単位必須。既定 `グラム`（列166） |
| 検索キーワード | **1枠のみ**（空白結合）。5枠分割は SC **99016** |
| 画像URL | C1は `Amazon MAIN/PT URL` のみ。サブはU4で楽天→R2 |
| マッチングsheet | 楽天サブとAmazon PTの**二重ドラッグ不要**（MAIN白抜きのみ手動） |
| 冪等 | 同じ親の再GENERATEDは D内レ点「失敗後の再GENERATED」または 21-④ `UPLOAD_FAILED` 後 |

### M2 正本リンク
- [org/D_MENU_M2_HUMAN_RUN.md](org/D_MENU_M2_HUMAN_RUN.md)  
- [org/LV4_M2_IMPLEMENTATION_APPROVAL.md](org/LV4_M2_IMPLEMENTATION_APPROVAL.md)  
- [org/LV4_M2_TRACK_A_GAP_ANALYSIS.md](org/LV4_M2_TRACK_A_GAP_ANALYSIS.md)  
- `tools/m2_offer_packaged/`  
- [org/D_MENU_SPAPI_SMOKE_HUMAN_RUN.md](org/D_MENU_SPAPI_SMOKE_HUMAN_RUN.md)／`tools/spapi_smoke/`  
- [org/D_MENU_SPAPI_LISTINGS_WRITE_HUMAN_RUN.md](org/D_MENU_SPAPI_LISTINGS_WRITE_HUMAN_RUN.md)／`tools/spapi_listings_write/`  
- [org/LV4_SPAPI_LISTINGS_WRITE_BATCH_APPROVAL.md](org/LV4_SPAPI_LISTINGS_WRITE_BATCH_APPROVAL.md)（v1.1）  
- [org/D_MENU_SPAPI_SHEET_BRIDGE_HUMAN_RUN.md](org/D_MENU_SPAPI_SHEET_BRIDGE_HUMAN_RUN.md)／[org/LV4_SPAPI_SHEET_BRIDGE_APPROVAL.md](org/LV4_SPAPI_SHEET_BRIDGE_APPROVAL.md)  
- [org/LV4_SPAPI_CHECKBOX_EXPORT_APPROVAL.md](org/LV4_SPAPI_CHECKBOX_EXPORT_APPROVAL.md)（v1.2c）  
- [org/LV4_SPAPI_DRIVE_FETCH_APPROVAL.md](org/LV4_SPAPI_DRIVE_FETCH_APPROVAL.md)／[org/LV4_SPAPI_APPROVED_EXPORT_APPROVAL.md](org/LV4_SPAPI_APPROVED_EXPORT_APPROVAL.md)  
- [org/LV4_SPAPI_GAS_PUT_APPROVAL.md](org/LV4_SPAPI_GAS_PUT_APPROVAL.md)／[org/D_MENU_SPAPI_GAS_PUT_HUMAN_RUN.md](org/D_MENU_SPAPI_GAS_PUT_HUMAN_RUN.md)（v1.4 第1段・API実機合格／SC最終更新は反映待ち）  
- [org/LV4_SPAPI_GAS_PUT_STAGE2_APPROVAL.md](org/LV4_SPAPI_GAS_PUT_STAGE2_APPROVAL.md)（v1.4 **第2段・実機合格（API）**）  
- [org/LV4_SPAPI_D_ENTRY_APPROVAL.md](org/LV4_SPAPI_D_ENTRY_APPROVAL.md)／[org/D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md](org/D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md)（**Dレ点本線・相乗り自己発 dry_run／prod 実機合格**）
- [org/LV4_AMAZON_CHECKBOX_MAINLINE_SELLER_SKU_APPROVAL.md](org/LV4_AMAZON_CHECKBOX_MAINLINE_SELLER_SKU_APPROVAL.md)（**レ点本線・拡大実機合格**）
- [org/LV4_A1_FBA_COMPLIANCE_ATTRS_APPROVAL.md](org/LV4_A1_FBA_COMPLIANCE_ATTRS_APPROVAL.md)（**FBA属性・実機合格**）
- [org/LV4_DUAL_OFFER_MFN_FBA_APPROVAL.md](org/LV4_DUAL_OFFER_MFN_FBA_APPROVAL.md)（**デュアル Phase1・検収OK**）
- [org/LV4_DUAL_OFFER_PHASE2_APPROVAL.md](org/LV4_DUAL_OFFER_PHASE2_APPROVAL.md)／[org/D_MENU_DUAL_OFFER_PHASE2_HUMAN_RUN.md](org/D_MENU_DUAL_OFFER_PHASE2_HUMAN_RUN.md)（**Phase2・検収OK**）
- [org/LV4_LANE_A1_FBA_OFFER_APPROVAL.md](org/LV4_LANE_A1_FBA_OFFER_APPROVAL.md)／[org/D_MENU_LANE_A1_FBA_HUMAN_RUN.md](org/D_MENU_LANE_A1_FBA_HUMAN_RUN.md)（**レーンA1・検収OK**）
- [org/LV4_P2_DC_123_INVESTIGATION_APPROVAL.md](org/LV4_P2_DC_123_INVESTIGATION_APPROVAL.md)／[org/D_MENU_P2_DC_HUMAN_RUN.md](org/D_MENU_P2_DC_HUMAN_RUN.md)（**P2調査完了＋§7.4方針ロック**・③は低優先）
- [org/LV4_P1_FILE_MIN_APPROVAL.md](org/LV4_P1_FILE_MIN_APPROVAL.md)／[org/D_MENU_P1_HUMAN_RUN.md](org/D_MENU_P1_HUMAN_RUN.md)（**P1 File最少・検収OK**）
- [org/LV4_AMAZON_CATEGORY_PT_POC_APPROVAL.md](org/LV4_AMAZON_CATEGORY_PT_POC_APPROVAL.md)／[org/LV4_AMAZON_CATEGORY_PT_POC_HUMAN_RUN.md](org/LV4_AMAZON_CATEGORY_PT_POC_HUMAN_RUN.md)（**P4a読取PoC・実機合格**）
- [org/LV4_AMAZON_CATEGORY_PT_P4B_APPROVAL.md](org/LV4_AMAZON_CATEGORY_PT_P4B_APPROVAL.md)／[org/D_MENU_P4B_CATEGORY_PT_HUMAN_RUN.md](org/D_MENU_P4B_CATEGORY_PT_HUMAN_RUN.md)（**P4b-a／b合格。2026-09-17 貼付◎本線**）
- [org/LV4_D_NEW_PT_SHELF_GATE_APPROVAL.md](org/LV4_D_NEW_PT_SHELF_GATE_APPROVAL.md)／[org/D_MENU_D_NEW_PT_SHELF_GATE_HUMAN_RUN.md](org/D_MENU_D_NEW_PT_SHELF_GATE_HUMAN_RUN.md)（**D新規ゲート＋Cursor手渡し・実装済**）
- [org/LV4_SHELF_BROWSE_CATALOG_APPROVAL.md](org/LV4_SHELF_BROWSE_CATALOG_APPROVAL.md)／[org/D_MENU_SHELF_BROWSE_CATALOG_HUMAN_RUN.md](org/D_MENU_SHELF_BROWSE_CATALOG_HUMAN_RUN.md)（**SHELF Browse 網羅・Node ルーティング**）
- [org/LV4_MAP_SHEET_JSON_SYNC_APPROVAL.md](org/LV4_MAP_SHEET_JSON_SYNC_APPROVAL.md)／[org/D_MENU_MAP_SHEET_JSON_SYNC_HUMAN_RUN.md](org/D_MENU_MAP_SHEET_JSON_SYNC_HUMAN_RUN.md)／[org/MAP_SC_ERROR_LEDGER.md](org/MAP_SC_ERROR_LEDGER.md)（**正本=sheet／派生=MD／実行=JSON**）
- [org/AMAZON_DEV_ROADMAP_P0_P4.md](org/AMAZON_DEV_ROADMAP_P0_P4.md)
- [org/LV4_C_COURSE_CONSOLIDATION_APPROVAL.md](org/LV4_C_COURSE_CONSOLIDATION_APPROVAL.md)／[org/D_MENU_C_IMAGE_COURSE_HUMAN_RUN.md](org/D_MENU_C_IMAGE_COURSE_HUMAN_RUN.md)
- [org/LV4_D_P0_E_ABSORB_INVENTORY_APPROVAL.md](org/LV4_D_P0_E_ABSORB_INVENTORY_APPROVAL.md)

### 外出先チェックリスト
[org/D_MENU_C1_ANMIN_REMOTE_CHECKLIST.md](org/D_MENU_C1_ANMIN_REMOTE_CHECKLIST.md) — 安眠完了済み（履歴）  

### 正本・手順リンク
- [org/LV4_M2_IMPLEMENTATION_APPROVAL.md](org/LV4_M2_IMPLEMENTATION_APPROVAL.md)（**M2 v1・実機合格**）  
- [org/LV4_M2_TRACK_A_GAP_ANALYSIS.md](org/LV4_M2_TRACK_A_GAP_ANALYSIS.md)／[org/D_MENU_M2_HUMAN_RUN.md](org/D_MENU_M2_HUMAN_RUN.md)（**公式Loader正**）  
- [org/D_MENU_C1_ANMIN_REMOTE_CHECKLIST.md](org/D_MENU_C1_ANMIN_REMOTE_CHECKLIST.md)（安眠完了・履歴）  
- [org/D_MENU_E_AMAZON_COURSE_HUMAN_RUN.md](org/D_MENU_E_AMAZON_COURSE_HUMAN_RUN.md)  
- [org/D_MENU_AMAZON_AI_ADOPT_HUMAN_RUN.md](org/D_MENU_AMAZON_AI_ADOPT_HUMAN_RUN.md)／[org/D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md](org/D_MENU_AMAZON_AI_ADOPT_REQUIREMENTS.md)  
- [YAHOO_CATEGORY_BRAND_STAGE.md](YAHOO_CATEGORY_BRAND_STAGE.md)／[org/LV4_YAHOO_CATEGORY_BRAND_IMPLEMENTATION_APPROVAL.md](org/LV4_YAHOO_CATEGORY_BRAND_IMPLEMENTATION_APPROVAL.md)／[org/D_MENU_YAHOO_CATEGORY_BRAND_HUMAN_RUN.md](org/D_MENU_YAHOO_CATEGORY_BRAND_HUMAN_RUN.md)  
- [org/D_MENU_C1_HUMAN_RUN.md](org/D_MENU_C1_HUMAN_RUN.md)  
- [org/D_MENU_C1_MASTER_HPC_COLUMN_MAP.md](org/D_MENU_C1_MASTER_HPC_COLUMN_MAP.md)  
- [org/D_MENU_C1_MASTER_FOOD_SEASONING_COLUMN_MAP.md](org/D_MENU_C1_MASTER_FOOD_SEASONING_COLUMN_MAP.md)  
- [org/D_MENU_C1_PACKAGED_XLSM_REQUIREMENTS.md](org/D_MENU_C1_PACKAGED_XLSM_REQUIREMENTS.md)  
- [org/D_MENU_U4_HUMAN_RUN.md](org/D_MENU_U4_HUMAN_RUN.md)  
- [org/YAHOO_OAUTH_REAUTH_HUMAN_RUN.md](org/YAHOO_OAUTH_REAUTH_HUMAN_RUN.md)（Yahoo再認証半自動・C14／B17メール）  

---

## 1. プロジェクト全体（30秒）

| 項目 | 内容 |
|------|------|
| **目的** | 楽天・Yahoo・Amazon 向けに、リサーチ〜商品情報・価格・セット構成〜出品に必要なデータを **スプレッドシート＋GAS** で整備・自動化する。一人社長＋AI部門モデルは [org/AI_ORG_CHARTER.md](org/AI_ORG_CHARTER.md)。 |
| **クリティカルパス** | [FLOW_AND_PRIORITY.md](FLOW_AND_PRIORITY.md) — リサーチ・見積もり → 出品情報 → 各モール出品。 |
| **実装の中心** | コード.js、主シートは ▼商品マスタ(人間作業用)、AI情報取得data、00_設定マスタ。 |
| **6領域・成果物** | [PROJECT_OVERVIEW.md](PROJECT_OVERVIEW.md)。 |

---

## 2. 現在のフェーズ（いま優先している開発）

- **フォーカス領域**: Amazon出品は **本番常時ONセット適用済**。販促 **2層減衰実装**（カレンダー`減衰中%`＋B中1%。restore=減衰中%・[要件§10.13/10.14](org/D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)）。並行でサブ画像楽天品質。  
- **Lv4（Amazonバルク）**:  
  - HPC／FOOD／T2／U3／**U2**／**U4（楽天サブ流用）**: 実装済・運用合格。  
  - C1 SEASONING: 七味 **出品中**。  
  - C1 GROCERY: 缶飯 **出品中**（100521待ち解消・人間確認済）。  
  - C1 SOUP: 食べるおだし PACKAGED。**必須だけ常時埋め**。寸法・重量は必須から除外。新PTは 09＋E③のあと B-T1 auto-onboard。  
  - 本線UX: [org/D_MENU_U2_HUMAN_RUN.md](org/D_MENU_U2_HUMAN_RUN.md)／[org/D_MENU_U4_HUMAN_RUN.md](org/D_MENU_U4_HUMAN_RUN.md)／[org/D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md](org/D_MENU_SPAPI_D_ENTRY_HUMAN_RUN.md)。  
- **楽天Nav Stage1**: 実機PASS。Propertyはfalse。  
- **Lv3（Yahoo）**: 2026-07-20 人間検収完了。  
- **Lv2（楽天）**: 2026-07-20 人間検収完了。  
- **Lv1（承認キュー）**: 2026-07-17 人間検収完了。  

- **このフェーズの完了条件（目安）**  
  - ~~缶飯新規カタログが公開~~ **済（人間SC確認）**。  
  - ~~サブ画像 Amazon（A3）~~ **合格（人間確認）**。  

- **並行・継続（後回し可）**: T3 ZIP自動／Dc API強化／レーンC・P3／ε。  

- **スコープ外（次モール着手前）**  
  - 承認②、販売中SKU無人上書き、clasp push 自動化。  
  - generateRakutenCSV 本体・Yahoo.js 出品API本体の改変。

---

## 3. 直近で確定した仕様・前提（docs 反映済み）

- **AI組織**: **ハイブリッド**（承認①掲載は軽い／本命は承認②在庫反映）→ 日中無人出品（**ユニーク画像50枚＋実働25分**で自動分割・原則12:00前）→ 夜確認。残リストは明示取消まで残し、マスタ無しは画面非表示（履歴残す）。在庫0原則。販売中(在庫>0)は原則スキップ。CS／問屋は送信禁止・下書き可。詳細は org 憲章・マトリクス・多数決メモ §4.2・§4.3。  
- **キーワード案・商品名案の本線は OpenAI**。Gemini 版の商品名案は **コードに存在するが本流メニュー未配線**。  
- **バリエーション単位**は **Gemini → OpenAI 補完**（◎ASIN 経路も同様）。**内容量パース**は Gemini 中心（ChatGPT 自動フォールバックなし）。  
- **11-③（runProductNameProposalsForRows）** は、**OpenAI 失敗時もマスタ商品名があればバリエーションに届く**（既定）。旧挙動は PRODUCT_NAME_PROPOSALS_CONTINUE_VARIATION_ON_OPENAI_FAIL=false。**B Step7** は **variation まで進み得る**（商品名が空でも単位・内容量のみ更新されうる）。  
- 詳細は [AI_ROUTING_GEMINI_OPENAI.md](AI_ROUTING_GEMINI_OPENAI.md)、[商品マスタ_人間作業エリアとマスタエリア_要件定義.md](商品マスタ_人間作業エリアとマスタエリア_要件定義.md)。

---

## 4. 次にやること（優先順）

1. （済）8/14 taper prod: Cinderellas b/s 22→18。**店頭ともに18%確認済**。  
2. 日次 `taper_send.py --poll --mail`（dry_run）。Windows タスク **`OctasAmazonTaperDryRun`** 09:00。安定後に `--prod`。  
3. **開発完了後**: タイムセール運用マニュアル化。  
4. 8/27 Smile apply リマインド。9/4 restore＝減衰中%（仮想: 14）。  
5. T3／ε／U7（B SP-API品番・Dc④⑤）は各ゲート承認後。  
6. remote `git push` は指示時。  
7. スコープ外（後で再提案）: B開始自動 apply／B終了自動 restore／数量メール `--send`／P1c-2／P1c-C／レーンA。  

---

## 5. 深掘りリンク（領域別）

| テーマ | ドキュメント |
|--------|----------------|
| **AI組織・承認** | [org/AI_ORG_CHARTER.md](org/AI_ORG_CHARTER.md)、[org/AI_APPROVAL_MATRIX.md](org/AI_APPROVAL_MATRIX.md)、[org/THREE_REVIEW_RUNBOOK.md](org/THREE_REVIEW_RUNBOOK.md)、[org/PHASE0_THREE_REVIEW_MAJORITY.md](org/PHASE0_THREE_REVIEW_MAJORITY.md)、[org/LEVELLED_IMPLEMENTATION_PLAN.md](org/LEVELLED_IMPLEMENTATION_PLAN.md)、[org/LV1_APPROVAL_QUEUE_REQUIREMENTS.md](org/LV1_APPROVAL_QUEUE_REQUIREMENTS.md)、[org/LV2_RAKUTEN_ORCHESTRATION_REQUIREMENTS.md](org/LV2_RAKUTEN_ORCHESTRATION_REQUIREMENTS.md)、[org/LV3_YAHOO_ORCHESTRATION_REQUIREMENTS.md](org/LV3_YAHOO_ORCHESTRATION_REQUIREMENTS.md)、[org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md](org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md)、[org/LV4_THREE_REVIEW_MAJORITY.md](org/LV4_THREE_REVIEW_MAJORITY.md) |
| 楽天ジャンル Nav（並行） | [RAKUTEN_NAV_GENRE_DIAG.md](RAKUTEN_NAV_GENRE_DIAG.md)（Stage0）、[RAKUTEN_NAV_GENRE_STAGE1.md](RAKUTEN_NAV_GENRE_STAGE1.md)（Stage1）、[RAKUTEN_NAV_GENRE_STAGE3.md](RAKUTEN_NAV_GENRE_STAGE3.md)（**Stage3都度API・実装済・要push**） |
| Gemini / OpenAI・11-③ vs B Step7・429 | [AI_ROUTING_GEMINI_OPENAI.md](AI_ROUTING_GEMINI_OPENAI.md) |
| 商品マスタ人間作業エリア | [商品マスタ_人間作業エリアとマスタエリア_要件定義.md](商品マスタ_人間作業エリアとマスタエリア_要件定義.md) |
| 価格・送料・再③・B 統合 | [PRICING_V1_REQUIREMENTS.md](PRICING_V1_REQUIREMENTS.md)、AGENT_HANDOVER **§8** |
| リサーチ・セット構成 | [RESEARCH_AND_ESTIMATE.md](RESEARCH_AND_ESTIMATE.md) |
| エージェント共通・必読一覧 | [AGENT_HANDOVER.md](AGENT_HANDOVER.md) **§2** |

---

## 6. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-11-03 | **P3-W第2**: WRAP確定後手持ち候補UI。`B4_WRAP_ASSET_LANE=0`切戻し。clasp不要。[ALIGN §4.5.11](org/B45_MAIN_SUB_ALIGN_REQUIREMENTS.md) |
| 2026-11-03 | **P3-W hasWrap鍵**: 親SKU→商品名ベース→JAN。export全件 bool。[ALIGN §4.5.11](org/B45_MAIN_SUB_ALIGN_REQUIREMENTS.md) |
| 2026-11-03 | **P3-W第1**: WRAP→個包装訴求 themeId=21 +1（最大11）。切戻し `B4_WRAP_APPEAL_THEME=false`。要 clasp。[ALIGN §4.5.11](org/B45_MAIN_SUB_ALIGN_REQUIREMENTS.md) |
| 2026-11-02 | **MAIN→サブ寄せ P3第1配線**（G7/G6/G1→dry_sub）。切戻し `B4_SUB_PICTURE_GATE=0`。[ALIGN §4.5.10](org/B45_MAIN_SUB_ALIGN_REQUIREMENTS.md) |
| 2026-11-02 | **MAIN→サブ寄せ P2第2 人間検収OK**（訴求軸・つなぎ）。次＝P3。[ALIGN §4.5.9](org/B45_MAIN_SUB_ALIGN_REQUIREMENTS.md) |
| 2026-11-02 | **MAIN→サブ寄せ P1 人間検収OK**（再利用・レ点外し・U1高速化）。次＝P2。[ALIGN §4.5.7](org/B45_MAIN_SUB_ALIGN_REQUIREMENTS.md) |
| 2026-11-02 | **MAIN→サブ寄せ P0**: [ALIGN §4.5](org/B45_MAIN_SUB_ALIGN_REQUIREMENTS.md)。 |
| 2026-09-18 | **バルクごと必須埋め**: 純正「データ定義」の必須／条件付き必須を読み、値がある列だけ書く。GROCERY黒セルomitは缶飯専用。[T2](org/D_MENU_LANE_B_BULK_T2_HUMAN_RUN.md) |
| 2026-09-18 | **PACKAGED auto-onboard**: 新バルクは 09+E③のあと B-T1 が指紋＋列マップを自動登録して 03 へ。[T1 PROD](org/D_MENU_LANE_B_BULK_T1_PROD_HUMAN_RUN.md) |
| 2026-09-16 | **Yahoo説明 P6c**: pool＝マスタ `楽天サブ画像1〜10` の入っている枠。落選・空欄は売れそう順に入れない。[要件 §0.1](org/YAHOO_CAPTION_TEMPLATE_LOCK_REQUIREMENTS.md) |
| 2026-09-15 | **Yahoo説明 §0.1**: 完成形＝採用テーマからAI売れそう順の上位3。人手型禁止。実装未。[要件](org/YAHOO_CAPTION_TEMPLATE_LOCK_REQUIREMENTS.md) |
| 2026-09-04 | **文字色共有 実装**: Prop＋Drive JSON。要 clasp。[ALIGN §3](org/B45_MAIN_SUB_ALIGN_REQUIREMENTS.md)。 |
| 2026-09-05 | **B-⑥ Amazon MAIN**: メニュー本線化。CからAmazon量産を外し補助化。要 clasp。[B6](org/D_MENU_B6_AMAZON_MAIN_REQUIREMENTS.md)／[ALIGN §1.2](org/B45_MAIN_SUB_ALIGN_REQUIREMENTS.md) |
| 2026-09-04 | **B一括＋B-③ 案1**: メニュー増なし。B起動確認→`B_INTEGRATED_INCLUDE_B3_IMAGES`。全パック後に約1分で B-③（Amazon自動）。要 clasp。[ALIGN §1.1](org/B45_MAIN_SUB_ALIGN_REQUIREMENTS.md) |
| 2026-09-01 | **§7.14 個品見本共有**: MAIN／サブで同一 IDENTITY。構図は別。P-align-3 要件ロック・実装未。[ALIGN](org/B45_MAIN_SUB_ALIGN_REQUIREMENTS.md)。 |
| 2026-08-29 | **B-⑤ 画像本線=OpenAI**（仮想v12人間ロック）。次は P1 メニュー。[HUMAN_RUN](org/D_MENU_B5_RAKUTEN_MAIN_BASE_HUMAN_RUN.md)。 |
| 2026-08-29 | **B-⑤ 仮想v12**: ③b/c主婦向け IMAGE_STYLE_TYPO。T3新規30（既存・退避除外）。①は abc11 流用。[HUMAN_RUN](org/D_MENU_B5_RAKUTEN_MAIN_BASE_HUMAN_RUN.md)。 |
| 2026-08-28 | **B-⑤ 仮想v11**: ③b/cスマホ主役サイズ（金丸半分）。Gemini/OpenAI再比較。`…213826_abc11`。[HUMAN_RUN](org/D_MENU_B5_RAKUTEN_MAIN_BASE_HUMAN_RUN.md)。 |
| 2026-08-28 | **B-⑤ 仮想v10**: ③に写真実写＋キャッチ必須1本。Gemini/OpenAI再比較。`…200941_abc10`。[HUMAN_RUN](org/D_MENU_B5_RAKUTEN_MAIN_BASE_HUMAN_RUN.md)。 |
| 2026-08-28 | **B-⑤ 仮想v9**: Gemini/OpenAI同じ指示比較。右上空け→Pillow金丸。gpt-image-2。[HUMAN_RUN](org/D_MENU_B5_RAKUTEN_MAIN_BASE_HUMAN_RUN.md)。 |
| 2026-08-28 | **B-⑤ 仮想v8**: ③b/c=スプシ材料・IMAGE_GOLDをAI配置・文字20%超は再作成。[HUMAN_RUN](org/D_MENU_B5_RAKUTEN_MAIN_BASE_HUMAN_RUN.md)。 |
| 2026-08-28 | **B-⑤ 仮想v7**: ③b/cに虚偽禁止・金丸固定・白背景不可・文字20%・材料内売り。[HUMAN_RUN](org/D_MENU_B5_RAKUTEN_MAIN_BASE_HUMAN_RUN.md)。 |
| 2026-08-28 | **B-⑤ 仮想v6**: ③a添付構図固定。③b/c短い売れる指示。[HUMAN_RUN](org/D_MENU_B5_RAKUTEN_MAIN_BASE_HUMAN_RUN.md)。 |
| 2026-08-28 | **B-⑤ 仮想v5**: ①流用。②見出しルマ切替。③トンマナ＝人間ベスト。[HUMAN_RUN](org/D_MENU_B5_RAKUTEN_MAIN_BASE_HUMAN_RUN.md)。 |
| 2026-08-28 | **B-⑤ 仮想v4**: ①シズル+袋左下+Octas。②③空金丸。`4906283045119_20260828_004133_abc4`。[HUMAN_RUN](org/D_MENU_B5_RAKUTEN_MAIN_BASE_HUMAN_RUN.md)。 |
| 2026-08-27 | **B-⑤ 見出しAI**: 金丸は98後貼り。文字≤20%。色はCの textColors。[要件](org/D_MENU_B5_RAKUTEN_MAIN_BASE_REQUIREMENTS.md)。 |
| 2026-08-26 | **B-⑤ 案Aロック**: テーマ対最大4→人1枚。Pillowアピール。[要件](org/D_MENU_B5_RAKUTEN_MAIN_BASE_REQUIREMENTS.md)。実装は別承認。 |
| 2026-08-12 | **B番犬P0実装**: ハード死後自動再開・Step5行CP・プルダウンTIME_SLICE・サマリ・進捗メニュー。要 clasp push。[要件](org/B_WATCHDOG_RESUME_REQUIREMENTS.md)。 |
| 2026-08-12 | **販促2層減衰**: カレンダー減衰中%＋B中1%オーバーレイ。restore=減衰中%。`減衰開始日`。fetchは現在%のみ。要 clasp push／`--schema-only`。 |
| 2026-08-10 | **本番常時ONセット**: PUT／ALLOW_PROD／LV4／SCサマリの未設定既定=true。MASTER_QTY・P4B・U2・U4は常時OFF。A1〜B2・A3合格反映。要 clasp push。 |
| 2026-08-02 | **セットMAIN Phase A 実装**（`tools/set_main_image`・スモークOK）。実素材はHUMAN_RUN。[承認](org/LV4_SET_MAIN_IMAGE_PHASE_A_APPROVAL.md)。 |
| 2026-08-02 | **セットMAIN Phase A 方針ロック**（Pillow・07/C接続・一括・Amazon先）。実装は別承認。[承認](org/LV4_SET_MAIN_IMAGE_PHASE_A_APPROVAL.md)。 |
| 2026-08-02 | **七味出品中確定**（48h待ち削除）。D内レ点＋ダイアログ **clasp push済**。缶飯は100521待ち。 |
| 2026-08-02 | **D内レ点実装**（失敗後再GENERATED・案A）。[承認](org/LV4_D_REMAKE_MENU_APPROVAL.md)。 |
| 2026-08-02 | **D内レ点方針ロック**＋MAP sync HUMAN_RUN固め。[承認](org/LV4_D_REMAKE_MENU_APPROVAL.md)／[MAP HUMAN_RUN](org/D_MENU_MAP_SHEET_JSON_SYNC_HUMAN_RUN.md)。 |
| 2026-08-02 | **SHELF Browse 網羅**（抽出・MAP・P4b/D Node ルーティング・缶飯 GROCERY）。[承認](org/LV4_SHELF_BROWSE_CATALOG_APPROVAL.md)。 |
| 2026-08-02 | **D新規ゲート＋Cursor手渡し実装**（Drive `AMAZON_SHELF_REGISTRY_FILE_ID`・3者スキップ）。要 clasp push＋File ID設定。[承認](org/LV4_D_NEW_PT_SHELF_GATE_APPROVAL.md)。 |
| 2026-08-01 | **Amazon候補を1子SKU＝1枚**（C画像コース）。要 clasp push。[C HUMAN_RUN](org/D_MENU_C_IMAGE_COURSE_HUMAN_RUN.md)。 |
| 2026-08-01 | **相乗り ASIN空 soft skip**（相乗りのみ含む）。`コード.js`。要 clasp push。[承認§2.1](org/LV4_D_NEW_PT_SHELF_GATE_APPROVAL.md)。 |
| 2026-08-01 | **D新規ゲート＋Cursor手渡し方針ロック**（docs）。実装は別承認。[LV4_D_NEW_PT_SHELF_GATE_APPROVAL](org/LV4_D_NEW_PT_SHELF_GATE_APPROVAL.md)。 |
| 2026-07-29 | **A（D入口版）実装**: D ラジオに既存相乗り dry/prod。要 clasp push→実機。 |
| 2026-07-29 | **A（D入口版）承認**（社長「承認」）。実装可・範囲は承認 §2・§3。 |
| 2026-07-29 | **A（D入口版）承認起草**: D ラジオで Amazon 新規／既存相乗り（dry_run／prod）。コードなし。 |
| 2026-07-29 | **SP-API v1.4 第2段 実機合格（API）**: 21-⑫ VALID／21-⑬ ACCEPTED（batch `…b7a053`）。SC目視は反映待ち。 |
| 2026-07-29 | **SP-API v1.4 第2段 実装**: 21-⑫⑬／承認①済抽出。 |
| 2026-07-29 | **SP-API v1.4 第2段 承認**（社長「承認する」）。実装可・範囲は承認 §2・§3。 |
| 2026-07-29 | **SP-API v1.4 第2段 承認起草**（承認①済→GAS直PUT・コードなし）。三者レビュー不要。 |
| 2026-07-29 | **SP-API v1.4 実機合格（API）**＋ENDPOINT正規化。SC最終更新は反映待ち。 |
| 2026-07-29 | **SP-API v1.4 実装**: 21-⑩⑪／`AmazonSpapiPut.js`。要 clasp push→実機。 |
| 2026-07-29 | **SP-API v1.4 承認起草**（GAS直呼び・コードなし）。三者レビュー不要。 |
| 2026-07-29 | **SP-API v1.2b 実機合格**＋親レ点出さないを完了扱い。 |
| 2026-07-29 | **SP-API v1.2c／v1.3 実機合格**: 子レ点→Drive→fetch-drive prod（`…48s11`）。 |
| 2026-07-28 | **SP-API v1.2c**: 21-⑧＝子SKUレ点のみ（選択行廃止）。 |
| 2026-07-26 | **メニュー8緊急修正**: sync除外。要 clasp push。 |
| 2026-07-26 | **メニュー8 v1実装**: Z→7.5・空欄のみ採用。次＝clasp push→HUMAN_RUN。 |
| 2026-07-26 | **メニュー8要件＋承認パッケージ**: Amazon AI生成＆一括採用（M-A・空欄のみ・要確認列ごと）。実装承認待ち。 |
| 2026-07-26 | **C1-1b実装**（master_csv・必須列・タックスはマスタ）。次＝未送信SKUでSC。 |
| 2026-07-26 | **C1列マップ下書き**（成功PACKAGED差分・SC必須列）。次＝C1-1b。 |
| 2026-07-26 | **C1実装承認**（`tools/c1_hpc_packaged`・HUMAN_RUN手順）。次＝実機。 |
| 2026-07-26 | **C1三点反映**（MAJORITY・URL空スキップ・親一式除外・指紋v1本番停止）。次＝実装承認。 |
| 2026-07-26 | **C1要件起草**（ローカル本線・HPC・DRY_RUN・03新規）。次＝3者。 |
| 2026-07-26 | **U4 実機合格**: 21-⑦ `U4_20260726_090920_1366af`・マスタ URL 確認。D冪等は想定内。 |
| 2026-07-26 | **U4 v1 実装**: 21-⑦・GENERATED Amazon URL優先。次=clasp push＋HUMAN_RUN。 |
| 2026-07-26 | **U4要件＋承認パッケージ起草**。次＝社長実装承認。 |
| 2026-07-26 | **T2再検証合格**: `80s10` URL単独・18320なし・店頭OK。T3急がない。候補退避 clasp push済。 |
| 2026-07-25 | **U2 実機合格**: ②③④・`02` MAIN1件。Dは冪等ブロック想定内（`LV4_20260725_094425_914290`）。 |
| 2026-07-25 | **C子レ点選定**＋候補`07` Property済。次=clasp push→HUMAN_RUN。 |
| 2026-07-25 | **U2三点＋社長回答反映**（MAJORITY新規）。コミット待ち。 |
| 2026-07-25 | **U2方針確定**: 案α本線・MAIN=sheet／02=出口・εバックログ。 |
| 2026-07-25 | **U3実機合格**＋**U2要件起草**。runId `LV4_20260725_072635_948892`。 |
| 2026-07-25 | **D×Amazon U3 v1**: D `amazon`/`full_amazon`・即時ファサード。clasp push待ち。 |
| 2026-07-25 | **D×Amazon U0クローズ**: 3者反映＋社長回答。手ZIP正／T3実装待ち／将来API必須。 |
| 2026-07-24 | **D×Amazon要件U0**: [org/D_MENU_AMAZON_FACADE_REQUIREMENTS.md](org/D_MENU_AMAZON_FACADE_REQUIREMENTS.md)。T3保留。次=社長承認→3者。 |
| 2026-07-24 | **T2 PoC成功**: runId `R2T2_20260724_221107_7f9cf7`・URL画像表示。トグルfalseへ。次=T3要承認。 |
| 2026-07-24 | **T2 clasp push済**: 8 files（`AmazonDriveImageExport.js`含む）。次=Property＋21-⑥→URL200→トグルoff。 |
| 2026-07-24 | **帰宅引き継ぎ**: §0全面更新。FOOD成功・Nav PASS・T2実装済・次=自宅 clasp push＋21-⑥。 |
| 2026-07-23 | **帰宅引き継ぎ**: HPC suburl試験UP中・FOOD v4再解析（再UP禁止）。§0全面更新。 |
| 2026-07-23 | **§11.0 HPCクローズ**: 1〜8・10＋U5。画像ZIP優先。正本 titlefix。FOOD／M2／21-⑤は別ゲート。 |
| 2026-07-22 | **帰宅**: R2＋8枚MAINは200確認。§0をURL埋め→SC再UP向けに更新。 |
| 2026-07-22 | **自宅IDE引き継ぎ**: DRY_RUN＋本GENERATED成功（..._B2）。当時はPACKAGED→SC。 |
| 2026-07-21 | **セッション終了（深夜）**: clasp push成功。当時は次＝ドライラン（完了済み）。 |
| 2026-07-21 | **Lv4**: docsを新subBatchIdに統一。clasp pushは invalid_rapt で未完了 → clasp login 後に再実行。 |
| 2026-07-20 | **Lv4 実装**: AmazonApprovalExport／ApprovalQueue amazon加算／メニュー21。次は人間検収。 |
| 2026-07-20 | **Lv4 第3回三点＋Q15/Q16**: TRACK未設定＝非実行／amazon抽出同一チケット／列メモ注記。次は実装承認。 |
| 2026-07-20 | **Lv4 Q11–Q14反映**（inventoryMode・親単位再試行・GTIN証跡シート・送料無料パターン）。方針Q&Aは一通り閉じた。次は三点レビュー再実施→実装承認。 |
| 2026-07-20 | **Lv4 Q7–Q10b反映**（3モール同一承認①・TRACK=B強制・出品者SKU=子SKU／メーカー品番）。次は残未決→実装承認。 |
| 2026-07-20 | **Lv4社長Q&A反映**（純正xlsm・D-1 PACKAGED・BはマスタJAN残し・再生成上書き＋ログ追記）。次は残Q（Lv1抽出等）→実装承認。 |
| 2026-07-20 | **Lv4三点レビュー反映**（[org/LV4_THREE_REVIEW_MAJORITY.md](org/LV4_THREE_REVIEW_MAJORITY.md)）。在庫書込禁止・DONE分離・親SKU抽出・GTIN着手ゲート。次は実装承認。 |
| 2026-07-20 | **Lv4要件ドラフト**（[org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md](org/LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md)）。M1=Bノーブランド新規／M2=A既存。 |
| 2026-07-20 | **Lv3人間検収完了**（本番 runId LV3_20260720_112238_777998）。フォーカスを Lv4 Amazon へ。在庫列「在庫数／在庫数計算」分離を§8.2相当で記録。 |
| 2026-07-20 | **Lv2案A修正**: 一時レ点を親＋子に変更（バリエーション商品のシングルSKU誤認を解消）。CSV本体非改変。次は再検収（clasp push → 19-①）。 |
| 2026-07-19 | **Lv2実装**（RakutenApprovalExport.js・メニュー19・案A）。次は人間検収。手動楽天CSVは非改変。 |
| 2026-07-17 | **Lv1人間検収完了**（batch A1_20260717_224813_ad7e65 → APPROVED）。フォーカスを Lv2要件確認／実装承認へ。 |
| 2026-07-17 | **Lv2要件ドラフト**（[org/LV2_RAKUTEN_ORCHESTRATION_REQUIREMENTS.md](org/LV2_RAKUTEN_ORCHESTRATION_REQUIREMENTS.md)）。実装は Lv1検収後。 |
| 2026-07-17 | ハイブリッド・残リストA・実行分割仮既定（25分／ユニーク50）を§3に反映。 |
| 2026-07-17 | **Lv1実装**（ApprovalQueue.js・メニュー18・Web approval_queue）。次は人間検収。 |
| 2026-07-17 | **Lv1要件**追加（[org/LV1_APPROVAL_QUEUE_REQUIREMENTS.md](org/LV1_APPROVAL_QUEUE_REQUIREMENTS.md)）。 |
| 2026-07-17 | **Lv0最終承認**反映。モール順楽天先。フォーカスを Lv1 へ。上書きは当面手動(U1)。 |
| 2026-07-17 | **Lv別実装プラン叩き台**追加（[org/LEVELLED_IMPLEMENTATION_PLAN.md](org/LEVELLED_IMPLEMENTATION_PLAN.md)、コードなし）。 |
| 2026-07-17 | 並行: 楽天ジャンル Nav Stage1 実装（専用診断シート追記・Property 既定オフ）。 |
| 2026-07-15 | **AI組織 Phase0**: 3者多数決反映・RUNBOOK（親1＋並列3）・多数決メモ追加。実装は次フェーズ。 |
| 2026-07-15 | **AI組織 Phase0**: org 憲章・承認マトリクスをフォーカスに。実装は次フェーズ。 |
| 2026-03-19 | 初版。全体＋現在フェーズの引き継ぎを本ファイルに集約。 |
| 2026-03-22 | OpenAI 失敗時の 11-③ バリエーション継続（既定）を §3 に反映。 |
