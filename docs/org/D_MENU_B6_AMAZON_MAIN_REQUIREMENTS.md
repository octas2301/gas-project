# B-⑥ Amazon MAIN — 要件（P0）

**文書種別**: requirements  
**状態**: 2026-09-05 実装（メニュー入口）。エンジンは既存 Pillow 本線  
**関連**: [B45_MAIN_SUB_ALIGN_REQUIREMENTS.md](B45_MAIN_SUB_ALIGN_REQUIREMENTS.md) §1／[D_MENU_SET_MAIN_IMAGE_PHASE_A_HUMAN_RUN.md](D_MENU_SET_MAIN_IMAGE_PHASE_A_HUMAN_RUN.md)

---

## 1. 役割

| 項目 | 正 |
|------|-----|
| メニュー | **B-⑥ Amazon MAIN（透過確認＋Cursor指示）** |
| 仕事 | 出品CKレ点の Amazon セットMAINを Pillow で量産（`01` 白抜き → `07` `{子SKU}_amazon.jpg`） |
| GAS | **合成しない**。パターン選択＋Cursor／Agent 指示のみ |
| マップ | **しない**（楽天 B-④ のテーマ割当とは別。Amazon と同時マップしない） |
| 旧入口 | C「セットMAIN量産（Amazon）」→ **廃止（モーダルから削除）**。dispatch 互換のみ |

## 2. C との境界

| メニュー | 役割 |
|----------|------|
| **B-⑥** | Amazon MAIN **本線入口** |
| **C** | 仕上げ・救済・補助（金丸楽天一括／Ama①②／楽天マトリクス／前提チェック） |
| **B-④ Web** | 楽天金丸の常用（弱い統合） |

## 3. 人フロー（P0）

1. 正本を `01.amazon白抜きベース` へ（透過PNG＋Canva背景リムーバー）
2. マスタで出品CK（親＋使いたい子）
3. **B-⑥** → パターン（正方形／縦長／横長）→ 指示文コピー → Cursor  
   - **ミックス**（セット構成2種以上・数量同一）は選んだパターンに関わらず **列積み**（`mix_column_layout`）。単一種だけ aspect 別本線。
4. `check_amazon_base_matte.py --require-ok` → `amazon_paste_batch.py … --aspect …`
5. 07 を目視 → C「Ama新カタログ①」以降／D

## 4. 実装メモ

- `menuB6AmazonSetMainCursorPrompt` … パターン選択ダイアログ
- `runB6AmazonSetMainWithAspect` … HtmlService 公開入口
- `menuCSetMainAmazonCursorPrompt_` … 指示文本体（関数名互換）
- ミックス列積み: `mix_column_layout.py`／`run_paste_mix_file`。切戻し `layout_rules.json` → `mixColumnLayout.enabled=false`

切戻し: メニュー行を消す／C に Amazon ラジオを戻す（Git revert）。
