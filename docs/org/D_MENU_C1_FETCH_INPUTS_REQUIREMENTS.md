# C1 入力自動取得（案A: fetchスクリプト）— 要件

**文書種別**: requirements  
**最終更新**: 2026-07-27  
**状態**: **実装済**（ローカル Python）  
**手順**: [D_MENU_C1_HUMAN_RUN.md](D_MENU_C1_HUMAN_RUN.md) §1b  
**親**: [D_MENU_C1_PACKAGED_XLSM_REQUIREMENTS.md](D_MENU_C1_PACKAGED_XLSM_REQUIREMENTS.md)  
**ツール**: `tools/c1_hpc_packaged/c1_fetch_inputs.py`

---

## 0. 目的

人間が毎回やっていた次を **1コマンド**にする。

1. Drive 上の `*_GENERATED.csv` をローカル `input/` へ取得  
2. ▼商品マスタを CSV 相当でローカル `input/` へ取得  

その後は従来どおり `c1_packaged.py`（dry_run／prod）。

---

## 1. スコープ

### 1.1 作るもの

| # | 成果 |
|---|------|
| 1 | `c1_fetch_inputs.py`（OAuth・Drive／Sheets 読取のみ） |
| 2 | `config` の `fetch` ブロック（example 更新） |
| 3 | HUMAN_RUN／台帳 |

### 1.2 作らないもの

- GAS 改変・楽天／Yahoo 聖域  
- C1 PACKAGED 本体の仕様変更  
- サービスアカウント必須化（v1 は **ユーザ OAuth**）  
- Drive へのマスタ書込・GENERATED 削除  

---

## 2. 入力契約（config）

| キー | 必須 | 内容 |
|------|------|------|
| `spreadsheet_id` | ○ | 出品用スプレッドシート ID |
| `fetch.credentials_path` | ○ | GCP OAuth クライアント JSON（Desktop） |
| `fetch.token_path` | ○ | 初回認可後の token（Git 禁止） |
| `fetch.input_dir` | ○ | 保存先（例: `…/input`） |
| `fetch.master_sheet_name` | ○ | 既定 `▼商品マスタ(人間作業用)` |
| `fetch.log_sheet_name` | △ | 既定 `▼Lv4実行ログ(Amazon)` |
| `fetch.generated_folder_id` | △ | 名前検索用。空ならログの fileUrl 優先 |
| `generated_csv` / `master_csv` | △ | 取得後にパスをこのファイル名で上書き保存（config 記載パスへ） |

CLI:

| 引数 | 内容 |
|------|------|
| `--sub-batch ID` | `{ID}_GENERATED.csv` を取得（folder または Drive 名検索） |
| `--generated-file-id ID` | Drive ファイル ID 直指定 |
| `--latest` | ログシート最新の `GENERATED`＋fileUrl を採用 |
| `--skip-master` / `--skip-generated` | 片方だけ |

---

## 3. 処理フロー

```text
OAuth（drive.readonly + spreadsheets.readonly）
  → GENERATED:
       --generated-file-id
       なければ --sub-batch で folder_id 内ファイル名一致
       なければ --latest でログ sheet の fileUrl から ID 抽出
  → Drive files.get_media → input/{subBatchId}_GENERATED.csv
  → マスタ:
       Sheets values を CSV 化 → input/master_export.csv
       （手作業「ファイル→ダウンロード→CSV」と同等の列内容を目指す）
  → 終了コード 0／ログにパス
```

---

## 4. 検収

- [ ] 初回 OAuth 後、2回目はブラウザなしで取得できる  
- [ ] `--latest` または `--sub-batch` で今回の B2 GENERATED が取れる  
- [ ] master CSV に `親SKU`／`子SKU`／`Amazon MAIN URL` 見出しがある  
- [ ] 続けて `c1_packaged.py --mode dry_run` が通る（指紋・テンプレは従来）  
- [ ] `credentials.json` / `token.json` が Git に含まれない  

---

## 5. 戻し方

- スクリプト・docs を revert  
- ローカル token 削除  
- 手ダウンロード運用に戻す  

---

## 6. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-27 | 案A 要件＋実装。 |
