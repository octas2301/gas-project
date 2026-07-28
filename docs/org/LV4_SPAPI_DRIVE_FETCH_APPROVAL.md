# SP-API v1.3 Drive自動取得 — 実装承認パッケージ

**日付**: 2026-07-28  
**状態**: **承認済・実装済・実機合格**（2026-07-28／29・`--fetch-drive` dry_run／prod）  
**前提**: 橋渡し v1.2 実機合格  
**手順**: [D_MENU_SPAPI_SHEET_BRIDGE_HUMAN_RUN.md](D_MENU_SPAPI_SHEET_BRIDGE_HUMAN_RUN.md) § Drive取得

---

## 1. 目的

21-⑧／21-⑨ で Drive に出た `*_SPAPI_ITEMS.csv` を、ローカルが **自動取得**して dry_run／prod。手ダウンロード省略。

**同梱 v1.2a**: CSV 文字コード UTF-8／cp932 自動判別。

---

## 2. 変更ファイル

| 種別 | パス |
|------|------|
| 新規 | `tools/spapi_listings_write/spapi_fetch_drive_csv.py` |
| 改修 | `spapi_listings_write.py`（`--fetch-drive`／文字コード） |
| 改修 | `config.example.json`／`requirements.txt` |
| 新規 | 本ファイル |
| 更新 | HUMAN_RUN／CURRENT_PHASE／HANDOVER／CHANGE_LEDGER |

**やらない**: GAS からの SP-API PUT、全件、在庫>0無人

---

## 3. 仕様

- OAuth: C1 と同じ `credentials.json` 流用可（drive.readonly）
- 最新: フォルダ内 `name contains SPAPI_ITEMS` の最新 CSV（または `file_id`）
- 保存先: `items.csv`（gitignore）
- `python spapi_listings_write.py --fetch-drive --mode dry_run`

---

## 4. 社長承認欄

- [x] **承認する**（おすすめ順・v1.3／2026-07-28）  
- [ ] 却下／条件付き
