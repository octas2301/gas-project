# プロジェクト別 Drive 自動バックアップ — HUMAN_RUN

**正本**: [PROJECT_DAILY_BACKUP_REQUIREMENTS.md](PROJECT_DAILY_BACKUP_REQUIREMENTS.md)  
**日付**: 2026-10-17（2026-11-15: G1＋G3・次段順を追記）

Agent は `clasp push` しない。人間が PowerShell で行う。

**確定方針**: スプシ＝Drive日次（全PJ横展開）。コード＝**G1（毎日 main push）＋G3（全claspローカルsnapshot）**。次段順＝広告→在庫→競合Jitta→税務。

---

## 0. 共通（初回）

Drive フォルダは **2026-10-17 作成済**（My Drive 直下）。

| フォルダ | ID | URL |
|----------|-----|-----|
| `99.バックアップ` | `1K2OF1hpA5mfugJIAPZCR6VkZ0zzmYFZE` | https://drive.google.com/drive/folders/1K2OF1hpA5mfugJIAPZCR6VkZ0zzmYFZE |
| `出品` | `14Cd0yoQ8BWrB-1SHbC3XyhV00A-mTxYh` | https://drive.google.com/drive/folders/14Cd0yoQ8BWrB-1SHbC3XyhV00A-mTxYh |
| `仕入れ検討①` | `1Uz1Px-7UEFi2MmPylMS4ZNCPKwtGdPJe` | https://drive.google.com/drive/folders/1Uz1Px-7UEFi2MmPylMS4ZNCPKwtGdPJe |
| `広告` | `1CscA4rPU7J69gf8ANgOLttda2z36TmIR` | https://drive.google.com/drive/folders/1CscA4rPU7J69gf8ANgOLttda2z36TmIR |

報告メールを使うなら宛先を Property に入れる（未設定でも ▼ログ／①ログには残る）。

### 0.1 Drive 目視結果（2026-11-15 確認）

API（`drive.readonly`）で `99.バックアップ` 配下を一覧。**今日朝の控えあり**。

| フォルダ | 直近ファイル | modified（UTC） | 備考 |
|----------|--------------|-----------------|------|
| 出品 | `出品_backup_2026-11-15` | 2026-11-14T23:26Z（JST 15日 8:26頃） | 日次トリガー稼働 |
| 仕入れ検討① | `競合DB_backup_2026-11-15` | 2026-11-14T23:05Z（JST 15日 8:05頃） | 今日朝あり |
| 仕入れ検討① | `①調査_backup_2026-11-14` が最新 | — | 15日分は無し＝更新窓内に変更なしでスキップの可能性（要件どおり成功扱い） |

出品フォルダは 11/03〜11/15 までほぼ毎日控えあり。

---

## 1. 出品（第1版・先）

```powershell
cd C:\Users\takuy\Desktop\gas-project
clasp push
```

### Script Properties（出品 GAS）

| Property | 推奨 | 説明 |
|----------|------|------|
| `PROJECT_DAILY_BACKUP_ENABLED` | 省略=ON | `false` で切戻し |
| `PROJECT_DAILY_BACKUP_ROOT_FOLDER_ID` | `1K2OF1hpA5mfugJIAPZCR6VkZ0zzmYFZE` | `99.バックアップ` |
| `PROJECT_DAILY_BACKUP_PJ_FOLDER_ID` | `14Cd0yoQ8BWrB-1SHbC3XyhV00A-mTxYh` | `出品` |
| `PROJECT_DAILY_BACKUP_TARGET_SS_ID` | 任意 | 空なら Active（出品マスタ） |
| `PROJECT_DAILY_BACKUP_RETENTION_DAYS` | 省略=14 | 残日数 |
| `PROJECT_DAILY_BACKUP_SHRINK_RATIO` | 省略=0.8 | 良品比これ未満は破損疑い |
| `PROJECT_DAILY_BACKUP_MAIL_TO` | 任意 | 破損・失敗の報告先 |
| `PROJECT_DAILY_BACKUP_REQUIRED_SHEETS` | 省略=`▼商品マスタ(人間作業用)` | `\|` 区切り |

### メニュー

`AI出品ツール` → `99.日次バックアップ（Drive）`

1. **強制コピー（検収用）** で1回成功を確認  
2. **毎朝8時トリガー設置**  
3. 翌日以降: `99.バックアップ/出品/` に `出品_backup_YYYYMMDD` が増えるか（更新窓内に変更があった日）

### 検収

要件 §7 T1〜T5（出品分）。

---

## 2. 仕入れ検討①（出品の次）

```powershell
cd C:\Users\takuy\Desktop\gas-project\purchase_research
clasp push
```

### Script Properties（① GAS）

| Property | 推奨 | 説明 |
|----------|------|------|
| `PR_PROJECT_DAILY_BACKUP_ENABLED` | 省略=ON | 切戻し `false` |
| `PR_PROJECT_DAILY_BACKUP_ROOT_FOLDER_ID` | `1K2OF1hpA5mfugJIAPZCR6VkZ0zzmYFZE` | `99.バックアップ` |
| `PR_PROJECT_DAILY_BACKUP_PJ_FOLDER_ID` | `1Uz1Px-7UEFi2MmPylMS4ZNCPKwtGdPJe` | `仕入れ検討①` |
| `PR_PROJECT_DAILY_BACKUP_RESEARCH_SS_ID` | 任意 | 空なら Active（調査複製） |
| `PR_PROJECT_DAILY_BACKUP_COMPETITOR_SS_ID` | 任意 | 空ならコード定数の競合DB |
| `PR_PROJECT_DAILY_BACKUP_RETENTION_DAYS` | 省略=14 | |
| `PR_PROJECT_DAILY_BACKUP_SHRINK_RATIO` | 省略=0.8 | |
| `PR_PROJECT_DAILY_BACKUP_MAIL_TO` | 任意 | |

### メニュー

`自動リサーチ` → `99. 運用・切り分け` → 30〜33

競合DBは大きい。初回強制コピーの所要時間に注意。

---

## 3. 切戻し

- 出品: `PROJECT_DAILY_BACKUP_ENABLED=false` またはトリガー削除  
- ①: `PR_PROJECT_DAILY_BACKUP_ENABLED=false` またはトリガー削除  
- 広告: `ADS_PROJECT_DAILY_BACKUP_ENABLED=false` またはトリガー削除  

正本スプシは消えない（コピーと古いコピーのゴミ箱のみ）。

---

## 4. 次段スプシ日次

順: **広告 → 在庫 → 競合Jitta → 税務**。各 PJ でフォルダ作成 → Property → `ProjectDailyBackup.js` 相当 → clasp push → 8時トリガー → 強制コピー1回。

| 順 | PJ | Drive フォルダ | clasp | 対象 SS ID | PJ フォルダ ID |
|----|-----|----------------|-------|------------|----------------|
| 1 | 広告 | `99.バックアップ/広告` | `広告運用GAS/`（push は clone 経由） | `17jwaao79ApQrHzA27BPVIl1sQ9G2tVAqWPkVvrGGelo` | `1CscA4rPU7J69gf8ANgOLttda2z36TmIR` |
| 2 | 在庫 | `99.バックアップ/在庫` | `inventory-app/` | | |
| 3 | 競合Jitta | `99.バックアップ/競合Jitta` | `competitor_jitta/` | | |
| 4 | 税務 | `99.バックアップ/税務` | `export-tax-list/` | | |

### 4.1 広告（2026-11-15 投入）

コード: `広告運用GAS/ProjectDailyBackup.js`（`ADS_` 接頭 Property）。clasp push 済（clone 経由）。

Drive:
| 項目 | 値 |
|------|-----|
| 親 | `99.バックアップ` = `1K2OF1hpA5mfugJIAPZCR6VkZ0zzmYFZE` |
| PJ | `広告` = `1CscA4rPU7J69gf8ANgOLttda2z36TmIR` |
| 正スプシ | [広告運用](https://docs.google.com/spreadsheets/d/17jwaao79ApQrHzA27BPVIl1sQ9G2tVAqWPkVvrGGelo/edit) |
| 検収コピー | `広告_backup_force_20261115_113226`（Drive API・2026-11-15） |

**人間が1回だけ（トリガー／Property）**:

1. 広告運用スプシを開く（新スコープ `script.scriptapp` の承認が出たら許可）
2. `🚀 広告運用` → `99.日次バックアップ（Drive）` → **0. 初回セットアップ（Property＋毎朝8時）**
3. 任意: **1. 強制コピー（検収用）**（GAS 経路の確認）

| Property | 推奨 |
|----------|------|
| `ADS_PROJECT_DAILY_BACKUP_ENABLED` | 省略=ON／`false` で切戻し |
| `ADS_PROJECT_DAILY_BACKUP_ROOT_FOLDER_ID` | `1K2OF1hpA5mfugJIAPZCR6VkZ0zzmYFZE` |
| `ADS_PROJECT_DAILY_BACKUP_PJ_FOLDER_ID` | `1CscA4rPU7J69gf8ANgOLttda2z36TmIR` |
| `ADS_PROJECT_DAILY_BACKUP_TARGET_SS_ID` | `17jwaao79ApQrHzA27BPVIl1sQ9G2tVAqWPkVvrGGelo` |
| `ADS_PROJECT_DAILY_BACKUP_REQUIRED_SHEETS` | 省略=`元データ_Amazon\|履歴データ` |
| `ADS_PROJECT_DAILY_BACKUP_MAIL_TO` | 任意 |

push:

```powershell
Copy-Item -Path "C:\Users\takuy\Desktop\gas-project\広告運用GAS\*" -Destination "C:\Users\takuy\Desktop\広告運用GAS_clone\" -Force
cd C:\Users\takuy\Desktop\広告運用GAS_clone
clasp push
```

---

## 5. GitHub G1（毎日 main）＋ ローカル G3

### 5.1 タスク登録（初回・人間）

```powershell
cd C:\Users\takuy\Desktop\gas-project
powershell -NoProfile -ExecutionPolicy Bypass -File tools\register_backup_tasks.ps1
```

| タスク名 | 時刻 | スクリプト |
|----------|------|------------|
| `gas-project-git-daily-backup` | 08:05 | `tools/git_daily_backup.ps1`（G1） |
| `gas-project-clasp-all-snapshot` | 08:00 | `tools/backup_all_clasp_snapshots.ps1`（G3・全clasp） |
| `gas-project-gas-snapshot` | 08:00 | `tools/backup_gas_snapshot.ps1`（出品ルートのみ・既存） |

ログ:
- G1: `_local_backup/git_daily/_last_result.txt`
- G3: `_local_backup/clasp_all/_last_result.txt`

### 5.2 G1（GitHub 毎日 main）

- 変更があれば ignore 遵守で commit → `git push origin main`
- 秘密・巨大物は `.gitignore`（`**/secrets/` 等）。ステージに秘密が混ざると中止
- 失敗は `_local_backup/git_daily/` にログ
- **初回**: 未 commit の大量差分を一度 commit/push してから自動に任せる

手動1回（確認用）:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\git_daily_backup.ps1 -DryRun
powershell -NoProfile -ExecutionPolicy Bypass -File tools\git_daily_backup.ps1
```

### 5.3 G3（全 clasp ローカル保険）

- 対象: listing / purchase_research / competitor_jitta / ads / inventory / export_tax
- 置き場: `_local_backup/clasp_all/（pj）/yyyyMMdd_HHmmss/`
- 急縮小は `_rejected`（壊れた断面を良品にしない）

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\backup_all_clasp_snapshots.ps1
```
