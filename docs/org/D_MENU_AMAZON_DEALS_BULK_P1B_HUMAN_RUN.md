# P1b — Amazonタイムセール人間手順

**状態**: 2026-08-12 P1c実装済（2層減衰・メニュー分割）。clasp push 済。レーンAは**未運用**／数量確認メール§9.7。次は 8/14 taper 1SKU prod  
**要件**: [D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md](D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)（§0.1 不変条件・§2.1.0 事実記載）

---

## GASメニュー（2026-08-12）

広告スプシに **2本**。タイムセールは `🚀 広告運用` から分離。減衰の本線送信は **日次 `taper_send.py --poll`**。手動送信は 99-⑦。

### ⏱ Amazonタイムセール・ポイント減衰設定（本線）

| # | メニュー | いつ |
|---|---------|------|
| **1.** | 公式タイムセールの提出xlsxを作る（Agent依頼文） | 公式SaleがSCに出たとき |
| **1-②** | 提出xlsxを作り直す（修正後・Agent依頼文） | スプシ修正後 |
| **2.** | セール開始の21日前・14日前：数量確認メールを送る（下書き送信） | T-21／T-14。下書きは先に Python |
| **3.** | セール開始前日／終了翌日：ポイント作業の催促メールを送る（下書き送信） | 開始前日／終了翌日。下書きは先に Python |
| **4.** | セール開始時：期間中ポイント%をAmazonに載せる（Agent依頼文） | セール開始 |
| **5.** | セール終了後：減衰中ポイント%をAmazonに戻す（Agent依頼文） | セール終了後（カレンダー減衰中%へ。終着へ一気に戻さない） |
| **6.** | 新しいSKUに、何日かけて何%ずつ下げるかを記入する（選択行） | 目標売価・販促%入力後の最初の1回 |
| **6-②** | 予定日より前に、今すぐ1段下げたいSKUに印を付ける（選択行） | 期日前に下げたいとき → 次の日次が拾う |

### 99.テスト用（タイムセール）

| # | メニュー | いつ |
|---|---------|------|
| **99-①** | 【危険】タイムセールシートを初期化する | タブ初期化のみ |
| **99-②** | 列の並びを目的グループに直す | 列が崩れたとき |
| **99-③** | ヘッダ色・入力黄・計算式を付け直す | 色・数式が消えたとき |
| **99-④** | 期間・間隔・実行依頼の入力規則を直す | プルダウン／チェック破損 |
| **99-⑤** | 計画が空の有効SKUへ、減衰の進め方を一括記入する（初回向け） | 初回一括 |
| **99-⑥** | 提出xlsxは作らず、タイムセール表だけ最新にする（Agent依頼文） | xlsx不要・表だけ更新。**1.の中にも sync あり** |
| **99-⑦** | 減衰を手動で1回回す（日次失敗時・Agent依頼文） | 日次失敗・PCオフ翌日 |

### 🚀 広告運用（タイムセール抜き）

| # | メニュー |
|---|---------|
| **1.** | 広告レポートを分析し、除外・入札案を出す |
| **1-②** | 分析して広告バルクまで一気に作る |
| **2.** | 広告バルクをDriveに保存する |
| **2-②** | 広告バルクの誤りをチェックする |
| **3.** | 楽天お宝キーワードの広告CSVを作る |
| **99-①** | 操作マニュアルシートを作る |
| **99-②** | 古いDriveファイルを削除する |

---

## フロー要約（公式Sale＝出たら即登録）

```text
① SCで推奨バルクDL → 必ずフォルダ②へ保存（最新価格・候補のため）
   フォルダ名: 02.amazon公式タイムセールバルクファイル保存（人間が保存）
② メニュー「1. 公式タイムセールの提出xlsxを作る（Agent依頼文）」
   → ②にxlsxがあるか確認 → Cursor文をコピー（ダイアログ下部にUL固定手順）
③ Agent: sync --write → build_submit_xlsx.py --write → フォルダ③
   ・②の当該SKU **スケジュール列ドロップダウンの日付付き全候補**＋セル値を正（カタログだけのBF等は入れない）
   ・作成成功で状態=UL済（再作成可）。Smile等の再ULしたくない枠は提出対象=いいえ
   ・カスタムは公式と非重なりで提出対象へ。公式と衝突したら独自側を短縮しメール通知
   ・レーンA（discounted_price）は未運用
④ 人: ③を確認 → SCへUL（手順は②のダイアログ下部）→ 反映待ち
⑤ 修正が必要ならスプシを直す → 「1-② 提出xlsxを作り直す（修正後・Agent依頼文）」
⑥ 開始21日前・14日前: 「2. …数量確認メールを送る」→ 改定はSC画面編集（§9.7）
```

**方針**: 名付き公式は候補に出たら即登録。レーンA（`discounted_price`）は **未運用**（生成しない）。

---

## 数量確認メール（§9.7）— 本線

| タイミング | 内容 |
|------------|------|
| **T-21** | 第1報。再計算案と現状。改定是非 |
| **T-14** | 最終＋FBA。改定なら **SC画面編集**（バルク再UL不可） |

対象: レーンBで状態が `予定`／`UL済`／`実施中` など（**提出対象=いいえの登録済も含む**）。

```text
cd tools/amazon_deals_bulk
python mail_qty_confirm.py                    # 今日がちょうど T-21/T-14 の行
python mail_qty_confirm.py --days 14
python mail_qty_confirm.py --days 14 --tol 2  # ±2日で拾う（運用ゆらぎ）
python mail_qty_confirm.py --today 2026-08-14 --days 14   # 日付指定の予行
python mail_qty_confirm.py --days 14 --tol 3 --send       # 実送信（下書き＋送信）
```

**送信経路（優先順）**

1. SMTP（`config.local.json` の `smtp_host` 等）
2. **Gmail API**（初回ブラウザ同意 → `tools/c1_hpc_packaged/secrets/token_gmail_send.json`。GCPで Gmail API 有効化が必要）
3. GAS WebApp（`qty_mail_web_url`）／メニュー **2. セール開始の21日前・14日前：数量確認メールを送る（下書き送信）**

宛先既定: `contact@octas2301.com`（`notify_email_to`）  
**件名（固定）**: `[amazonタイムセール] 数量確認 （スプシで確認・改定はSC画面）`  
下書きファイル: `tools/amazon_deals_bulk/_work/qty_confirm_T*.txt`  
下書きシート: 広告スプシ `⏱数量確認メール下書き`  
リンク先はタイムセール行。改定は PC の SC Deal 編集。

**Smile（2026-08-28開始）目安**: T-21≈8/7（過ぎていれば第1報スキップ可）／**T-14≈8/14** に `--days 14` を実行。  
**2026-08-11 検証**: `--days 14 --tol 3 --send` で Smile×2 を拾い、**Gmail API で contact@octas2301.com へ送信成功**。

---

## BF

- **当該商品の②行／候補にBFが出てから**登録（カタログだけのBFは自動追加しない）  
- 早期割引は候補出現後〜**9/30**

---

## 公式B提出xlsx（GASメニュー）

1. **SC** → 推奨バルクDL → **②** へ保存  
2. **1. 公式タイムセールの提出xlsxを作る（Agent依頼文）** → Cursorへ貼付（下部のUL手順）  
3. Agent: `sync_master_and_exec.py --write` → `build_submit_xlsx.py --write` → **③**  
4. 人: ③チェック → SC UL → 反映確認  
5. 修正後: **1-② 提出xlsxを作り直す（修正後・Agent依頼文）**  

**提出xlsxの形（成功UL準拠: `DEALS_…_114817_submitv2`）**:
- **対象SKU行のみ残す**（他行削除＝compact）。全行残し＋結合解除は VLOOKUP 参照壊れで「ファイル処理失敗」になる
- **開始/終了は YYYY-MM-DD 直書き**（書式 `yyyy\-mm\-dd`）。テンプレの VLOOKUP は提出ファイルでは使わない
- **シート保護OFF**、参加中/スケジュールの DV を残り行だけ再構築
- **同一ASINは1ファイル1スケジュール**（8月と9月は別xlsx）
- ②は最新のみ残し、古いDLは `02/使用済み/` へ（`--archive-02`）

```text
# 8月カスタムのみ（既定が成功形）
python build_submit_xlsx.py --write --only-schedule "カスタム - (2026-08-12 - 2026-08-13)"
# 9月カスタムのみ
python build_submit_xlsx.py --write --only-schedule "カスタム - (2026-09-18 - 2026-09-24)"
```

手でSKU行を削った提出ファイルや、旧 `--in-place` 全行残しは使わない。ULは③の最新をSCへ。

---

## ポイント §10.10 Phase0

| 列 | 意味 |
|----|------|
| `期間中ポイント%`／円 | タイム期間中の出品者付与（%空＝1） |
| `セール前ポイント%`／円 | **最終終着%**（減衰フロア。例: 1%）。**restore 先ではない**。fetch で埋めない |
| `出品者ポイント現在%`／円 | 直近同期後 |
| `ポイント状態`／`ポイントメモ` | 状態・注記（語彙は下表／要件§10.10） |

**ポイント状態（正）**: `セール前退避済`／`期間中適用済`／`セール前復元済`／`フィード{STATUS}`。空＝未設定。  
※ `セール前復元済`＝**減衰中%へ戻した**直後（終着へ一気に戻した、ではない）。

```text
cd tools/amazon_deals_bulk
python points_fetch.py --write                 # 出品者現在%スナップのみ（終着は触らない）
python points_send.py --mode apply             # dry_run: TSVのみ
python taper_send.py --poll --mail             # B終了後は先にカレンダー同期
python points_send.py --mode restore           # 減衰中%へ
# 承認後（EC書込）
python points_send.py --mode apply --prod --i-confirm-prod --wait --update-sheet
python points_send.py --mode restore --prod --i-confirm-prod --wait --update-sheet
# 既定は施策Bの接近/直後SKUのみ。全マスタは --all-master
```

**Smile（2026-08-28〜09-03）目安**（連携は要件§10.13）:

| いつ | 何 |
|------|-----|
| 〜8/27 | `points_fetch --write`（現在%スナップ。終着は触らない。Cinderellas b/s は店頭1%・減衰中22%・開始日8/14） |
| 8/14 | カレンダー1段目: 減衰中 22→18（現行TS終了後。B中でなければ Amazon 18%） |
| 8/27〜28 | `points_send --mode apply` → 承認後 `--prod …`（**現在も1%なら差分なし・SC確認のみ可**） |
| Smile中 | 店頭1%。カレンダー減衰中は進む（8/28→14%想定。シートのみ） |
| 9/4〜 | `taper_send --poll` → `points_send --mode restore`（**14%へ**。22%へ戻さない） |

価格戻しGAS提案はいつでも可。**戻しAPIはB終了＋restore後**（§10.13）。

### Phase0 残ギャップ（2026-08-11 洗い出し）

| ID | ギャップ | 重要度 | メモ |
|----|----------|--------|------|
| P0-G1 | ~~本番フィード未検収~~ → **2026-08-11 検収済**（1SKU往復 2%→1%） | — | feed `185494020676`（2%）／`185495020676`（1%）共に DONE。SKU=`originalM-1803--KOUS--Cinderellas b`。SC目視は人間確認 |
| P0-G2 | ~~日程トリガー無し~~ → **リマインドで代替（自動実行は持たない）** | — | `mail_points_remind`：apply=T-1／restore=終了+1。自動 apply/restore は当面しない |
| P0-G3 | ~~リマインドメール無し~~ → **2026-08-11 実装** | — | `mail_points_remind.py`＋下書きシート＋GASメニュー（要 clasp push） |
| P0-G4 | ~~対象SKUがマスタ有効全件~~ → **2026-08-11 施策連動** | — | `sale_skus_for_points`＋`points_send`既定。解除=`--all-master`／`--sku` |
| P0-G5 | ~~GASメニュー無し~~ → **2026-08-11 Cursor指示で対称化** | — | apply／restore は SP-API のため GAS本体では送らない。メニューは Cursor指示＋リマインド送信。要 clasp push |
| P0-G6 | **ポイント円列はフィード未使用**（%のみ送信） | 低 | 列は監査・表示用。仕様どおりで可 |
| P0-G7 | ~~状態語彙が未固定~~ → **2026-08-11 固定** | — | `セール前退避済`／`期間中適用済`／`セール前復元済`（＝減衰中%へ戻した）／`フィード*` |
| P0-G8 | ~~fetch失敗時も apply 続行~~ → **2026-08-11 ガード** → **2026-08-12 改訂** | — | 終着空でも apply 可（フロア1%）。fetch は現在%のみ。`--allow-missing-before` は互換残 |
| P0-G9 | ~~listings GETで%が取れない~~ → **2026-08-11 本線確定** | — | `offers[].points.pointsNumber`÷price→%。Cinderellas b/s 実機1%。GETフィードは環境により400で任意フォールバック |
| P0-G10 | ~~§10.13連携の実装ガード無し~~ → **2026-08-11 実装** | — | `price_recovery_send`: B中スキップ／期間中適用済は中止。解除=`--allow-during-b`／`--allow-before-restore` |

**次**: 運用（8/14 taper 1SKU prod → 8/27 Smile apply リマインド → 9/4 restore＝減衰中%）。開発後続は P1c-2／C／D／任意 P2。レーンAは未運用。  

## ポイントリマインド（G2/G3）

```text
cd tools/amazon_deals_bulk
python mail_points_remind.py --kind apply --days 1          # 開始前日
python mail_points_remind.py --kind restore --days 1        # 終了翌日
python mail_points_remind.py --kind both --tol 1            # 予行
python mail_points_remind.py --kind apply --days 1 --send   # 送信
```

既に `期間中適用済`／`セール前復元済` のSKUは通常スキップ。再送・予行は `--include-done`。  
未適用なら **現在%＝期間中%でも催促する**（差分ゼロ注記付き。SC確認／送信スキップ可）。  
GAS（`⏱ Amazonタイムセール・ポイント減衰設定`）:
- **3.** セール開始前日／終了翌日：ポイント作業の催促メールを送る（下書き送信）
- **4.** セール開始時：期間中ポイント%をAmazonに載せる（Agent依頼文）
- **5.** セール終了後：減衰中ポイント%をAmazonに戻す（Agent依頼文）（P0-G5・SP-APIはCLI）
※要 `clasp push`（広告運用GAS）。

### 運用予行ログ（2026-08-11）

| 項目 | 結果 |
|------|------|
| Smile apply T-1（`--today 2026-08-27`） | **2件**下書き生成（b/s）。Draft: `_work/points_remind_apply_D1_*.txt`＋シート`⏱ポイントリマインド下書き` |
| Smile restore 終了+1（`--today 2026-09-04`） | 該当なし（現状は退避済・未 apply のため正常） |
| マスタ b | G1後にセール前%空→`points_fetch --write`で **1%退避**・状態=セール前退避済 |
| マスタ s | セール前%=1／状態=セール前退避済（先に退避済） |
| バグ修正 | 旧: `needs_sync` 依存で「既に1%」が全スキップ → **日程ヒット＋未適用なら催促**に変更 |

**人間残り**: 8/14 taper 1SKU prod（都度承認）。8/27 に `--send` または **3.** 催促メール送信。

### G1 本番検収ログ（2026-08-11）

| 回 | 送信% | feedId | processingStatus |
|----|-------|--------|------------------|
| 1 | 2 | `185494020676` | DONE |
| 2 | 1 | `185495020676` | DONE |

メタ: `tools/amazon_deals_bulk/_work/points_send_apply_20260811_123812.json`／`…123850.json`  
マスタ: 期間中%=1／出品者現在%=1／状態=期間中適用済（`--update-sheet`）

---

## 実質戻し §10.12／次開発 P1c §10.14

**正本**: [要件 §10.12–10.14](D_MENU_AMAZON_DEALS_BULK_REQUIREMENTS.md)

1. マスタに **目標売価円**・**販促ポイント%**・**減衰開始日** を入力（例: 4480／22／2026-08-14）  
2. **販促ポイント円・実質価格円**はシート数式で常時表示  
3. メニュー **6.** 新しいSKUに、何日かけて何%ずつ下げるかを記入する（選択行）→ **減衰期間／減衰段%／減衰間隔** を確認  
4. 人が見る進捗: `出品者ポイント現在%`（店頭スナップ）／`減衰中ポイント%`（カレンダー運用目標）／`次回減衰後%`／`次回減衰日`／`減衰開始日`  
5. 要 `clasp push`（`TimeSalePriceRecovery.js`）

タイムセールの赤札は **Bのみ**。売価段階上げは廃止。

**運用**: 週次CLなし。減衰後は **結果メール**（商品URL）を見て異常時のみ **手動**修正（SC or `points_send --sku`）。リカバリURLは後続。

**起動**: 本線は日次 `taper_send.py --poll --mail`。期日前は **6-②** で印 → 次の日次。失敗時は **99-⑦**。最初1週間は dry_run。C／Dは後続。

### ポイント減衰送信（§10.14／`taper_send.py`）

```text
cd tools/amazon_deals_bulk
python taper_send.py --poll --mail
python taper_send.py --start --sku "…" --mail
python taper_send.py --sku "…" --prod --i-confirm-prod --wait --update-sheet --mail
python price_recovery_send.py --snap --sku "…"        # 売価スナップ（別）
```

| ガード | 既定 | 解除 |
|--------|------|------|
| B期間中 | Amazonスキップ・シートはカレンダー更新 | `--allow-during-b` |
| ポイント=期間中適用済 | Amazonスキップ・シートはカレンダー更新 | `--allow-before-restore` |
| 本日実行済 | スキップ | `減衰実行依頼=TRUE` |

GAS: **6-②** 今すぐ印／**99-⑦** 手動1回（日次失敗時）。※要 clasp push。  
日次: `python taper_send.py --poll --mail`（安定後に `--prod --i-confirm-prod --wait --update-sheet`）。  
dry_run ではフラグを消さない（同じ計画が毎日メールされる）。本番成功後に `減衰実行依頼` をクリア。

### 日次タスク（Windows・人手で登録）

タスクスケジューラ例（最初1週間は dry_run）:

```text
プログラム: python
引数: taper_send.py --poll --mail
開始: C:\Users\takuy\Desktop\gas-project\tools\amazon_deals_bulk
時刻: 例 09:00（PC起動中のみ）
```

PCが寝ていると動かない。HTTP化（C）は後続。

### マスタ列並び・色分け

1. **99-②** 列の並びを目的グループに直す … 人入力を各グループ先頭に。`最終売価円`→`目標売価円`
2. **99-③** ヘッダ色・入力黄・計算式を付け直す（または `sync --schema-only`）: 基本灰／SC青／ポイント紫／**実質戻し琥珀**／A緑／メモヘッダ黄。人入力セル(行2〜)=`#FFF2CC`

---

## 新Sale検知（月次バルク後）

```text
python detect_new_schedules.py
python detect_new_schedules.py --save
```

---

## レーンA（未運用・触らない）

**2026-08-11 確定: 運用しない。** sync は A行を生成せず削除する。  
過去に `lane_a_send` の検証ログが残っていても **本線手順ではない**（要件§2.1.0）。再開は社長明示＋要件改訂後のみ。

---

## コマンド（本線）

```text
cd tools/amazon_deals_bulk
python ops_status.py
python sync_master_and_exec.py --write
python build_submit_xlsx.py --write
python mail_qty_confirm.py --days 14 --tol 2
python revise_b_qty.py --within-days 21
python detect_new_schedules.py
python points_send.py
python taper_send.py --poll --mail
```
