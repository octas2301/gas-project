# B統合 — ハード死番犬再開（P0）

**文書種別**: 要件定義＋実装  
**日付**: 2026-08-12  
**状態**: **実装済**（要 `clasp push`）  
**親**: [AGENT_HANDOVER.md](../AGENT_HANDOVER.md) §8（B統合境界・ログ標準）／[CURRENT_PHASE.md](../CURRENT_PHASE.md)  
**コード**: `コード.js`（`B_WATCHDOG_*` / `runBWatchdogFromTrigger` / `generateListingDataComparison` / `setTitleDropdownForParentRowsWithGivenRows`）  
**ゴール一文**: ハードタイムアウト後も人間が B を押さず、**最後の保存位置から**自動で続き、Step6・Step7 まで到達できる。

---

## 1. 背景（実機）

| 事象 | 内容 |
|------|------|
| ⑩転記なし | Step6 未到達（Step5 途中ハード死） |
| 別回 | Step7 末尾 `[キーワードプルダウン]` 行処理中にハード死 |
| 穴 | Bの21分は **Step合間のみ**／Step5の22分は **行先頭のみ**かつ B経過を知らない／ハード死では `GEN_LISTING_STATE`・トリガーが残らない |

死んだ実行のメモリは読めない。読めるのは Properties／シート。**別実行（番犬）がそれを読んで続きを起動する。**

---

## 2. 作るもの（P0・実装済）

| # | 内容 |
|---|------|
| 1 | **1発予約**: **メニュー開始時のみ** `B_WATCHDOG_AFTER_MIN`（既定20）分後の `runBIntegratedFromTrigger` を予約。ソフト再開トリガーでは再予約しない（走行中二重起動防止）。同ハンドラの既存は削除してから1本 |
| 2 | **定期番犬**: `runBWatchdogFromTrigger` を `B_WATCHDOG_PERIODIC_MIN`（既定15）分ごと。stale＋状態残＋実行中でない → 続き起動＋メール。ソフト再開後のハード死はこちらがカバー |
| 3 | **Step5 残り時間**: `min(GEN予算, B開始+約25分までの残り)`。Gemini/GPT **前後**でもチェック → `TIME_SLICE` |
| 4 | **行ごと CP**: 各行成功後に `GEN_LISTING_STATE.nextRow`（トグル `B_STEP5_ROW_CHECKPOINT`） |
| 5 | **再開位置**: `nextStepIndex`＋`GEN_LISTING_STATE.nextRow`。**Step1からやり直さない** |
| 6 | **▼B実行サマリ**: 1実行1行（時刻 / runId / phase / nextStepIndex / nextRow / 終了理由） |
| 7 | **`[メインKW]` 詳細**: 既定OFF（`B_VERBOSE_MAIN_KW_LOG=true` でON） |
| 8 | **メニュー「Bの進捗を表示」**: Z→11。`LAST_PROGRESS` と `B_INTEGRATED_RUN_STATE` をダイアログ |
| 9 | **メール**: 番犬が stale 再開したとき（ハード死推定） |
| 10 | **プルダウン TIME_SLICE**: 親行単位で予算超過 → カーソル保存＋`TIME_SLICE`（Step7 は再throw） |

## 3. 作らないもの

- `generateRakutenCSV` / `Yahoo.js` / B Step **順序変更**
- 死んだ実行内での `catch` 再開
- GCP／実行ログAPI必須化
- 完了済み B の自動再実行
- Step1からの自動やり直し

---

## 4. Script Properties

| Key | 既定 | 意味 |
|-----|------|------|
| `B_WATCHDOG_ENABLED` | **未設定=true** | 番犬全体 |
| `B_WATCHDOG_AFTER_MIN` | 20 | B開始から1発予約（分） |
| `B_WATCHDOG_STALE_MIN` | 15 | 定期番犬の「古い」判定（分） |
| `B_WATCHDOG_PERIODIC_MIN` | 15 | 定期間隔（分） |
| `B_STEP5_ROW_CHECKPOINT` | **未設定=true** | 行ごと GEN_LISTING 更新 |
| `B_VERBOSE_MAIN_KW_LOG` | **未設定=false** | `[メインKW]` 詳細ログ |
| `B_INTEGRATED_SLICE_STARTED_AT_MS` | （自動） | 当該スライス開始時刻 |
| `B_DROPDOWN_RESUME_INDEX` | （自動） | プルダウン再開インデックス |

緊急停止: `B_WATCHDOG_ENABLED=false`。

**2026-08-14 ハード死対策中**: 番犬は **運用OFF必須**。B開始で20分1発を立てない。再開は Step1後の1分トリガーのみ。詳細は [B_HARD_DEATH_SCOPE_REQUIREMENTS.md](B_HARD_DEATH_SCOPE_REQUIREMENTS.md)。コード既定の未設定=true は残るが、B開始・トリガー再開では番犬を新規設置しない。

---

## 5. 番犬起動条件（定期）

すべて満たすときのみ:

1. `B_WATCHDOG_ENABLED` が true（未設定含む）
2. `B_INTEGRATED_RUN_STATE` がある
3. `LAST_PROGRESS.phase` が `runBIntegratedSteps:completed` でない
4. `LAST_PROGRESS.updatedAt` が `STALE_MIN` より古い
5. Lock が取れる（他の B 実行中でない）
6. `triggerRunCount` ≤ `B_INTEGRATED_MAX_TRIGGER_RUNS`

---

## 6. 検収

- [ ] Step5 途中ソフト打ち切り → 約1分後に `nextRow` から続き、AI列を全消ししない
- [ ] ハード死相当（番犬 stale）→ 人間無操作で続き、Step1しない、メールあり
- [ ] Step6 で同期 N>0、キャッチ／箇条書きが親に入る
- [ ] Step7 プルダウン途中打ち切り → カーソルから続き
- [ ] 完了後は番犬が再起動しない
- [ ] `B_WATCHDOG_ENABLED=false` で1発・定期とも動かない

---

## 7. 復元

- Property: `B_WATCHDOG_ENABLED=false` / `B_STEP5_ROW_CHECKPOINT=false`
- Git: 本変更のコミットを `git revert`

---

## 7.1 運用メモ（⑨⑩・時間制限に当たったとき）— 2026-08-13

社長方針: **改修を急がず、運用しながら検証**。時間制限／ハード死が出たら **まず本節＋§2〜§5 を参照**してからコード変更する。

| # | 事象 | 方針（2026-08-13） | 参照 |
|---|------|-------------------|------|
| **⑨** バリエーション提案欠け | Step7 未到達・途中切れで単位／内容量が空に見える | **いま改修不要**。多くは時間制限。番犬再開＋レ点を新規親に絞る運用で足りる想定 | 本ファイル §1・§2 #10（プルダウン TIME_SLICE） |
| **⑩** AI→マスタ同期ゼロ | Step6 `syncAiDataToMaster` が走っていない／空転記 | **忘れないメモ**。対策は番犬P0（実装済）。**運用検証待ち**。再発時は `LAST_PROGRESS` / `B_INTEGRATED_RUN_STATE` / Step5・6 ログを確認し、本要件の検収に沿って切り分け | §1（⑩背景）・§5・§6 |

### 時間制限ヒット時の手順（推奨）

1. Script Properties: `LAST_PROGRESS`・`B_INTEGRATED_RUN_STATE`（手編集しない）
2. 番犬が動いているか: 本対策中は **OFF**（`B_WATCHDOG_ENABLED=false`）。未設定=ONは旧既定。再開は1分トリガー
3. B→「はい」または自動続きで **Step1 からやり直さない**こと
4. それでも⑨⑩が残る場合だけ Agent に本節を渡して追加対策

---

## 8. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-14 | ハード死対策中は番犬OFF・B開始で非設置。正本 [B_HARD_DEATH_SCOPE_REQUIREMENTS.md](B_HARD_DEATH_SCOPE_REQUIREMENTS.md) |
| 2026-08-13 | §7.1 ⑨改修不要／⑩運用検証メモ（時間制限時の参照先） |
| 2026-08-12 | P0 要件確定＋実装（社長承認: 20分1発／定期両方／未設定ON／メール／docs新規） |
