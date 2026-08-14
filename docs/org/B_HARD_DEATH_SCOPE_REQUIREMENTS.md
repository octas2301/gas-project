# B統合 — ハード死対策（件数・切断・当B集合）

**文書種別**: 要件定義＋実装  
**日付**: 2026-08-14  
**状態**: 実装済（ローカル。要 `clasp push`／`B_WATCHDOG_ENABLED=false`）  
**親**: [CURRENT_PHASE.md](../CURRENT_PHASE.md)／[AGENT_HANDOVER.md](../AGENT_HANDOVER.md) §8  
**コード**: `コード.js`  
**ゴール一文**: AIストックは残し、未処理を上から最大3行パック（シリーズ温存・4行以上は3+余り）して夜中に連鎖。Step1後に新スライス。当B挿入商品だけ処理。JANでは紐づけない。

---

## 1. 背景

2026-08-14 の全件B: Step1 後に同一実行で Step2（画像Gemini）へ入り約30分ハード死。番犬は検知したが Step2 先頭 `SpreadsheetApp.getUi()` で即死し、失敗を無限リトライした。

---

## 2. 対策 ID（これ以外は今やらない）

| ID | 内容 |
|----|------|
| G | B開始ゲート: AI対象行の商品名と卸値(税込)必須。欠けたら実行しない |
| A | 1回最大3行。人間の今回CKなし。`B済`空の行を上から自動パック |
| B | 親あたりセット数最大10（和集合・`[1,2,3]` の後） |
| C | Step1完了後は必ず終了 → 約1分後トリガーで Step2 を新スライス |
| S | 当B集合: `商品名ベース+卸値(税込)` 記録。照合は親SKU優先。子は親SKU。JANは照合に使わない |
| S2 | B統合 Step2〜8 および Step5/6 はこの集合だけ。Z単独は従来 |
| E | Step2 の `getUi` をトリガー安全に |
| F | 番犬OFF。B開始で20分1発を立てない |

やらない: 画像LIMIT自動、ガードOFF、書いてからシリーズ揃え、ハード死でStep1やり直し、Step2内部JAN切れ目。

シリーズ揃え（セット数）は現行どおり **挿入前**（`alignSeriesSetCountsInPlans_`）。  
パックは **シリーズIDの連続グループを割らない**（次が3行なら今が2でも止める）。4行以上は3チャンク＋余りを次回で再計算。

---

## 3. 整合（破綻防止）

1. 人手でZを切るな、と C は矛盾しない。トリガー再開（`startStepIndex>=1`）では再切断しない。  
2. `saveBIntegratedState` は既存JSONをマージし `insertedProducts` を消さない。完了時だけ削除。  
3. `insertedProducts` 空で Step2 以降に入ったら no-op（全レ点へ戻さない）。  
4. G/A は `startStepIndex===0`（メニュー・連鎖トリガーとも）。再開（>=1）では再パックしない。  
4b. `queuedAiRows` を状態に保存し Step1 はその行だけ（Step1途中死で再パックして別商品を拾わない）。  
4c. 全Step完了後、未処理があれば状態を新規runId・Step0で1分後連鎖。失敗時は連鎖しない。  
5. 上限10は `ensureMinThreeVariationSetCounts123_` の後。1,2,3は残す。  
6. B統合はレ点で対象を決めない。行番号キーは使わない。  
7. API検索は親行のJANフィールドを使う。紐づけキーは名前+原価／親SKU。  
8. Step5/6 も同じキーで AI／マスタを絞る。  
9. C の再開は `setBIntegratedTrigger` のみ。E とセット。  
10. 非 TIME_SLICE 失敗時も全レ点へ戻さない。

---

## 4. データ

`B_INTEGRATED_RUN_STATE.insertedProducts[]`: `nameBase`, `costTaxIn`（小数2桁）, `parentSku`, `jan`（監査）。

正規化: 名前 trim・連続空白畳み。税込は Number。空・0 は G で開始不可。

実行中フラグ: `B_INTEGRATED_STEPS_ACTIVE_`（`runBIntegratedSteps` 内のみ true）。

トグル:

| Key | 既定 |
|-----|------|
| `B_SCOPE_INSERTED_PRODUCTS_ONLY` | 未設定=ON |
| `B_STEP1_FORCE_SLICE_AFTER` | 未設定=ON |
| `B_WATCHDOG_ENABLED` | **運用は false**（本対策中） |

---

## 5. 入口ゲート（G+A）とキュー

AI情報取得data **1列目 `B済`**（無ければ挿入。Step5は列名引き）。

未処理 = 中身ありかつ B済が空。シート上から:

1. 同じシリーズIDの連続行をグループ。ID空は1行1グループ。  
2. グループが4行以上 → 先頭から3行ずつチャンク。余りは次回。  
3. チャンクを先頭から足す。次を足すと3超なら **1行も足さず** 今のバッチで終了（次の3まとまりを温存）。

ゲートは **パックした1〜3行だけ** 名前／卸値(税込)。未処理0はキュー空で正常終了。

`B済` は Step1 で親を書いた直後に GAS が日時・runId・parentSku を書く。Zセット構成単独は全中身あり行のまま・B済は書かない。

---

## 6. 検収

- AIストック10行でも1回は最大3親。B済の次から連鎖  
- 前2＋次シリーズ3 → 今は2で止める  
- 同一シリーズ4行 → 3＋余り1  
- パック行の名前または税込空 → 始まらない  
- 未処理0 → キュー空で終了  
- Step1後に切断、約1分後 Step2（getUi例外なし）  
- 同名同原価の既存親は親SKUで当B親だけ  
- セット数11以上 → 10で書込  
- Z単独横断は全レ点  
- 番犬トリガーが増えない  

復元: Property トグルOFF／Git revert。`B_WATCHDOG_ENABLED=false` は残してよい。
