# Yahoo.js コードレビューレポート

| 項目 | 内容 |
|------|------|
| 対象 | `Yahoo.js`（約 2,654 行 / Ver 3.x） |
| 実施日 | 2026-07-11 |
| レビュー基準 | `AGENTS.md` §5.2〜§5.3、`HANDOVER.md` §2〜§5.3 |
| 制約 | **Yahoo.js 本体は未変更**（指摘・提案のみ）。EC API / スプレッドシート / clasp 操作なし |
| レビュー範囲 | 静的解析（本番実行・実機検証は未実施） |

---

## 総評

出品コア（子SKUレ点のみ・`seller_id` の画像=クエリ / 商品=ボディ・`submitItem` の it-07004 成功扱い・`grouping_id` サニタイズ・`normalizeYahooItemPathForApi_`）は仕様どおり実装されている。一方で、**マスタヘッダー行のハードコード不一致**、**完了メールの成功/失敗混在**、**6分制限に対するオートレジューム未実装**、**taxable / 文字数上限の送信前ガード不足**が実運用リスクとして目立つ。チェックボックスは出品CK経路では概ね両対応だが、ヘルパーと削除・CSV経路で判定が揃っていない。

---

## 指摘一覧

### 1. マスタヘッダー行の検出方法が経路ごとに不一致

| 項目 | 内容 |
|------|------|
| **重要度** | 高 |
| **該当箇所** | `YahooDataBuilder._loadMasterData`（871–885 行付近）／`updateMasterYahooId`（2061–2079）／`updateMasterDeleteFlag`（2345–2365 付近）／`showDeleteSelectionDialog`（41–42, 62）／`listDeletableItems`（265–278） |
| **なぜ問題か** | Builder は先頭20行から「親SKU」「出品CK」を探してヘッダーを動的検出する一方、出品済ID更新・削除CK・削除フラグ更新は **`data[7]`（8行目固定）**。ヘッダー位置がずれると、出品は成功してもマスタ未更新、削除対象の取りこぼし／誤更新が起きる。 |
| **推奨する対処** | ヘッダー検出を共通化し、全経路で同じ `headerRowIdx` を使う。 |

```javascript
function findMasterHeaderRowIndex_(values) {
  for (var r = 0; r < Math.min(20, values.length); r++) {
    if (values[r].includes('親SKU') && values[r].includes('出品CK')) return r;
  }
  return -1;
}

function updateMasterYahooId(masterSheet, childSku, yahooItemCode) {
  var data = masterSheet.getDataRange().getValues();
  var headerRowIdx = findMasterHeaderRowIndex_(data);
  if (headerRowIdx < 0) { /* warn & return */ }
  var headers = data[headerRowIdx];
  // ... データ行は headerRowIdx + 1 から
}
```

---

### 2. 完了メールの「出品済み一覧」に失敗商品も含まれる

| 項目 | 内容 |
|------|------|
| **重要度** | 高 |
| **該当箇所** | `sendYahooCompletionEmail`（2121–2133 行付近）、呼び出し元 `runYahooExport`（440–441） |
| **なぜ問題か** | 引数 `products` は出品**試行**全件。失敗分も「出品済み商品一覧」と削除用URLに載る。運用者が失敗品を成功と誤認し、誤削除リンクを踏むリスクがある。失敗は別セクションにあるが、一覧側のフィルタがない。 |
| **推奨する対処** | 成功した `item_code` の Set を渡すか、メール用に成功分だけを渡す。 |

```javascript
// runYahooExport 側で成功コードを蓄積
var successCodes = [];
// runOneProduct 成功時: successCodes.push(product.code);
sendYahooCompletionEmail(products, sellerId, successCount, errorCount, errorItems, successCodes);

// メール内
products.filter(function (p) { return successCodeSet.has(p.code); }).forEach(...);
```

---

### 3. 大量出品時の GAS 6分制限・オートレジューム未実装

| 項目 | 内容 |
|------|------|
| **重要度** | 高 |
| **該当箇所** | `runYahooExport`（303–450 行付近）。画像ループ＋商品ごと `editItem`→`setStock`→`submitItem`＋ `Utilities.sleep(200/1000/500)` |
| **なぜ問題か** | `HANDOVER.md` でも Yahoo 側にオートレジュームが必要と明記されているが未実装。件数増で実行時間切れ→途中まで出品済・マスタ一部更新・メール未送信などの**中途半端な状態**になりやすい。 |
| **推奨する対処** | 開始時刻を記録し、残り時間が閾値（例: 5分30秒）を切ったら Script Properties に進捗（次インデックス・runId）を保存してトリガー再開。楽天聖域・B統合境界は触らず、Yahoo 単独／一括の Yahoo 部分のみ。実装は EC 書き込みを含むため**事前承認必須**。 |

---

### 4. `taxable` を文字列 `"1"`/`"0"` で送信している

| 項目 | 内容 |
|------|------|
| **重要度** | 高 |
| **該当箇所** | `YahooApiClient.updateItem`（1080–1082）、デバッグ群（1543, 1621, 1829 など） |
| **なぜ問題か** | `AGENTS.md` / `HANDOVER.md` は **数値 `1/0`（文字列 `"1"` 不可）** と明記。`_coerceTaxable` は数値を返すが、直後に `String(...)` している。form-urlencoded では最終的に文字列化されるが、ペイロード組み立て時点の型規約違反であり、過去の taxable 系障害の再発余地がある。 |
| **推奨する対処** | |

```javascript
} else if (field === 'taxable') {
  value = this._coerceTaxable(value); // number 1 or 0 のまま
}
```

デバッグの型チェック（`val === 0 \|\| val === 1`）とも揃える。

---

### 5. `caption` / `explanation` / `headline` の文字数上限ガードがない

| 項目 | 内容 |
|------|------|
| **重要度** | 高 |
| **該当箇所** | `YahooDataBuilder._createSingleProduct`（name のみ 75 文字処理: 751–763）／`YahooApiClient.updateItem`（送信前に caption/explanation/headline を切り詰めていない） |
| **なぜ問題か** | 仕様: name 全角75・caption 5000（HTML可）・explanation 500（HTML不可）・headline 30。name 以外は超過時に API エラーで出品失敗しやすい。`[[JOIN_BULLETS]]` 等で explanation に `<br>` が入ると HTML 不可制約にも抵触しうる。 |
| **推奨する対処** | 送信直前にフィールド別サニタイズ。 |

```javascript
function clipYahooText_(s, maxChars) {
  s = String(s == null ? '' : s);
  return s.length > maxChars ? s.substring(0, maxChars) : s;
}
// name: 75 / headline: 30 / explanation: 500（HTML除去後）/ caption: 5000
```

explanation 向けマッピングでは `[[HTML_CLEAN]]` または送信前 `_stripHtml` を必須化する運用も推奨。

---

### 6. `updateMasterYahooId` / 削除CK解除が行単位 `setValue`（I/O 過多）

| 項目 | 内容 |
|------|------|
| **重要度** | 中 |
| **該当箇所** | `updateMasterYahooId`（2074–2077）※商品ごと全件スキャン＋1セル書き込み／`showDeleteSelectionDialog`（92–96）※対象行ごと `setValue(false)` |
| **なぜ問題か** | 成功 N 件でマスタを最大 N 回フルスキャン＋N 回書き込み。削除CK解除も同様。スプレッドシート I/O がボトルネックになり、指摘3のタイムアウトを悪化させる。 |
| **推奨する対処** | 子SKU→行番号の Map を1回構築。書き込みは列レンジをまとめて `setValues`。 |

---

### 7. チェックボックス判定ヘルパーの不統一（`1` 未対応・CSVは trim なし）

| 項目 | 内容 |
|------|------|
| **重要度** | 中 |
| **該当箇所** | `yahooMasterCheckboxIsTrue_`（701–703）／削除CK（64）／`YahooCSVGenerator._loadMapping`（558）／`YahooApiClient._loadMapping`（1028） |
| **なぜ問題か** | 出品CKは `true` / `"TRUE"` 対応で仕様を満たす。一方 (a) 数値 `1` はヘルパー未対応（削除CKは `=== 1` あり）、(b) CSV の出力フラグは `=== "TRUE"` のみで trim/`toUpperCase` なし、(c) API マッピングは trim+UPPER あり。経路によってレ点が無視される。 |
| **推奨する対処** | |

```javascript
function yahooMasterCheckboxIsTrue_(cell) {
  if (cell === true || cell === 1) return true;
  return String(cell == null ? '' : cell).trim().toUpperCase() === 'TRUE';
}
// 削除CK・出力フラグもすべてこのヘルパーに寄せる
```

---

### 8. `logYahooEvent` / 完了メールが `getActiveSpreadsheet()` 固定

| 項目 | 内容 |
|------|------|
| **重要度** | 中 |
| **該当箇所** | `logYahooEvent`（969–974）、`sendYahooCompletionEmail`（2088）、`runYahooExport(ssOverride)`（303–305） |
| **なぜ問題か** | トリガー等で `ssOverride` を渡しても、ログ・メール内 URL 取得が Active 依存。Active が無い／別ブックのときログ欠落・メール失敗・誤ブック参照が起きうる。 |
| **推奨する対処** | `ss` を引数で受け渡し、`runYahooExport` の `ss` をログ／メールまで貫通させる。 |

---

### 9. `grouping_id` の全角→半角が Builder と `updateItem` で非対称

| 項目 | 内容 |
|------|------|
| **重要度** | 中 |
| **該当箇所** | `_createSingleProduct`（813–826）／`updateItem` の grouping-id 分岐（1086–1089） |
| **なぜ問題か** | Builder は全角英数字を半角化してからサニタイズ。`updateItem` はマッピング値を `replace(/[^a-zA-Z0-9-]/g, '')` のみ。マッピング経由の全角は**削除**され、同一親でも ID が空／不一致になりバリエーション分断の原因になる。 |
| **推奨する対処** | 全角正規化＋サニタイズを共通関数化し、Builder / updateItem / CSV の `grouping-id` で共用。 |

---

### 10. Web 削除エンドポイントに認証がない

| 項目 | 内容 |
|------|------|
| **重要度** | 中 |
| **該当箇所** | `doGet` / `doPost`（2220–2287）、`deleteYahooItem`（2294–） |
| **なぜ問題か** | メールの削除用 URL を知っていれば（または推測できれば）第三者でも削除 API を叩ける設計。運用上の利便性とのトレードオフだが、EC 書き込みの入口としてはリスクが高い。`doPost` の Yahoo 削除は `action` 必須チェックも弱い。 |
| **推奨する対処** | 署名付きトークン（HMAC + 有効期限）を URL に付与して検証。または Script Properties の共有シークレット。楽天削除 Web 経路も同様に見直し。 |

---

### 11. `ship_weight` が `"0"` のときも送信されうる

| 項目 | 内容 |
|------|------|
| **重要度** | 中 |
| **該当箇所** | `updateItem` の空判定（1092–1094）、`_post` / `_preparePayloadForPost` の空削除（1166–1168, 1377–1381） |
| **なぜ問題か** | 仕様は「空ならパラメータ自体を送らない」。`""`/`null` は削除されるが、`"0"` や `0` は残る。重量未設定を 0 と書いたマスタがあると API エラーや意図しない重量設定になりうる。 |
| **推奨する対処** | |

```javascript
if (apiParam === 'ship_weight') {
  var n = Number(value);
  if (value === '' || value == null || !isFinite(n) || n <= 0) continue; // 送らない
}
```

---

### 12. 画像アップロード失敗が商品単位で握りつぶされ、ステータス未反映

| 項目 | 内容 |
|------|------|
| **重要度** | 中 |
| **該当箇所** | `runYahooExport` 画像ループ（348–360）、`YahooImageUploader.upload`（467–530） |
| **なぜ問題か** | 画像 NG でも editItem 以降へ進む。マスタの「Yahoo画像」ステータス更新もこの経路では見当たらない。画像なし／一部欠けのまま出品完了扱いになる。 |
| **推奨する対処** | 商品ごとに画像成功数を記録し、0 枚ならスキップまたは警告付きで続行する方針を明示。可能ならマスタ「Yahoo画像」列へ `済`/`Error` をバッチ書き込み（EC・マスタ更新のため承認前提）。 |

---

### 13. HTML 生成でのエスケープ不足（XSS／表示崩れ）

| 項目 | 内容 |
|------|------|
| **重要度** | 中（ダイアログ）／低〜中（Web） |
| **該当箇所** | `createDeleteSelectionHtml`（150–151: `item.name` 生埋め込み）、`getDeleteConfirmHtml` 等の商品名表示 |
| **なぜ問題か** | 商品名に `<` や引用符が含まれると HTML が壊れ、最悪スクリプト実行。GAS HtmlService でもユーザ入力の生埋め込みは避けるべき。 |
| **推奨する対処** | `HtmlService` のテンプレートまたは手動エスケープ（`& < > "`）。 |

---

### 14. デバッグ関数が本番 `updateItem` とペイロード組み立てが乖離

| 項目 | 内容 |
|------|------|
| **重要度** | 中（保守）／一部 低 |
| **該当箇所** | `debugPayloadConstruction` / `debugYahooApiRequest` / `debugVariationAndGrouping`（1444–1898）、`debugPriceAndImages`（2017: `new YahooApiClient(mapSheet)`） |
| **なぜ問題か** | デバッグは固定キーで payload を手組みしており、本番の「マッピング＋outputFlag＋fileType=data」動的組み立てと一致しない。調査結果が本番とずれる。`debugPriceAndImages` はコンストラクタ引数順 `(token, sellerId, mapSheet)` を誤っており、即時に壊れている。 |
| **推奨する対処** | 調査は `client.updateItem` 前に `_preparePayloadForPost` 相当を本番と同じ関数から取得するようリファクタ。壊れているデバッグは修正かメニュー非表示。 |

---

### 15. `suffixName` が長いと name が 75 超／空寄せになる

| 項目 | 内容 |
|------|------|
| **重要度** | 低〜中 |
| **該当箇所** | `_createSingleProduct`（749–763） |
| **なぜ問題か** | `maxLen = 75 - suffixName.length`。バリエーション値が極端に長いと `maxLen <= 0` になり、結果が suffix のみで 75 超、または実質空に近い name になる。 |
| **推奨する対処** | suffix 側も上限を分け、最終 `yahooName` を必ず 75 で clip。 |

---

### 16. 兄弟画像の収集対象が「レ点付き子」のみ

| 項目 | 内容 |
|------|------|
| **重要度** | 低（仕様確認） |
| **該当箇所** | `_groupMasterData`（921–924: siblings に isCheck のみ）、`_createSingleProduct`（778–786） |
| **なぜ問題か** | 未レ点の兄弟メイン画像は載らない。意図的ならドキュメント化で十分。意図が「同一親の全バリエーション画像」なら不足。 |
| **推奨する対処** | 仕様を `HANDOVER.md` に明記。全兄弟が必要なら siblings と children を分離（出品対象 vs 画像参照用）。 |

---

### 17. 削除・出品の API 連打に共通のレート制御が弱い

| 項目 | 内容 |
|------|------|
| **重要度** | 低〜中 |
| **該当箇所** | `executeDeleteFromDialog`（206–214: sleep なし）、`submitItem` コメントの 1クエリ/秒、`runOneProduct` の sleep |
| **なぜ問題か** | 出品側は概ね待機あり。まとめて削除は連続 `deleteItem` で制限に当たりやすい。 |
| **推奨する対処** | 削除も 1 秒間隔など共通 `yahooThrottle_()` を挟む。 |

---

### 18. 巨大ファイル＋調査用デバッグの同居（可読性）

| 項目 | 内容 |
|------|------|
| **重要度** | 低 |
| **該当箇所** | ファイル全体。特に 1440 行以降の debug*、HTML 生成関数 |
| **なぜ問題か** | 本番経路（export / builder / api client / delete）と調査・UI HTML が同一ファイルで、差分レビューと clasp 同期の認知負荷が高い。 |
| **推奨する対処** | 将来的に `YahooDebug.js` / `YahooDeleteUi.js` へ分割を検討（新規ファイルは承認対象）。当面はセクション見出しと「本番エントリ一覧」をファイル先頭コメントに列挙。 |

---

### 19. 出品対象判定（子SKUレ点のみ）は仕様どおり（良好）

| 項目 | 内容 |
|------|------|
| **重要度** | （問題なし・確認結果） |
| **該当箇所** | `_groupMasterData`（900–924） |
| **なぜ問題か** | — |
| **推奨する対処** | 親レ点はマージ元親行の選択にのみ使い、子の `isCheck` でのみ `children` に追加。`HANDOVER.md` §5.2.1 の禁止仕様（親だけレ点で全子出品）にはなっていない。**維持すること。** |

---

### 20. `seller_id` 配置・`submitItem` の it-07004 扱いは仕様どおり（良好）

| 項目 | 内容 |
|------|------|
| **重要度** | （問題なし・確認結果） |
| **該当箇所** | `YahooImageUploader.upload`（502–503）／`_post` / `_postAndReturnBody`（1373–1374, 1409）／`runOneProduct`（388–394） |
| **なぜ問題か** | — |
| **推奨する対処** | 画像=クエリ、商品・在庫・反映・削除=ボディ。新規の it-07004 は成功扱い＋手動反映案内。**維持すること。** |

---

## 特に優先して直すべき TOP5

| 優先 | 指摘 | 理由 |
|------|------|------|
| 1 | **#1 ヘッダー行ハードコード不一致** | 出品成功とマスタ更新・削除の前提が壊れ、データ不整合の直撃になる |
| 2 | **#3 オートレジューム未実装** | 件数増で必ず表面化する運用障害。途中状態の復旧手段がない |
| 3 | **#2 完了メールの成功/失敗混在** | 誤操作（誤削除）と運用判断ミスに直結 |
| 4 | **#4 taxable の型** / **#5 文字数・HTML ガード** | API 仕様違反による出品失敗の再発防止（まとめて送信前バリデーション層に） |
| 5 | **#6 マスタ書き込みのバッチ化**（＋#7 チェックボックス統一） | タイムアウト緩和とレ点取りこぼし防止。比較的安全に効く |

---

## レビュー観点チェック（AGENTS.md 対応表）

| 観点 | 結果 |
|------|------|
| バグ・境界・getRange | ヘッダー固定（#1）、メール一覧（#2）、name suffix（#15）を指摘。getRange の「行数・列数」取り違えは当該箇所では目立たず |
| Yahoo API（seller_id / パラメータ名 / taxable / grouping / 文字数 / submitItem） | seller_id・submitItem は良好。taxable・文字数・grouping 非対称・ship_weight 0 を指摘 |
| チェックボックス true / "TRUE" | 出品CKは概ね OK。ヘルパー・CSV・削除で不統一（#7） |
| パフォーマンス | オートレジュームなし、行単位 setValue、全件スキャン繰り返し（#3, #6） |
| 保守性 | デバッグ乖離・巨大単一ファイル（#14, #18） |

---

## 修正時の注意（実装しない／承認前提）

本レポートの修正案をコードに落とす場合:

- **Yahoo.js 変更は EC 出品・削除・マスタ更新に関わるため、規模に関わらず事前承認**（`AGENTS.md` §3.2 / ユーザー承認ルール）。
- 楽天ロジック（`コード.js` の `generateRakutenCSV` 等）には触れない。
- 必須3点セット（docs 更新・調査ログ・復元手段/トグル）を遵守。
- Cloud 上では `clasp push` しない。実機確認はローカルで人間が実施。

---

## 実機での確認手順（人間向け・参考）

レポート自体は静的レビューのため未実施。修正後の確認例:

1. ローカルで対象差分のみ `clasp push`
2. 出品CK付き少数件で `runYahooExport` → `▼ログ(システム用)` で editItem / setStock / submitItem
3. ヘッダー行をずらしたコピーシートで `updateMasterYahooId` 相当が追従するか
4. 意図的失敗を混ぜ、完了メールの一覧が成功のみか
5. taxable / 長文 caption・explanation / 空 ship_weight の送信ログ確認

---

*本ファイルはレビュー成果物です。Yahoo.js への適用コミットは含みません。*
