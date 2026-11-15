# D: R-Cabinet フォルダ自動切替（上限時）

**状態**: 要件ロック（2026-10-21）／**コード済**（要出品 clasp・Property ON）  
**対象**: 出品 `コード.js` の D 画像アップロード（R-Cabinet）  
**関連**: [BATCH_EXPORT_IMAGE_GATE_REQUIREMENTS.md](../BATCH_EXPORT_IMAGE_GATE_REQUIREMENTS.md)／一括出品ログ診断（2026-10-21）

---

## 0. 確定事項（社長 2026-10-21）

| 項目 | 確定 |
|------|------|
| 新フォルダの置き場 | **ルート直下**（基本フォルダ ID 0 の直下には作らない API 制約に従う。上位は未指定＝ルート） |
| 上限ヒット後 | **同じ D 実行の続き**で新フォルダへ切替し、失敗した画像から再送する |
| 停止時 | **原因追及できるログ**を一括出品ログへ必ず残す（§5） |

---

## 1. 目的

フォルダ内画像数上限（`Number of files is upper limit`）で D の画像アップロードが止まらないようにする。  
達した／達しそうなとき、機械が新フォルダを作り、そこへ上げ、設定の上げ先を更新する。

---

## 2. 現状

| 項目 | 内容 |
|------|------|
| 上げ先 | `▼設定(マッピング)` の **B9**（フォルダID）。実運用例: `12644827` |
| 失敗例 | `Number of files is upper limit` |
| 既存コード | `cabinet.file.insert`／`files/search` のみ。フォルダ作成なし |

---

## 3. API

| API | 用途 |
|-----|------|
| `cabinet.usage.get` | 契約上限・空き（容量／フォルダ数／FileMax 等） |
| `cabinet.folder.files.get` 等 | 現フォルダ枚数の確認（取れる範囲） |
| `cabinet.folder.insert` | 新フォルダ作成 |
| 既存 `cabinet.file.insert` | 新 FolderId へアップロード |

---

## 4. いつ切り替えるか

### 4.1 予防（達しそう）

上げ開始時、または親 SKU 開始時に利用状況を見る。

- 現フォルダの残り枚数が、この実行の見込み枚数より少ない  
- または残り枚数がしきい値以下（既定 **50**。Property で変更可）

### 4.2 事後（達した）

`file.insert` が `Number of files is upper limit`（または同等）を返した。

### 4.3 作らず止める（店舗全体）

次のときは **新フォルダを作らない**。ログに理由を残し、人へ知らせる（トースト／PENDING_ALERT 可）。

- 店舗全体の容量不足（`AvailSpace` 不足）
- フォルダ数上限（`AvailFolderCount=0`）

---

## 5. 切替後の動き（同じ D の続き）

1. 新フォルダを作成する  
   - 名前: `auto_YYYYMMDD_HHMM`（Asia/Tokyo）  
   - 置き場: **ルート直下**（§0）  
2. 返ってきた FolderId を、**この実行の上げ先**にする  
3. **B9 を新 FolderId で上書き**する（以後の D もここを使う）  
4. 上限で失敗したスロットから、**同じ D の続き**で再送する  
5. 1 回の D 実行あたりの新規フォルダは **最大 1 つ**（超えそう／2 つ目が必要なら止めて人へ）

---

## 6. 停止時・切替時のログ（原因追及・必須）

届け口は既存の **一括出品ログ**（秘密・鍵は出さない）。

### 6.1 切替が起きたとき

1 行に要約できる形で残す（複数行可）。含める項目:

| 項目 | 例 |
|------|-----|
| 理由 | `予防` / `上限エラー` |
| 旧フォルダID | `12644827` |
| 新フォルダID | （作成結果） |
| フォルダ名 | `auto_20261021_2055` |
| usage 要約 | FileMax／使用数／AvailFolderCount／AvailSpace（数値のみ） |
| 直前の親SKU・sheet行 | どこまで進んでいたか |
| 直前の失敗 msg 先頭 | 180 字まで（鍵なし） |

### 6.2 作らず止めたとき

| 項目 | 例 |
|------|-----|
| 停止理由コード | `STORE_SPACE_FULL` / `FOLDER_COUNT_MAX` / `AUTO_FOLDER_CAP` 等 |
| usage 要約 | 同上 |
| 当時の上げ先フォルダID | B9 または実行中 ID |
| 親SKU・sheet行 | 停止時点 |
| 人がやること 1 行 | 例: 「RMSで容量を空ける／不要画像を減らす」 |

### 6.3 実行の生死（既存診断と併用）

既にある D 診断と矛盾させない。

- `Dトリガー開始`／`Dトリガー終了`  
- `画像UP 親開始`／`親終了`／心拍  
- **`Dトリガー終了` が無い**＝時間切れ等の強制終了の手がかり（従来どおり）

切替後も心拍・親終了・ループ完了を同じ規則で出す。

### 6.4 失敗分類の拡張

既存 `batchExportClassifyCabinetErr_` 系に、少なくとも次を足す（または同等の class 文字列）:

- `CABINET_FILE_LIMIT` … `Number of files is upper limit`  
- （既存）`URLFETCH_UNAVAILABLE` / `SIZE_CAPACITY` / `AUTH` / `TIME` / `OTHER`

### 6.5 UrlFetch 一時失敗の限定再試行（2026-10-23）

方針（理想分岐）:

| 状態 | 再試行 |
|------|--------|
| fileId 無し（insert 失敗） | **insert のみ**（`policy=insert_only`） |
| fileId あり（search 失敗／FileUrl欠落） | **search のみ**（`policy=search_only`）。insert しない |
| 外側 | Search 失敗で fileId が残っていれば `search_only_resume` をもう一巡。容量系だけ縮小して insert し直す |

`callRCabinetInsertFile` = insert_only →（成功時）search_only。Yahoo出品 `_post` と同型の一時判定:

| 項目 | 内容 |
|------|------|
| 対象 | 通信例外（使用できないアドレス／timeout 等）／HTTP 429・5xx／search の FileUrl欠落 |
| 回数 | 各フェーズ最大3（待機 1秒→2秒）。恒常4xxは再試行しない |
| ログ | **一括出品ログ**へ `transient`／`recovered`／`final_fail`／`host_probe`／`search_only_resume`。付帯: folderId・newUpStreak・sec・parent・slot・row・col・name |
| 返却 | `fileId` フィールド＋メッセージ末尾の `searchAttempts=` / `insertAttempts=` / `lastMs=` / `fileId=` |
| 切戻し | Script Property `RAKUTEN_CABINET_URLFETCH_RETRY_ENABLED=false`（未設定=ON） |

Google 側が「使用できないアドレス」の内側理由を返さないため、**再試行で通るか**と **host_probe の http／throw** で一時通信かどうかを切り分ける。  
`recovered` がシートに出れば「一時通信だった」と特定できる。`final_fail`＋`host_probe http=` は出口復旧、`host_probe throw=` は出口側も不通の手がかり。

---

## 7. 人間の役割

- 検収: 一括出品ログの切替行と B9 の値  
- 店舗全体満杯／フォルダ数上限のときの整理  
- clasp・実装承認  
- 型・文案の発明は求めない  

---

## 8. 禁止・聖域

- 楽天 CSV 生成本体・MAKE 連携の仕様変更はしない（上げ先フォルダが変わるだけ）  
- 未使用画像の自動削除はしない（別要件）  
- 鍵・トークンをログに出さない  
- 1 実行でフォルダを無制限に増やさない（§5）  

---

## 9. 切戻し

| 手段 | 内容 |
|------|------|
| Property | `RAKUTEN_CABINET_AUTO_FOLDER_ENABLED`（**未設定=OFF**＝従来どおり B9 固定） |
| 手 | B9 を旧フォルダIDに戻す |

しきい値案: `RAKUTEN_CABINET_AUTO_FOLDER_REMAIN_THRESHOLD`（未設定=50）。

---

## 10. 検収

- [ ] フォルダ内上限エラーでも、空きがあれば新フォルダ作成→B9更新→同じ D で続きの画像が上がる  
- [ ] 店舗全体満杯／フォルダ数上限では作成せず、§6.2 のログを残して停止  
- [ ] 一括出品ログに旧→新フォルダID・理由・usage 要約がある  
- [ ] Property OFF で従来動作  
- [ ] 楽天 CSV 聖域を壊さない  

---

## 11. 実装スライス（承認後・コード改修は別承認）

1. usage 取得＋`CABINET_FILE_LIMIT` 判定＋§6 ログ  
2. `folder.insert`（ルート直下）＋ B9 更新  
3. 同一 D 内の上げ先切替と失敗スロット再送（最大フォルダ 1）  
4. Property 切戻し＋ HUMAN_RUN 追記  

`clasp push` は人間。切戻しは Property `RAKUTEN_CABINET_AUTO_FOLDER_ENABLED=false`（未設定も OFF）。

### 人間手順（ON）

1. 出品で clasp push  
2. Script Properties: `RAKUTEN_CABINET_AUTO_FOLDER_ENABLED=true`  
3. D をフル・画像オンで再実行  
4. 一括出品ログで `キャビネット切替`／`停止理由=` を確認  

---

## 改修履歴

| 日付 | 内容 |
|------|------|
| 2026-10-21 | コード投入。usage／folder.insert／B9更新／同じD再送／診断ログ。未設定=OFF |
| 2026-10-21 | 初版ロック。ルート直下・同じ D の続き・停止時原因ログ。未実装 |
