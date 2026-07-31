# U4 — Drive02→R2→マスタURL（人間手順）

**状態**: **実機合格**（2026-07-26）  
**承認**: 2026-07-26「U4 v1 承認」  
**正本**: [D_MENU_U4_R2_URL_EMBED_REQUIREMENTS.md](D_MENU_U4_R2_URL_EMBED_REQUIREMENTS.md)

---

## 0. 実機合格記録（2026-07-26）

| 項目 | 結果 |
|------|------|
| clasp push | 済（人間） |
| 21-⑦ | OK — `runId=U4_20260726_090920_1366af`・MAIN成功=1・失敗=0・マスタ更新SKU=1・列追加=`Amazon MAIN URL`,`Amazon PT URL` |
| マスタ | 対象子行に `Amazon MAIN URL`（`https://pub-….r2.dev/…`）あり → **本丸PASS** |
| 21-①／D Amazon | `idempotentBlocked=1`・親0＝既存バッチ冪等で想定内（新規GENERATEDなし）。URL確認はマスタで代替可 |
| Property | `AMAZON_U4_URL_EMBED_ENABLED=false` へ戻す（ダイアログ指示どおり） |

---

## 1. Script Properties

| Key | 値 |
|-----|-----|
| `AMAZON_U4_URL_EMBED_ENABLED` | `true`（実行時のみ → 終わったら **false**） |
| `AMAZON_U4_MAX_SKUS` | 省略可（既定 20） |
| `AMAZON_U4_SKU_LIST` | 任意。カンマ区切り子SKU（最優先） |
| `AMAZON_DRIVE_IMAGE_FOLDER_ID` | 空なら Drive `02` 既定 |
| `R2_*` | T2と同じ（ACCOUNT / ACCESS / SECRET。BUCKET・PUBLIC_BASE 省略可） |

※ 21-⑦は **U4トグルのみ**で R2 Put（`AMAZON_DRIVE_R2_UPLOAD_ENABLED` は 21-⑥用）。

---

## 2. clasp push

```text
clasp push
```

含む: `AmazonDriveImageExport.js`／`AmazonApprovalExport.js`／`コード.js`（21-⑦）

---

## 3. 操作順

1. U2 で対象子の MAIN を `02` に出す（④）  
2. Property: `AMAZON_U4_URL_EMBED_ENABLED=true`  
3. （任意）`AMAZON_U4_SKU_LIST=lifec-….80s10`。空ならマッチングの Amazon MAIN 付き子 → なければマスタ `Amazon MAIN 参照` 付き子  
4. **Z → 21 → 21-⑦** → 完了ダイアログで MAIN成功・マスタ更新を確認  
5. マスタ列 `Amazon MAIN URL`（必要なら `Amazon PT URL`）← **検収の正**  
6. （任意）21-①／D で GENERATED。冪等ブロック時はマスタURLで合否可。GENERATED列まで見るなら 21-②後に再実行  
7. `AMAZON_U4_URL_EMBED_ENABLED=false`

---

## 3b. D×Amazon 新規からの自動実行（2026-07-31 承認）

- **D×Amazon 新規**は GENERATED の直前に **U4 を自動実行**する（`batchExportAmazonAutoU4_`）。手動の 21-⑦ と `AMAZON_U4_URL_EMBED_ENABLED` は不要。
- 対象は **レ点行のうち URL が未充足の子SKU のみ**。`Amazon MAIN URL` があり、かつ（`Amazon PT 参照` が無い or `Amazon PT URL` がある）行は **スキップ**（ログ `autoU4 skip reason=urls_present`）。
- **失敗時はDを停止**する（画像なしバルクを作らないため）。原因は 21-⑦ 単独実行で確認する。
- 止めたいときは Property **`AMAZON_U4_AUTO_IN_D_ENABLED=false`**（既定 true）。
- **サブ画像**:
  1. `Amazon PT 参照`（Drive ID）があればそれを R2 へ上げる
  2. 無ければ（かつ `AMAZON_ONLY` でなければ）マスタ **`楽天サブ画像1〜8`**（子→親フォールバック）を取得して R2 へ上げ、`Amazon PT URL`（`|` 区切り）に書く
  3. **親行へも空欄時のみコピー**（C1は親行も読む）。PT が0件のときは既存値を消さない
  4. このため **マッチングsheetで楽天サブとAmazon PTを二重ドラッグする必要はない**（楽天側のサブ紐付けだけで足りる）。Amazon MAIN（白抜き）の紐付けは従来どおり必要
- MAIN は Drive02 `{SKU}.MAIN.jpg` を優先。無い場合でも既存 `Amazon MAIN URL` があればそれを維持して PT のみ処理する

### 帰宅後・既出品へのサブ追加（手ZIP）

1. Property: `AMAZON_U4_URL_EMBED_ENABLED=false` に戻す  
2. マスタ `Amazon PT URL`（例: PT01|PT02|PT03）と MAIN URL／Drive02 から JPG を用意  
3. 名前: `{子SKU}.MAIN.jpg` / `{子SKU}.PT01.jpg` …（フラット）  
4. ZIP化 → SC **Upload Images**  
5. 任意: Drive `04.SC用画像ZIP保存先` へコピー  

七味例: `sanky-B01N5A6ESU-19s13`・PT成功=3（`U4_20260731_124440_329c1a`）

---

## 4. 戻し方

- Property false（手動21-⑦は `AMAZON_U4_URL_EMBED_ENABLED`、D自動は `AMAZON_U4_AUTO_IN_D_ENABLED`）  
- `git revert`／メニュー 21-⑦ 削除  
