# 未広告ASIN掲載 — 人間手順（第1スライス）

**最終更新**: 2026-10-25  
**正本**: [広告運用_未広告ASIN掲載_第1スライス要件.md](広告運用_未広告ASIN掲載_第1スライス要件.md)  
**コード**: `広告運用GAS/UnadvertisedAsin.js`／メニューは `コード.js` の「🚀 広告運用」  
---

## 0. 初回だけ（人間）

### 0.1 clasp 反映

```powershell
cd C:\Users\takuy\Desktop\gas-project\広告運用GAS
clasp push
```

スプレッドシートを開き直してメニュー更新を確認。

### 0.2 Script Properties（広告運用GAS）

出品GASと同じ SP-API 読取用を、**広告運用プロジェクト**にも入れる。

| キー | 内容 |
|------|------|
| `SPAPI_LWA_CLIENT_ID` | LWA Client ID |
| `SPAPI_LWA_CLIENT_SECRET` | LWA Client Secret |
| `SPAPI_REFRESH_TOKEN` | Refresh Token |
| `SPAPI_SELLER_ID` | 出品者ID |
| `SPAPI_MARKETPLACE_ID` | 未設定なら JP `A1VC38T7YXB528` |
| `SPAPI_ENDPOINT` | 未設定なら `https://sellingpartnerapi-fe.amazon.com` |

任意:

| キー | 既定 | 意味 |
|------|------|------|
| `UNADVERTISED_ASIN_ENABLED` | 未設定=ON | `false` で機能停止 |
| `UNADVERTISED_ASIN_DAILY_BUDGET` | 1000 | 日予算（円）。シート優先可 |
| `UNADVERTISED_ASIN_DEFAULT_BID` | 11 | 基本入札（円） |
| `UNADVERTISED_ASIN_MAX_LISTINGS` | 200 | 1回の一覧上限 |
| `UNADVERTISED_ASIN_CSV_MAX_AGE_DAYS` | 14 | CSV古い警告の日数（一括DLも同日数でログ警告） |
| `UNADVERTISED_ASIN_ID_AUTO_FROM_BULK` | 未設定=ON | `false` で `ID参照用` の自動更新を停止 |
| `LISTING_MASTER_SS_ID` | 出品SS既定 | 商品名ベース／バリ／メーカー名（親くくりには使わない） |
| `UNADVERTISED_ASIN_OFFER_COUNT_ENABLED` | 未設定=ON | `false` でセラー数APIスキップ |
| `UNADVERTISED_ASIN_OFFER_SLEEP_MS` | 3000 | バッチ間隔 |
| `UNADVERTISED_ASIN_OFFER_429_WAIT_MS` | 90000 | 429待機 |
| `UNADVERTISED_ASIN_OFFER_429_RETRIES` | 5 | 429再試行 |

### 0.3 Driveフォルダ

Amazon用フォルダ（`AMZ_FOLDER_ID`）の中に次を置く（無ければGASが一覧／バルク時に作成可）。

| フォルダ | 中身 |
|----------|------|
| `📥広告中の商品レポート` | 広告中の商品CSV（最新1本） |
| `📥一括キャンペーンDL` | 一括操作のカスタムシート xlsx/CSV（最新1本）→ `ID参照用` 自動更新の種 |

---

## 1. 毎回の運用

1. Amazon Ads → レポート → **広告中の商品** をCSVダウンロード → `📥広告中の商品レポート`  
2. **（既存に足すとき）** Amazon Ads → 左 **一括操作** → **カスタマイズしたスプレッドシートを作成しダウンロード**（スポンサープロダクト。インプレッション0除外はOFF推奨）→ `📥一括キャンペーンDL`  
3. **🚀 広告運用** → **3. 未広告ASIN：一覧を更新する**（このとき `ID参照用` も一括DLから自動更新）  
4. `未広告ASIN一覧` で載せたい**子**だけレ点ON  
   - 並び: 新しい親ブロックが上  
   - 予定名: 商品が分かる長さ（短名強制なし）。`VGXT` のようなASIN末尾は付けない  
   - **相乗り**のキャンペーン名は出品タイトル（マスタ商品名ベースではない）。備考に `相乗り名=出品タイトル`  
   - **ブランド**列でノーブランド（`（なし）`）を確認  
   - 備考「相乗り」「名前衝突」は既定レ点OFF  
   - 相乗り: 新品セラー数 ≥ 2  
   - 更新のたびに直前内容が `未広告ASIN一覧_前回` に残る（合否の前後比較用）  
5. **3-② レ点からバルクを作る** → はい（直前にもう一度 `ID参照用` を自動更新）  
6. `📦生成されたバルクファイル` を一括操作でUL  
7. KW最適化はレーンB（従来）  

手順の見本（画像貼付枠つき）: メニュー **3-③ 未広告ASIN：掲載マニュアルを作る** → シート `📌未広告ASIN掲載マニュアル`（文言更新後は3-③で作り直し）

---

## 2. 一覧の見方

| 列 | 意味 |
|----|------|
| レ点 | 載せる子 |
| 親ASIN / 掲載方針 | 親ブロック。`新規ペア` / `既存に追加` |
| キャンペーン予定名 | 商品名ベース（なければ商品名）。長くてよい |
| ブランド | SP-APIブランド。空はメーカー名→`（なし）` |
| 広告グループ予定名 | 子ごと |
| 広告状態 | `あり` / `なし` |
| 備考 | 名前衝突・親フォールバック・相乗りなど |

親ASIN: **Listings → Catalog のみ**（マスタ親SKUは使わない）。  
同じ親の子は名前を共有する。「かぶり」は **別親どうしの名前文字列** の話（正本 §0）。

---

## 2.1 合否判定（前後比較を必ず入れる）

一覧更新のあと、エージェント／人間は次を見る。

**見る場所**
- `未広告ASIN一覧`（今回）
- `未広告ASIN一覧_前回`（直前）
- 実行ログ（親解決の Listings／Catalog／子自身 件数）

**今回の合否（絶対）**
- 焼鳥丼など: キャンペーン名に `VGXT`／親ASIN末尾が **付いていない**
- **ブランド**列がある。空は `（なし）` と分かる
- 備考に `親=マスタ親SKU` が **出ない**（Listings／Catalog／フォールバックのみ）

**前回との比較（必須）**
| 観点 | 前回（旧） | 今回（期待） |
|------|------------|--------------|
| 名前末尾 | `VGXT` / `33XK` / `4WHY` 等 | 無し（可読な長い名前） |
| ブランド列 | 無し | 有り |
| 親解決メモ | `親=マスタ親SKU` があり得る | Listings／Catalog／子自身のみ |
| 同じ親の塊 | 維持または改善 | 悪化していない（バラけ増えていない） |
| 名前衝突件数 | （参考） | 増減を報告。増えてもASIN末尾で隠さない |

合否回答の先頭に「見た場所」と「前回比（良くなった／同じ／悪くなった）」を1行ずつ書く。

---

## 3. 切り戻し

- `UNADVERTISED_ASIN_ENABLED=false`  
- またはメニューを使わない  
- Git: `UnadvertisedAsin.js` を revert  

---

## 4. 既知の注意

- バルクxlsxは公式テンプレ SPシート**31列**＋実DLの値。v1.7で AG仮ID一意化・既存は数値ID必須・除外重複排除。  
- Ads API申請とは別レーン。  
- 同じメニューを件数分割で何度も押さない。上限は `UNADVERTISED_ASIN_MAX_LISTINGS`。  

---

## 4.1 原因切り分けログ（3-②実行後）

Apps Script 実行ログで次を見る（`runId=` で絞る）。

| ログ断片 | 見ること |
|----------|----------|
| `idRefresh=` / `idCamps=` | 一括DLからID参照が更新されたか、キャンペーン件数 |
| `id_probe` | `うず巻き=` `唐辛子=` 等が 0 なら **DL／ID参照にその名が無い**（名寄せ以前） |
| `bulk_resolve` の `path=` | `create_by_name`＝名前で新規作成（already exists→連鎖の型）／`existing_with_id`・`promote_*`＝数値IDで既存足し／`skip_existing_no_id`＝既存方針なのにID無し |
| `autoHit=` `manHit=` | `miss` なら照合失敗。`autoNear=` に別表記候補が出れば名ゆれ |
| `autoId=(name)` | 数値IDが付かず名前をID代わりにしている |

渦巻・七味の再発時は、同じ `runId` の `id_probe` と当該親の `bulk_resolve` 1行を渡せば原因を特定できる。  

