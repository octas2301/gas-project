# 競合DB — メーカー実態AI＋ヒット数え直し＋Catalog増員

**日付**: 2026-09-19（2026-10-31 改訂P2コード）  
**状態**: 実態AI＋ヒット時計＋**Catalog増員（P2-1/P2-2・未設定=OFF）**  
**親**: [B_MAKER_JITTA_CLASSIFICATION.md](../docs/org/B_MAKER_JITTA_CLASSIFICATION.md)／[CYCLE §4.5bis／§5](../docs/org/B_PURCHASE_RESEARCH_CYCLE_REQUIREMENTS.md)／[競合JITTA要件 P2](../docs/org/B_COMPETITOR_JITTA_REQUIREMENTS.md)

---

## 0. 目的

競合DB側で **未調査メーカーの実態AI** と **セラーヒット／小規模出現の数え直し** を回し、①の Keepa／Catalog と実行枠を分ける。  
**Catalog（改訂P2）**: 係交代ではなく **増員並走**。①をCで黙らせない。Keepaは①のみ。

---

## 1. 範囲

| やる | やらない |
|------|----------|
| 未調査の実態AI（Gemini→OpenAI） | Keepa／門／自動レ点／C0／A-light |
| サマリ実態／ミラー・マスタキャッシュ書込 | サマリ行の物理削除・統合 |
| 法人格ゆれの同一キーミラー（heal） | |
| セラーヒット全店再走査（足し算禁止） | |
| 小規模出現（①候補の門通過を読むだけ） | |
| **Catalog列挙（規模確定・洗いキュー追記）** | 分割語フル列挙（①洗い側） |
| 時間トリガー（実態AI／ヒット／**Catalog**は別時計） | |

### Catalog（P2）

- 対象: メーカーサマリで製造メーカー／入口ブランドかつ Catalog見積が空（規模不明優先）
- 書く先: ①調査スプシ `①洗いキュー`（メモ=`comp_catalog`）／サマリの規模帯・Catalog見積・洗い進捗
- 親ロック: 競合スプシ `Catalog親ロック`（owner=`comp`）。①は `hb`
- 第1スライス: 本クエリで最大501件まで。超えたら特大見積で打ち切り
- 洗いキュー: **1実行1回読み**＋追記はメモリ継ぎ足し（①の SNAPSHOT と同思想。社ごと全読み禁止）
- 未済≥5万で新規追記停止。**追記0では規模帯を書かない**（`catalog_pending_cap`）。幽霊 `via=comp`＋キュー0は `catalog_phantom_heal`
- 未済≥上限で候補が全員キュー0なら **API前に batch_skip**（同じ社を試し続けない。時計は間隔維持）
- Prop `PR_WASH_QUEUE_PENDING_ROWS_MAX`（①と同名。未設定=50000。`0`でオフ）
- 失敗: `catalog_skip` に理由／`comp_enum` は残さない。バッチ先頭で stuck heal
- 落ちあと: **起動先頭で次予約**＋失敗時再予約＋実態AI／ヒット時計からの再武装（`catalog_rearm`）。`updated` 古ロック掃除。人手連打なし

---

## 2. 配置

| 項目 | 内容 |
|------|------|
| リポジトリ | `competitor_jitta/`（`JittaResolve.js`＋`CatalogEnum.js`） |
| 紐づけ | 競合スプシ `1Urd…` に **コンテナバインド**した GAS（scriptId を `.clasp.json` へ） |
| ①との二重防止 | ① `PR_JITTA_OWNER=comp`／当たり直しは① `PR_HIT_RECOUNT_OWNER` 省略=`comp` |
| ①洗いキュー | 調査スプシ `1tf7…`（`CJ_RESEARCH_SS_`） |

---

## 3. Property（競合側）

| Prop | 既定 | 意味 |
|------|------|------|
| `CJ_ENABLED` | ON | `false` で実態AI全停止 |
| `CJ_CAP` | 8（最大20） | 1回の件数 |
| `CJ_INTERVAL_MIN` | 10（1/5/10/15/30 に丸め） | 実態AIトリガー間隔 |
| `CJ_DAILY_CAP` | 80 | 1日の resolve 上限（通常。U≤ゲート） |
| `CJ_DAILY_CAP_BOOST` | 1000 | **U≥ゲート+1** のとき日次上限 |
| `CJ_BACKLOG_GATE` | 30 | 未調査ユニーク社数 U の閾値 |
| `CJ_HIT_ENABLED` | ON | `false` でヒット時計停止 |
| `CJ_HIT_INTERVAL_MIN` | 実態AIと同じ | ヒット時計間隔 |
| `CJ_CATALOG_ENABLED` | **OFF（未設定=OFF）** | `true` でCatalog増員ON |
| `CJ_CATALOG_CAP` | 3（最大8） | 1回の列挙社数 |
| `CJ_CATALOG_INTERVAL_MIN` | 10 | Catalog時計間隔 |
| `CJ_CATALOG_REARM_MIN` | 5 | 落ちあと再予約分 |
| `CJ_CATALOG_LOCK_STALE_MIN` | 35 | 親ロック `updated` 古さで掃除 |
| `CJ_CATALOG_STALE_GRACE_MIN` | 10 | 発火予定超過の幽霊判定猶予 |
| `SPAPI_LWA_CLIENT_ID` / `SECRET` / `SPAPI_REFRESH_TOKEN` | （Catalogに必須） | ①と同じ鍵を競合にも置く |
| `SPAPI_ENDPOINT` / `SPAPI_MARKETPLACE_ID` | 省略可 | FE／JP既定 |
| `GEMINI_API_KEY` | （必須・実態AI） | 実態AI |
| `OPENAI_API_KEY` | 任意 | Gemini失敗時 |

## 3.1 Property（①側）

| Prop | 既定 | 意味 |
|------|------|------|
| `PR_JITTA_OWNER` | `hb` | `comp`=競合に委譲 |
| `PR_HIT_RECOUNT_OWNER` | `comp` | 省略も委譲。切戻し `hb` |
| `PR_SHARED_CATALOG_LOCK` | ON | `false`＝共有シートロックOFF（①単独） |

---

## 4. 人間手順

```powershell
cd C:\Users\takuy\Desktop\gas-project\competitor_jitta
clasp push

cd C:\Users\takuy\Desktop\gas-project\purchase_research
clasp push
```

1. 競合 Script Properties に SP-API 3点を入れる（秘密はチャットに書かない）  
2. `CJ_CATALOG_ENABLED=true`  
3. メニュー **実態AI（競合）→ 8. Catalog時計 開始**（1回）  
4. 合否: `実態AIログ` に `catalog_batch`。競合スプシに `Catalog親ロック`。①ログに①の列挙が止まっていない  

**① clasp と混ぜない。** Agent は clasp しない。

---

## 5. 合格ゲート

| ID | 合格 |
|----|------|
| G1〜G8 | 実態AI／ヒット（従来どおり） |
| P2-G1 | Cが厚くても①Catalogが止まらない（競合OFFでも） |
| P2-G2 | 別親並走・同一親は同時にならない（`Catalog親ロック`） |
| P2-G3 | Keepaは①のみ |
| P2-G4 | 競合落下後も①継続。競合は自動再開 |
| P2-G5 | `CJ_CATALOG_ENABLED=false` で①単独 |

---

## 6. 切戻し

1. Catalog: **9. Catalog時計 停止** または `CJ_CATALOG_ENABLED=false`  
2. ヒット: **ヒット時計 停止** または `CJ_HIT_ENABLED=false`／① `PR_HIT_RECOUNT_OWNER=hb`  
3. 実態AI: トリガー停止＋`PR_JITTA_OWNER=hb`  
4. 共有ロックだけ切る: ① `PR_SHARED_CATALOG_LOCK=false`  

---

## 7. 改訂

| 日付 | 内容 |
|------|------|
| 2026-11-11 | Catalog: 未済上限時 API前 batch_skip（空転止め） |
| 2026-11-10 | Catalog: 幽霊規模帯禁止＋pending_cap シートログ＋phantom heal |
| 2026-11-01 | Catalog: キュー1回読み＋失敗札クリア＋heal（①で踏んだ全読み穴の移植） |
| 2026-10-31 | **P2-1/P2-2コード**: `CatalogEnum.js`・共有 `Catalog親ロック`・①側共有ロック。未設定=Catalog OFF |
| 2026-10-31 | **改訂P2要件**: Catalog増員並走・①沈黙禁止 |
| 2026-09-19 | ヒット数え直し（別時計） |
| 2026-09-05 | 初版 |
