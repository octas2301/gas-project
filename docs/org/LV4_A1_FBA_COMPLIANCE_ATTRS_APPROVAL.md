# レーンA1補 — FBAオファー必須属性（電池・危険物）承認パッケージ

**日付**: 2026-08-01  
**状態**: **実機合格**（2026-08-01。七味 FBA dry_run VALID → prod ACCEPTED）  
**親**: [LV4_LANE_A1_FBA_OFFER_APPROVAL.md](LV4_LANE_A1_FBA_OFFER_APPROVAL.md)／[D_MENU_LANE_A1_FBA_HUMAN_RUN.md](D_MENU_LANE_A1_FBA_HUMAN_RUN.md)  
**関連**: [LV4_DUAL_OFFER_MFN_FBA_APPROVAL.md](LV4_DUAL_OFFER_MFN_FBA_APPROVAL.md)（列分離は本包と独立）  
**三者レビュー**: **スキップ**（社長承認）

---

## 1. 目的（概要）

同一ASIN・同一日の実機対比により、**FBA（`AMAZON_JP`）の相乗りPUTだけ**が 90220（危険物規制・電池）で INVALID になることが確定した。  
FBA経路で **compliance 系属性を扱えるようにし**、かつ運用は次の**折衷**とする（初手から本番・dry_runスキップは採用しない）。

```text
（既定）dry_run → VALIDなら prod
（失敗時）issues＋おすすめ値を人間に提示
        → 人間が値を確認
        → 必ず dry_run で VALID を取る（コンプライアンス系はスキップ不可）
        → その後だけ prod
```

---

## 2. 実機対比（根拠・2026-08-01）

| | 自己発（MFN） | FBA |
|--|--------------|-----|
| ASIN | `B01N5A6ESU`（七味・既存カタログ） | 同左 |
| SKU | `sanky-B01N5A6ESU-19as13` | `sanky-B01N5A6ESU-19af13` |
| GET | **200**（既存） | **404**（新規） |
| dry_run | **VALID issues=0** | **INVALID 90220** |
| runId | `SPAPI_PUT_OFFER_CK_DRY_20260801_105918_199878` | `SPAPI_PUT_OFFER_CK_DRY_20260801_105527_f64afc` |

90220 文言:

- 「危険物規制」は必須ですが、入力されていません。  
- 「電池本体、電池が必要な商品」は必須ですが、入力されていません。  

豚汁・牛吸い・アンチョビの FBA でも同系統。自己発は通る → **FBAチャネル特有（または FBA新規SKU）の必須増**と結論。

---

## 3. 社長へ確認したい方針（承認時にチェック）— **折衷案**

### 3.0 運用フロー（本包の正）

| 段階 | 動作 |
|------|------|
| 既定の初手 | **必ず dry_run**（`VALIDATION_PREVIEW`）。初手 prod はしない |
| 成功 | VALID 後、従来どおり確認のうえ **prod**（ALLOW_PROD） |
| 失敗 | 本番に載せず、**issues 全文＋許可リスト内のおすすめ値**をダイアログ／ログで人間に示す |
| 人間判断 | おすすめを採る／捨てる／手で変える（Phase1は採否が主。手編集UIは最小で可） |
| 再検証 | コンプライアンス属性（電池・危険物等）を触ったら **必ず再度 dry_run**。VALIDになるまで prod 禁止 |
| prod | dry_run VALID 後のみ。おすすめ採用を理由に dry_run を飛ばさない |

**採用しない**

- 常に dry_run なしで初実行し、エラー時だけ人間通知  
- 「リスクなし」と判断したら dry_run なしで再本番  
- issues から任意属性を推測して無限自動リトライ  

### 3.1 論点表

| # | 論点 | 提案（既定＝折衷） | 承認 |
|---|------|-------------------|------|
| 1 | 初手 | **dry_run 必須**（prod 直は不可） | **済** |
| 2 | 失敗時UX | **何がエラーか＋おすすめ値**を人間に提示（許可コード／属性のみ。未知は止めて原文表示） | **済** |
| 3 | 再実行 | 人間確認後、**先に dry_run** → VALID のみ prod | **済** |
| 4 | 属性をいつbodyへ | **FBA時**。初回から安全既定を載せる **または** 90220検知後に載せて dry_run 再実行（どちらも可。実装は後者でも前者でも、**prod前に VALID 必須**） | **済** |
| 5 | 自己発 | body／フローとも **現状維持**（回帰回避） | **済** |
| 6 | 値の出所 | Phase1: 許可リストの安全既定（非電池・該当なし相当）。Phase2: マスタ列 | **済** |
| 7 | requirements | **`LISTING_OFFER_ONLY` 維持**（必要属性を同梱） | **済** |
| 8 | 試験 | 七味 FBA: dry_run（属性付き）VALID → 1SKU prod | **済** |
| 9 | 三点 | **スキップ**（契約変更時は再判定） | **済（スキップ）** |
| 10 | デュアル列 | **本包に含めない** | **済** |

---

## 4. 仕様案（実装承認後）

### 4.1 変更予定ファイル

| 種別 | パス | 内容 |
|------|------|------|
| 改修 | `AmazonSpapiPut.js` | FBA用 compliance 属性の付与。issues→おすすめ文言の整形。許可リスト外は自動補完しない |
| 改修 | `コード.js`（必要時） | 失敗ダイアログにおすすめ提示。再 dry_run／prod の導線（既存確認と整合） |
| 更新 | A1 HUMAN_RUN／本承認／PHASE／HANDOVER／LEDGER | 折衷フロー・再検収 |

**やらない**

- 初手 prod・コンプライアンスの dry_run スキップ  
- 未知 issues の自動補完・無限リトライ  
- マスタ列必須化（Phase2）／デュアル列／楽天・Yahoo・B統合改変  

### 4.2 属性候補（dry_run で確定）

| 目的（90220） | 候補キー | Phase1 おすすめ／既定（案） |
|---------------|----------|------------------------------|
| 電池が必要か | `batteries_required` | `false` |
| （必要なら）電池同梱 | `batteries_included` | `false` |
| 危険物規制 | `supplier_declared_dg_hz_regulation` | `not_applicable` 相当（Definitionsの enum に合わせる） |

- dry_run の issues／Definitions でキー・enum を確定  
- 許可リスト外の不足は **人間に原文のみ**（勝手に値を作らない）  

### 4.3 トグル（任意・推奨）

| Property | 既定 | 用途 |
|----------|------|------|
| `APPROVAL_AMAZON_SPAPI_PUT_FBA_COMPLIANCE_ATTRS` | **true** | FBA時の属性付与／おすすめ提示。false で旧挙動 |

### 4.4 検収

- [x] 方針承認（§3 折衷）… **2026-08-01**  
- [x] 実装承認＋コード… **2026-08-01**  
- [x] 七味 FBA dry_run VALID… `SPAPI_PUT_OFFER_CK_DRY_20260801_111613_41ce9e`  
- [x] 七味 FBA prod ACCEPTED… `SPAPI_PUT_OFFER_CK_PROD_20260801_111845_6bd20f`  
- [x] 自己発 dry_run 回帰（対比時 OK）  
- [x] dry_run なし prod をコンプライアンス経路の正としない（折衷維持）  
- [ ] （任意）失敗時おすすめダイアログの目視  

---

## 5. 想定リスクと緩和

| リスク | 緩和 |
|--------|------|
| 初手本番で誤属性 | **折衷で禁止**。dry_run 必須 |
| おすすめ誤採用 | 人間確認＋再 dry_run。prod は VALID 後のみ |
| 属性名・enum不一致 | dry_run／Definitions。許可リストのみ。**実機で not_applicable 系が通った** |
| 自己発回帰 | FBA分岐のみ |
| EC書込 | ALLOW_PROD＋確認 |

---

## 6. 復元

1. `APPROVAL_AMAZON_SPAPI_PUT_FBA_COMPLIANCE_ATTRS=false` で属性付与OFF  
2. PUT 系トグル false  
3. 当該 commit／差分 revert  

---

## 7. 次ゲート

1. ~~方針・実装・実機~~ **済**（A1合格）  
2. （別）デュアル Phase1／A2／A3  
3. （後）Phase2: マスタ列参照  

---

## 8. 社長確認

- [x] §3 折衷方針承認  
- [x] 三点スキップ  
- [x] 実装承認  
- [x] 実機検収… **2026-08-01**  

---

## 9. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-01 | ドラフト〜§3折衷・方針・実装。 |
| 2026-08-01 | **実機合格**: dry_run `…111613_41ce9e`／prod `…111845_6bd20f`。 |
