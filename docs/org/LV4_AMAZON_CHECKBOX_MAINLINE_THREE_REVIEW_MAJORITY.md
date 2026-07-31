# Amazonレ点本線＋Amazon相乗りSKU — 3者独立レビュー多数決

**日付**: 2026-07-30  
**mode**: independent-triple  
**状態**: **社長承認・正本反映済**  
**対象**: [LV4_AMAZON_CHECKBOX_MAINLINE_SELLER_SKU_APPROVAL.md](LV4_AMAZON_CHECKBOX_MAINLINE_SELLER_SKU_APPROVAL.md)  
**手順**: [THREE_REVIEW_RUNBOOK.md](THREE_REVIEW_RUNBOOK.md)  

---

## レビューア一覧（モデル・結論）

- Reviewer-1: Claude Opus 5 Thinking High — **条件付き**
- Reviewer-2: Cursor Grok 4.5 High Fast — **条件付き**
- Reviewer-3: GPT-5.6 Terra Medium — **条件付き**

子同士は結果を共有せず、コード・docsを書き換えていない。

## スコア表（3者）

| レビュー | 実現性 | 要件漏れ | 矛盾 | 聖域 | 逃げ漏れ |
|----------|-------:|---------:|-----:|-----:|---------:|
| Reviewer-1 | 3 | 2 | 2 | 5 | 3 |
| Reviewer-2 | 4 | 3 | 3 | 5 | 4 |
| Reviewer-3 | 4 | 3 | 3 | 5 | 3 |

## 採用（2/3以上）

1. 現行Dラジオを、本件のレ点＋複数チェック式UIで置換すると明記する
2. 憲章・承認マトリクスに「Amazon D手動本線に限り、人間の子SKUレ点を当面の承認①相当」と反映する
3. `出品CK` は boolean `true`／文字列 `"TRUE"` 両対応。親レ点のみは除外する
4. prodは主トグル＋`ALLOW_PROD`＋件数・SKU例・FORCE_QTY_0確認ダイアログを必須とする
5. `Amazon相乗りSKU` は既存の「出品者SKU＝子SKU」の **A/M2相乗り専用例外**とする
6. 列追加後、`[セット構成提案][列範囲チェック]` と親SKU／子SKU数式を実機確認する
7. 出荷方式ヘッダ・許容値、`target_val` 元ヘッダ・優先順の確定前はコード着手禁止
8. 旧D承認との置換・正本の優先関係を明記する

## 未決（1票または割れ）

- レ点再実行時のDONE／冪等管理の実装詳細
- 新規＋相乗り同時実行時の時間切れ・再開方式
- `Amazon相乗りSKU` の物理位置 → **NF列で確定**
- 現行A入口を先に実機試験するか、後続UIへ直接進むか

販売中・在庫>0スキップは1票の指摘だったが、既存の憲章・マトリクス・Lv4正本にある固定制約のため、例外承認なしで維持する。

## 社長回答（2026-07-30）

1. **多数決採用事項の正本反映を承認**
2. **フル＋既存相乗りprodを許可**
   - フル開始前にprodゲートと確認ダイアログを完了
   - 取消時は楽天／Yahooを含むフル全体を開始しない
3. sellerSkuの扱い
   - 新ASIN型／旧JAN型の各sellerSkuについて、Amazon上に同一sellerSkuがあれば更新
   - 未登録sellerSkuは新規登録
   - 新ASIN型と旧JAN型は別sellerSkuであり、相互上書きではない

## 総合

**条件付きYES**

正本反映は完了。コード実装は次を満たした後、別途社長承認を得て開始する。

1. 出荷方式列の正式ヘッダ名・許容値 → **`自己発送`／FBA等で確定**
2. 子SKU式からN列参照を削除 → **人間作業済み**
3. 実装承認欄 → **2026-07-30承認済み**
