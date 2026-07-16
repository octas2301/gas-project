# プロジェクト全体の位置づけと現在の開発フォーカス

**最終更新**: 2026-07-17  
**読み方**: 次の Agent は `docs/AGENT_HANDOVER.md` の **§1.5・§2** に従い、**本ファイルを最初に読み**、続けて §2 の必読一覧でプロジェクト全体をインプットする。

---

## 1. プロジェクト全体（30秒）

| 項目 | 内容 |
|------|------|
| **目的** | 楽天・Yahoo・Amazon 向けに、リサーチ〜商品情報・価格・セット構成〜出品に必要なデータを **スプレッドシート＋GAS** で整備・自動化する。一人社長＋AI部門モデルは [org/AI_ORG_CHARTER.md](org/AI_ORG_CHARTER.md)。 |
| **クリティカルパス** | [FLOW_AND_PRIORITY.md](FLOW_AND_PRIORITY.md) — リサーチ・見積もり → 出品情報 → 各モール出品。 |
| **実装の中心** | `コード.js`、主シートは `▼商品マスタ(人間作業用)`、`AI情報取得data`、`00_設定マスタ`。 |
| **6領域・成果物** | [PROJECT_OVERVIEW.md](PROJECT_OVERVIEW.md)。 |

---

## 2. 現在のフェーズ（いま優先している開発）

- **フォーカス領域**: **AI組織 Phase0** — 憲章・承認マトリクスの正本化と **3者多数決の反映**まで。手順の正は [org/THREE_REVIEW_RUNBOOK.md](org/THREE_REVIEW_RUNBOOK.md)（親1＋並列サブ3）。記録は [org/PHASE0_THREE_REVIEW_MAJORITY.md](org/PHASE0_THREE_REVIEW_MAJORITY.md)。**実装コードは書かない。**
- **このフェーズの完了条件（目安）**  
  - 憲章・承認マトリクスが docs にあり、多数決採用項を反映済み。  
  - 社長が多数決メモを最終承認し、未決一覧が洗い出されている。  
  - 次フェーズ（Lv別プラン／在庫0日中出品の実装）への着手条件が文書上はっきりしている。

- **並行・継続（後回し可）**: 商品情報まわりの OpenAI 429 切り分け、Z「11-③」と B Step7 の差の検証（[AI_ROUTING_GEMINI_OPENAI.md](AI_ROUTING_GEMINI_OPENAI.md)）。楽天ジャンル Nav は Stage0 完了・Stage1 実装済み（専用診断シートのみ・Property 既定オフ。[RAKUTEN_NAV_GENRE_STAGE1.md](RAKUTEN_NAV_GENRE_STAGE1.md)）。

- **スコープ外（本 Phase0）**  
  - `Yahoo.js` / `コード.js` の改修、GAS Web 承認UI実装、三モール出品ジョブ、`clasp push` 自動化。  
  - マスタの SKU/JAN 一括変更、CPO・価格ロジック変更（別チケット）。

---

## 3. 直近で確定した仕様・前提（docs 反映済み）

- **AI組織**: 朝モバイル承認（**当面社長のみ**）→ 日中無人出品（原則12:00前）→ 夜確認。在庫0原則（不可なら1・承認明示）。販売中(在庫>0)は**原則スキップ**。補充は別承認で `FLOOR(仕入÷セット)`。冪等・部分成功時は失敗側のみリトライ。CS／問屋は送信禁止・下書き可。詳細は org 憲章・マトリクス・多数決メモ。  
- **キーワード案・商品名案の本線は OpenAI**。Gemini 版の商品名案は **コードに存在するが本流メニュー未配線**。  
- **バリエーション単位**は **Gemini → OpenAI 補完**（◎ASIN 経路も同様）。**内容量パース**は Gemini 中心（ChatGPT 自動フォールバックなし）。  
- **11-③（`runProductNameProposalsForRows`）** は、**OpenAI 失敗時もマスタ商品名があればバリエーションに届く**（既定）。旧挙動は `PRODUCT_NAME_PROPOSALS_CONTINUE_VARIATION_ON_OPENAI_FAIL=false`。**B Step7** は **variation まで進み得る**（商品名が空でも単位・内容量のみ更新されうる）。  
- 詳細は [AI_ROUTING_GEMINI_OPENAI.md](AI_ROUTING_GEMINI_OPENAI.md)、[商品マスタ_人間作業エリアとマスタエリア_要件定義.md](商品マスタ_人間作業エリアとマスタエリア_要件定義.md)。

---

## 4. 次にやること（優先順）

1. 社長が [org/PHASE0_THREE_REVIEW_MAJORITY.md](org/PHASE0_THREE_REVIEW_MAJORITY.md) を最終承認する。  
2. **Lv別プラン**（在庫0日中出品・Amazonバルク等）を策定し実装フェーズへ。  
3. （並行可）OpenAI 429 / `insufficient_quota` の解消と、11-③ と B Step7 のログ比較。  
4. （並行可）楽天ジャンル Nav Stage1 の **人間検証**（`clasp push` → Property true → 17-⑥）。次は Stage2（別 scriptId）または組織 Lv プラン。

---

## 5. 深掘りリンク（領域別）

| テーマ | ドキュメント |
|--------|----------------|
| **AI組織・承認** | [org/AI_ORG_CHARTER.md](org/AI_ORG_CHARTER.md)、[org/AI_APPROVAL_MATRIX.md](org/AI_APPROVAL_MATRIX.md)、[org/THREE_REVIEW_RUNBOOK.md](org/THREE_REVIEW_RUNBOOK.md)、[org/PHASE0_THREE_REVIEW_MAJORITY.md](org/PHASE0_THREE_REVIEW_MAJORITY.md) |
| 楽天ジャンル Nav（並行） | [RAKUTEN_NAV_GENRE_DIAG.md](RAKUTEN_NAV_GENRE_DIAG.md)（Stage0）、[RAKUTEN_NAV_GENRE_STAGE1.md](RAKUTEN_NAV_GENRE_STAGE1.md)（Stage1要件） |
| Gemini / OpenAI・11-③ vs B Step7・429 | [AI_ROUTING_GEMINI_OPENAI.md](AI_ROUTING_GEMINI_OPENAI.md) |
| 商品マスタ人間作業エリア | [商品マスタ_人間作業エリアとマスタエリア_要件定義.md](商品マスタ_人間作業エリアとマスタエリア_要件定義.md) |
| 価格・送料・再③・B 統合 | [PRICING_V1_REQUIREMENTS.md](PRICING_V1_REQUIREMENTS.md)、AGENT_HANDOVER **§8** |
| リサーチ・セット構成 | [RESEARCH_AND_ESTIMATE.md](RESEARCH_AND_ESTIMATE.md) |
| エージェント共通・必読一覧 | [AGENT_HANDOVER.md](AGENT_HANDOVER.md) **§2** |

---

## 6. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-07-17 | 並行: 楽天ジャンル Nav Stage1 実装（専用診断シート追記・Property 既定オフ）。 |
| 2026-07-15 | **AI組織 Phase0**: 3者多数決反映・RUNBOOK（親1＋並列3）・多数決メモ追加。実装は次フェーズ。 |
| 2026-07-15 | **AI組織 Phase0**: org 憲章・承認マトリクスをフォーカスに。実装は次フェーズ。 |
| 2026-03-19 | 初版。全体＋現在フェーズの引き継ぎを本ファイルに集約。 |
| 2026-03-22 | OpenAI 失敗時の 11-③ バリエーション継続（既定）を §3 に反映。 |
