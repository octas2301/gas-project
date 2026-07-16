# 楽天・Yahoo! CPO V2 要件定義（第一段階）

## 1. 目的

- **旧⑤**（Z の旧CPOサブメニュー「利益率レール版」）の **利益率レール**（`applyProfitMarginRailsOnPrice_` 等）による売価のねじれを避ける。
- **楽天・Yahoo! CPO V2** は **Amazon CPO V2**（`docs/PRICING_CPO_V2_REQUIREMENTS.md`）と **同じ後処理**（競合−1アンカー・補間・単調）で、**税込売値のたたき台**と **単価グラデーションの骨格** に特化する。
- **利益200円・送料確定・最終単調**は **②.5・③／再③** に任せる（V2 では強制しない）。

## 2. 旧⑤との関係

| 項目 | 旧⑤ | 4-V2（本機能） |
|------|-----|------------------|
| メニュー | Z の **旧CPO（legacy）** → 楽天Yahoo: 利益率レール版（旧⑤） | Z の **4. 楽天Yahoo!販売価格をAI提案(V2・単体)**（主フロー） |
| 利益率レール | あり | **なし** |
| 競合負けのセル色（薄黄） | あり | あり（Amazon V2 同様） |
| 利益不足の赤背景 | あり | **なし**（V2 では付けない） |
| 実装 | `menuProposePriceAndSetToSelection` | `menuRakutenYahooCpoProposePricesV2Standalone` → `runRakutenYahooCpoProposePricesV2_` |

- **旧⑤は残す**（並行運用・比較用）。

## 3. 処理内容

### 3.1 対象行

- **出品CK**（レ点）が付いた行のみ。
- **親行**（子SKUが空）＋**子行**（セット数あり）が揃った JAN 単位で処理（旧⑤と同じ）。

### 3.2 入力（プロンプト）

- **②.5 前（既定）**: テーブルは `セット数 / セット卸値(税込) / 競合価格(楽天またはYahoo!)`。**送料列は含めない**。
- **②.5 後オプション**: `includeShippingReference: true` のとき **確定送料を参考列**として追加（送料は決めない）。※単体メニューは現状②.5前のみ。必要なら Amazon V2 と同様の呼び出しを追加可能。
- 手数料・販促費は **参考**（プロンプト本文）。後処理では手数料で売価を変えない。

### 3.3 出力（JSON）

- 楽天: `{"rakuten":[{"setCount":1,"price":...}, ...]}`
- Yahoo!: `{"yahoo":[{"setCount":1,"price":...}, ...]}`
- パース: `parseCPOJsonForMall`、戦略テキスト: `stripCPOJsonFromResponse` → **親行**の `楽天価格戦略` / `Yahoo!価格戦略`。

### 3.4 後処理

- `applyCpoV2PricePostProcess_` を **Amazon V2 と共通**（競合列だけ `buildCpoV2GroupForRakutenYahooMall_` で差し替え）。
- **競合あり**: `round(競合) − 1`
- **補間・1%・単調**: Amazon V2 と同じ。
- 子行の **楽天価格設定** / **Yahoo!価格設定** に数値を書き込み。競合より高い場合は **薄黄**（Amazon V2 と同様）。

### 3.5 API 呼び出し

- JAN あたり **楽天 → 400ms 待機 → Yahoo!** → 次 JAN まで 500ms（旧⑤に近い間隔）。

### 3.6 競合なし（モール単位）→ Amazon 価格同期

- **当該モール**（楽天または Yahoo!）について、**レ点付き子行のすべてで競合価格が空または 0 以下**（正の競合が1件もない）とき、そのモールの処理では **Gemini を呼ばない**。
- 各子行の **`販売価格amazon`** が **すべて正の数値**なら、その値（整数に丸め、最低 1 円）を **楽天価格設定** または **Yahoo!価格設定** にそのまま書き込む。親行の戦略列には「競合なし（楽天／Yahoo!）: …Gemini 未呼出」と記録。
- **いずれかの子行で `販売価格amazon` が欠損**（空・数値でない・1 未満）の場合は **Logger に setNum を出し**、従来どおり **Gemini + `applyCpoV2PricePostProcess_`** にフォールバックする。
- 運用上 **3.2 の前に** Amazon CPO（Z の「3. …(V2・単体)」または B Step **3**）で **Amazon 列を埋める**前提。

## 4. 参照コード・プロンプト

- [CPO_PROMPT_V2_RY.md](CPO_PROMPT_V2_RY.md)（人間可読・コード内テンプレと同期）
- `getRakutenYahooCPOPromptTemplateV2` / `buildRakutenYahooCPOPromptForJANV2`
- `runRakutenYahooCpoProposePricesV2_` / `menuRakutenYahooCpoProposePricesV2Standalone`
- `buildCpoV2GroupForRakutenYahooMall_` / `applyCpoV2PricePostProcess_`
- JSON 欠損時は **Logger のみ**（`maybeWriteCpoAiNotes_` は呼ばない）。

## 5. 変更履歴

- 2026-03-19: 第一段階（単体メニュー・旧⑤維持）として追加。
- 2026-03-22: §3.6 競合なしモールは `販売価格amazon` 同期（欠損時は Gemini フォールバック）。
- 2026-03-22: Z メニュー整理（主フローは V2、旧⑤は「旧CPO（legacy）」サブメニュー）。
