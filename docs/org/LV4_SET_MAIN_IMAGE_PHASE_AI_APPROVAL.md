# セットMAIN画像 — AI生成方針（ロック済／PoC実装承認済）

**状態**: **方針ロック＋API PoC実装承認**（2026-08-02 社長）。  
**前段**: [LV4_SET_MAIN_IMAGE_PHASE_A_APPROVAL.md](LV4_SET_MAIN_IMAGE_PHASE_A_APPROVAL.md)（Pillow・本線凍結）  
**手順**: [D_MENU_SET_MAIN_IMAGE_PHASE_AI_HUMAN_RUN.md](D_MENU_SET_MAIN_IMAGE_PHASE_AI_HUMAN_RUN.md)  
**三者**: PoC品質判定後に本線一括の要否を判断（本包では省略可）

---

## 1. 背景（Pillow Phase A の結論）

| モール | 評価 |
|--------|------|
| Amazon | **運用不可**（白枠貼付・極小グリッド） |
| 楽天 | 数字差し替え方向は正しいが Canva 級に未達 |

社長判断: **見本参照のAI生成へ振り切る。** 手作業: Google AI Studio → **NANOBANANA** で実現可能に見える。

---

## 2. ロックした目的（2026-08-02 補正）

| モール | やること | 禁止 |
|--------|----------|------|
| **Amazon** | 見本（03）の **配置・大きさ・重なり** をレイアウト設計図として転写。商品は01のみ | 見本を無視した自由構図、商品ラベルの改変 |
| **楽天／Yahoo** | **ベース画素は一切変更しない**。金丸内へ数量（＋マスタ単位）を **レイヤー重ね** のみ | AIによるバナー全面再生成・デザイン改変 |

楽天の評価ポイント（金丸の中）:
1. 書体がデザインに親和するか（`--font-id` / 04見本で目視選定）
2. マスタの **単位** が正しく出るか
3. 金丸内でバランスよく配置できるか

---

## 3. 見本フォルダ（作成済）

```text
G:\マイドライブ\05.画像生成（セットMAIN）\
  01.amazon白抜きベース/
  02.楽天ベース/
  03.amazon見本/          … 配置・大きさ・重なり
  04.楽天見本/            … 書体・金丸内バランス
  99.octas期限管理シール素材/
  00.テスト出力/          … PoC出力
```

Pillow用 97／98 は AI 本線では必須としない。

---

## 4. エンジン・モデル方針（ロック）

| 項目 | 決定 |
|------|------|
| 通称 | Google AI Studio の **NANOBANANA** ＝ Gemini Image |
| API | Gemini API（`google-genai`／Interactions 優先、失敗時 generateContent） |
| **モデル帯** | **通常 Flash Image のみ**（例: `gemini-3.1-flash-image` = Nano Banana 2） |
| **PRO** | **使わない**（`gemini-*-pro-image` 禁止） |
| **Lite** | 既定では使わない（コスト最優先に切り替えるときだけ明示） |
| **最新** | 実行時に `models.list` から通常 Flash Image の最新を選ぶ。失敗時フォールバック `gemini-3.1-flash-image`。環境変数 `SET_MAIN_IMAGE_MODEL` または `--model` で上書き可（PROは拒否） |
| キー | `GEMINI_API_KEY` または `tools/set_main_image/secrets/gemini_api_key.txt`（**リポジトリに入れない**） |

---

## 5. PoC スコープ（本承認の実装範囲）

- CLI: `tools/set_main_image/ai_compose_poc.py`
- **Amazon**: AI（役割ラベル付き見本入力）＋ `*_trace.json`（どのファイルが LAYOUT_BLUEPRINT か・プロンプト全文・apiPath）
- **楽天**: **レイヤーのみ**（`mode=rakuten_layer_only`）。フル生成はコード上も呼ばない
- 出力: `00.テスト出力/` ＋ 同名 `.json` / `*_trace.json`
- 本線一括・07接続は品質OK後

### 初回PoCで分かったこと（2026-08-02）
全面 AI edit だと楽天デザインが必ず壊れる／Amazon見本拘束も弱い → **楽天はレイヤー必須**、Amazonはトレース強化＋プロンプト拘束。

---

## 6. 検収

### PoC
- [ ] Amazon: 見本に近い重なり・大小比。白枠コピペ感がない  
- [ ] 楽天: ベースが目視で同一、数字だけ N、金丸内バランスが見本に近い  
- [ ] 楽天でデザイン改変が起きない  
- [ ] JPEG が `00.テスト出力` に出る。ログに **実際に使った modelId** がある  
- [ ] PRO モデルが選ばれていない  

### 本線（後続承認）
- 週規模コスト・再生成手順・07接続

---

## 7. リスクと緩和

| リスク | 緩和 |
|--------|------|
| 楽天デザイン改変 | 禁止プロンプト＋目視検収 |
| モデル改定 | list＋fallback。PROはコードで拒否 |
| キー漏洩 | secrets/ を gitignore。コミット禁止 |
| 課金 | PoCは少枚。本線前にコスト確認 |

---

## 8. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-08-02 | 提案起草。 |
| 2026-08-02 | **方針ロック**。見本03/04作成確認。**API PoC実装承認**。モデル＝通常Flash最新・PRO不要。 |
| 2026-08-02 | PoC結果を受け **楽天＝レイヤーのみ**／Amazon＝見本拘束＋`*_trace.json` に補正。 |
