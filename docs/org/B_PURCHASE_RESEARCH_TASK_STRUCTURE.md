# 仕入れ検討① — タスク構造化メモ（他PJ共有）

**文書種別**: 実装タスクの進捗つき構造（添付ツリーと同形式）  
**日付**: 2026-08-16  
**完成形のゴール・流れ**: [FLOW](B_PURCHASE_RESEARCH_FLOW.md)（A–J / A′。状況を聞かれたら FLOW §3）  
**要件**: [V1](B_PURCHASE_RESEARCH_V1_REQUIREMENTS.md)　**数字・ドライラン**: [NOW](B_PURCHASE_RESEARCH_NOW.md)  
**領域図**: [DOMAIN1](../DOMAIN1_RESEARCH_PURCHASING.md) §1　**倉庫**: [STORE](COMPETITOR_STORE_REQUIREMENTS.md)　**貼付**: [PASTE](B_AMAZON_COMPETITOR_PASTE_REQUIREMENTS.md)

下ツリーの **A 問屋 / B リサーチ / C 見積** は DOMAIN1 旧ラベル。完成形の A＝店起点ではない。

凡例: **済**／**次**／**後**／**外**／**出品PJ**

Keepa の正は競合DB `Keepaフル`（90日。出品と共用）。`①候補` は門の作業面（転記。倉庫ではない）。Catalog 生は `モールヒット` 禁止。出品 clasp 差し替え禁止。

---

## 構造（正）

```text
領域1
├─ A 問屋                         外
├─ B 商品リサーチ
│  ├─ ① 仕入れ検討 ★リサーチ見積もり agent
│  │  ├─ 進め方
│  │  │    Keepaの正=競合DB Keepaフル  ロック（90日。出品と共用）
│  │  │    ①候補=門の作業面（転記）   ロック（倉庫ではない）
│  │  ├─ 実行基盤
│  │  │    調査複製スプシ              済（①候補 181行。競合タブを増やさない）
│  │  │    専用 GAS purchase_research/ 済（出品 clasp 禁止。Catalog は未搭載）
│  │  ├─ 門検収 C1–C8                 済
│  │  ├─ 倉庫本線
│  │  │    W1 Keepaフルへ①書き       スライス5合格
│  │  │    W2 ①候補ASINバックフィル  済（181/181。skip5＋GET176）
│  │  │    W3 門はフル優先→①候補転記 済（メニュー20:20）
│  │  │    W4 貼付90日後読み          出品PJ（dry: paste60 skip39 need21）
│  │  ├─ 前段①–⑧
│  │  │    ① セラー貯め→ストア正     済（タブ「セラー」モリタ1行。フルから貯めない）
│  │  │    ② ピック                   済（食品root・モリタ。S4構成比）
│  │  │    ③ 巡回                     済（第1版1セラー。query total=30。productなし）
│  │  │    ④ /query GET               済（モリタ30）
│  │  │    ⑤⑥⑦ 分類・食品・メーカー  済・仮想（モリタ30）
│  │  │    ⑧ 週次メーカー決定         済（今週=永谷園。Catalog GETなし）
│  │  ├─ 品番リスト転記
│  │  │    O1–O3 出品者数 COUNT_NEW  済（QS一致。空61＝今オファーなし）
│  │  │    F0 直販・新品現在・現行順位 済（GETなし。いる34/いない186）
│  │  │    F1 Keepa列 flatten       済（cat217/梱包157/FBA164。BB価格はJSON空）
│  │  │    F2 画像／サブ画像          済
│  │  │    T1–T2 品番リスト          済（通過82 JOIN。式非書）
│  │  │    課題 K1–K8               止／流用／不要（NOW表）
│  │  └─ 見積書ルート                 後
│  └─ ② 出品用                       外（貼付 Catalog 階段と混ぜない）
└─ C 見積                             外
```

**次（①）**: フェーズ2停止（次店なし）。公式次語はふりかけ等。Eフラグは別承認。K1–K8は止。W4 本線は出品。  
**3者**: 今不要。E/F/G/J 実装前は FLOW §6。  
**次（出品・貼付）**: W4 は W1 後。洗い Catalog を貼付階段に使わない。

---

## 他PJが守ること

| 相手 | やる | やらない |
|------|------|----------|
| 出品・貼付 | W4（`Keepaフル` を90日内読む） | 洗い Catalog を A′–E′ に流用。ルート clasp を①へ差し替え |
| 競合ストア | W1 の `purpose=research` 追記を受ける。メーカーマスタ・セラー | Catalog 生を `モールヒット`。①候補を倉庫扱い。Keepaフルからセラー貯め |
| 出品 B Step2 | JAN の楽天・Yahoo（既存） | メーカー洗いで3モールAPI |
| 全体 | 楽天CSV聖域。Keepa 契約1本 | ①キーを git に置く |

載せない: ペヤング型／公式サイト照合／粗利で門打ち切り／出品マスタ書き／Keepa `csv[]`／検索1件目＝競合価格。課題 K1–K7（メーカー・FBA人数は無理にやらない。BB不要。おまけ等不要。標2b／自己発送は出品流用）正は [NOW](B_PURCHASE_RESEARCH_NOW.md)。

人間 clasp: [HUMAN_RUN](B_PURCHASE_RESEARCH_HUMAN_RUN.md)。Agent は `clasp push` しない。
