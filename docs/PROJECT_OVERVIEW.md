# EC業務プロジェクト 全体像

## 目的

Amazon・楽天・Yahoo! での売上高を、**利益の出る状態で早く拡大する**こと。

## 対象モール

- Amazon
- 楽天
- Yahoo!

## 6領域の枠組み（役割）

範囲は **6領域全部**。gas-project の実装が厚いのは **クリティカルパス（出品まで＝主に領域2）**。領域1の本線は 2026-08-14 から立て直し（正本: [DOMAIN1_RESEARCH_PURCHASING.md](DOMAIN1_RESEARCH_PURCHASING.md)）。

```text
目的: 利益付きで売上拡大
│
├─ 領域1 リサーチ・見積   役割: 「何を仕入れるか」。完成形は A–J / A′（[FLOW](org/B_PURCHASE_RESEARCH_FLOW.md)）
│     第1版スライスは D まで。F/J 送信は人でも可。問屋開拓は後回し。②は他PJキャッチアップ
├─ 領域2 出品            役割: マスタを三モールに載せる（gas-project 主戦場）
│     楽天CSV聖域 / Yahoo API / Amazonバルク / B統合
├─ 領域3 販売拡大        広告・実績。一部実装（タイムセール等）
├─ 領域4 商品管理        在庫・発注・物流。在庫AppSheetは別
├─ 領域5 顧客対応        要件段階。送信禁止
└─ 領域6 財務・経理      要件段階
```

| 領域 | 概要 |
|------|------|
| **1. リサーチ・見積もり** | 問屋・メーカーリサーチ、商品リサーチ、見積依頼・見積取得 |
| **2. 出品** | 商品出品情報の作成、Amazon/楽天/Yahoo! 出品・広告設定 |
| **3. 販売拡大** | 広告運用の調整、販売実績の取得・分析 |
| **4. 商品管理** | 適正発注数・発注アラート、在庫管理・入出荷・棚卸・物流 |
| **5. 顧客対応・レビュー** | 問い合わせ対応、返品・交換・クレーム、レビュー・Q&A返信 |
| **6. 財務・経理** | 売上・仕入計上、振込・手数料管理、利益集計・税務 |

上記のほとんどを **AI を活用して効率化・省人化** することを目指す。

## 大きいマイルストーン（管理用）

| ID | 内容 | 状態 |
|----|------|------|
| M1 | 基盤（要件・組織・承認） | ほぼ済 |
| M2 | 出品が回る（領域2） | 進行中・別スレッド（B統合ハード死・Amazon Lv4 等） |
| **M3** | **領域1: 外注なしリサーチ→見積下書き** | **領域1担当の本線**。[DOMAIN1](DOMAIN1_RESEARCH_PURCHASING.md)／[第1版要件](org/B_PURCHASE_RESEARCH_V1_REQUIREMENTS.md) |
| M4 | 販売拡大 | 広告・タイムセールは一部。実績分析は薄い |
| M5 | 在庫・発注 | 在庫AppSheetは別。発注本線は後 |
| M6 | CS・経理 | 要件段階 |

## 前提

- **体制**: 社員2名（1名は副業体制）
- **リソース**: 不足分は作業外注で補填
- **定型化**: 業務を定型化し、**日・週・月** でやることを分けて整理する
- **出庫・発送**: お客さまへの出庫は（1）FBA等モール倉庫への納品、（2）自己発送の2通り。いずれも**障がい者施設へ外注**しており、自社は施設へ商品を納品する、またはメーカー・問屋から施設へ納品する形。自己発送で施設がやりきれないもの・数が少ないものは**自宅から発送**。詳細は [REQUIREMENTS.md](REQUIREMENTS.md) の「4.1 お客さまへの出庫・発送」を参照

## 参考資料

- **メーカー・問屋仕入れ**: 仕入先選定・見積の前提として参照。
- **出品自動化スプシデータ**: プロジェクト内フォルダ `参考資料：出品自動化スプシデータ` に、現時点のマスタシート等をCSVで保存したファイルがある。このCSVの項目で Amazon・楽天・Yahoo! に必要な情報を揃えている。項目の確認はこのフォルダのCSVを参照する。

## 成果物・関連ドキュメント

| ドキュメント | 内容 |
|--------------|------|
| [AGENT_HANDOVER.md](AGENT_HANDOVER.md) | **エージェント向け引き継ぎ指示**。全 agent が読む資料一覧・共通認識・双方向インプットの必須ルール |
| [REQUIREMENTS.md](REQUIREMENTS.md) | 6領域ごとの要件定義・タスク一覧・AI効率化・優先度 |
| [FLOW_AND_PRIORITY.md](FLOW_AND_PRIORITY.md) | **フロー間の必須要件・クリティカルパス・自動化構築の優先順位**（細かい要件は各フロー構築時に詰める） |
| [MASTER_LINKAGE_TASKS.md](MASTER_LINKAGE_TASKS.md) | **既存マスタ連携の実装タスク一覧**（Phase 1〜3 の具体タスク） |
| [RUNBOOK_DAY_WEEK_MONTH.md](RUNBOOK_DAY_WEEK_MONTH.md) | 日次・週次・月次の定型タスク一覧（runbook） |
| [ROADMAP.md](ROADMAP.md) | 優先度・依存関係に基づくフェーズ案（ロードマップ） |
| [AMAZON_REQUIREMENTS.md](AMAZON_REQUIREMENTS.md) | Amazon 出品の要件整理 |
| [DOMAIN1_RESEARCH_PURCHASING.md](DOMAIN1_RESEARCH_PURCHASING.md) | **領域1の範囲・優先・調査ファイル／マニュアルURL・見積は下書きまで** |
| [org/B_PURCHASE_RESEARCH_V1_REQUIREMENTS.md](org/B_PURCHASE_RESEARCH_V1_REQUIREMENTS.md) | **①第1版要件**（品質CP＝Amazon検索の救い。Keepa件数キャップなし） |
| [org/B_PURCHASE_RESEARCH_NOW.md](org/B_PURCHASE_RESEARCH_NOW.md) | **領域1① いまここ** |
| [org/B_PURCHASE_RESEARCH_TASK_STRUCTURE.md](org/B_PURCHASE_RESEARCH_TASK_STRUCTURE.md) | **①の進捗つき構造（ツリー・他PJ共有）** |
| [RESEARCH_AND_ESTIMATE.md](RESEARCH_AND_ESTIMATE.md) | リサーチ・見積もりの整理（②出品用の詳細はここ＋SESSION） |
| [AI_LISTING_AND_TITLE_REQUIREMENTS.md](AI_LISTING_AND_TITLE_REQUIREMENTS.md) | 商品名・説明・AI提案の要件（列対応・AIおすすめ商品名・開発残） |

楽天・Yahoo! 出品の**開発・実装の詳細**は、本プロジェクト直下の [HANDOVER.md](../HANDOVER.md) を参照すること。**他エージェント**は [AGENT_HANDOVER.md](AGENT_HANDOVER.md) に従い、gas-project の資料をインプットしたうえで共通認識で開発し、新要件・前提は docs に反映すること。
