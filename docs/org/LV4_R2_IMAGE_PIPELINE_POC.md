# Lv4 Amazon 画像パイプライン設計（Drive起点 GAS 案）

**文書種別**: 設計（**GAS／Python 実装コードは未着手・要別承認**）  
**最終更新**: 2026-07-24（T2チケット詳細化・T1済反映）  
**親**: [LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md](LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md)（§11.0・§11.5・§1.4）  
**状態**: Drive起点案を正とする。サブは **REUSE_RAKUTEN / AMAZON_ONLY** 二モード対応が必要（社長意向・2026-07-24）。実装コードは未着手。

---

## 0. スコープ（確定）

| 対象 | 本設計 |
|------|--------|
| **Amazon**（Drive画像 → R2／ZIP／GENERATED連携） | **対象** |
| 楽天（`03.…\02.楽天アップロード画像保存場所` → RMS） | **サブ流用の正本候補**（参照）。Amazonパイプライン実装で `03` を壊さない |
| Yahoo!（楽天画像流用） | **Amazon設計では触らない**（現状どおり楽天流用） |

**一文**: MAINは常に Amazon 白抜き（`04\02`）。サブは **楽天ライン流用**または **Amazon単独PT** の二モード。どちらも `{SKU}.PT0n` に正規化してから R2／ZIP／GENERATEDへ。xlsm仕上げは当面 Cursor。GAS実装は別承認。

本運用の画像の正は **SC ZIP**（§11.0）。R2はステージング。

---

## 1. なぜ Drive起点か

楽天は既に:

`G:\マイドライブ\03.楽天・Yahoo!商品登録（CSV一括UL）\02.楽天アップロード画像保存場所`

へ画像を置き、**GASメニューで RMS へアップ**している。  
Amazonも同じ習慣に揃えれば、**ローカルパスをGASが読めない問題を避けつつ、メニュー中心**に寄せられる。

ファイル名（`{SKU}.MAIN.jpg`）は GAS でも生成・検証可能。論点は「名前が作れない」ではなく **バイナリの所在が Drive であること**。

---

## 2. Drive フォルダ構成（社長作成の `04` を活用）

根: `G:\マイドライブ\04.amazonカタログ作成（CSV一括UL）\`  
（Google Drive 上。GASは Folder ID で参照）

**フォルダ実体**: **2026-07-24 作成済**（01〜06）。ID控えのローカル正本: `Downloads/Lv4_Amazon_PACKAGED/DRIVE_04_FOLDER_IDS.md`（Git外）。

| サブフォルダ | Folder ID | 役割 | 誰が書く |
|--------------|-----------|------|----------|
| **04（根）** | `1K43z1vFW_QO41FN0m5Z_58qtdQY4ubor` | Amazon Drive 根 | — |
| **01.GENERATED保存先** | `15aMcznDKvKD7xco-bdj_0sDF0XMdhO5p` | 21-①埋め用データ（既存Drive運用と統合可） | GAS |
| **02.Amazonアップロード画像保存場所** | `1T6_E6T-qd9whSF8Re8lyRVB2n-P4BM84` | ★楽天02と同役割。人間が jpg を置く | **人間** |
| **03.PACKAGED_xlsm保存先** | `11juFW8OE-7x_QDvCdok6zrbM_j4s27o1` | 完成 `.xlsm` の置き場 | 人間／Cursor／（将来）ローカル処理 |
| **04.SC用画像ZIP保存先** | `1nGleQOSjcK47CEnR3cgVVjLkWNYfQRyq` | `{SKU}.MAIN.jpg` 等の ZIP | GAS（将来）または手作業 |
| **05.処理レポート・ログ** | `1Z2iBeWSaT5UhERrInRkUr3Me67O5oG-s` | processing-summary 退避 | 人間 |
| **06.テンプレ原本（純正xlsm）** | `1d8AEpOhx_1ymvskUeA5moZT_Si0SFGBS` | `HEALTH_PERSONAL_CARE.xlsm` / `FOOD.xlsm` 等・読取専用 | 人間（SCから取った純正を配置） |

**Script Properties への登録・GAS実装は別承認まで禁止**（キー案は `DRIVE_04_FOLDER_IDS.md`）。

### 02 配下の命名規約（確定方針）

```text
02.Amazonアップロード画像保存場所\
  {subBatchId または 親SKU}\
    {sellerSku}.MAIN.jpg
    {sellerSku}.PT01.jpg   … 任意
```

R2オブジェクトキーも **ファイル名と同じ**（`{sellerSku}.MAIN.jpg`）にする。

---

## 2.1 サブ画像の二モード（社長意向・必須対応）

Amazon MAINは **白抜き専用**のため楽天／Yahooと別。サブは多くの場合楽天・Yahooと共通だが、**Amazon単独出品**もある。

| モード | いつ | MAIN | サブの出所 |
|--------|------|------|------------|
| **`REUSE_RAKUTEN`** | 楽天／Yahooに既にある・同時出品 | `04\02\{SKU}.MAIN.jpg`（白抜き） | マスタ「楽天サブ画像1〜n」または `03\02` → **PT01…へ番号マップ**（中身判定なし） |
| **`AMAZON_ONLY`** | 楽天に無い／後からAmazonだけ | 同上 | `04\02\{SKU}.PT0n.jpg` のみ |

**切替**: SKUまたは親単位の明示フラグ（マスタ列 or バッチ指定）。  
**フォールバック**: `REUSE` だが楽天サブ0枚 → `AMAZON_ONLY` として `04` のPTを見る（ログに残す）。  
**安定性**: 楽天CDN／Yahoo URLのAmazon直埋めはしない。Driveからコピー／改名して R2・ZIPへ（スナップショット推奨）。  
**PT順番**: サブ画像1＝PT01候補、…。Amazonギャラリーは MAIN→PT01→PT02…。

共通出口（モード差をここで吸収）:

```text
MAIN必須（04） + SUB一覧（モード別） → {SKU}.MAIN / {SKU}.PT0n → R2 および/または ZIP → GENERATED
```

---

## 3. 操作フロー（目標像）

```text
① 人間
   02 に画像を置く（楽天と同じ習慣）

② GAS（将来メニュー・仮称）
   21-⑥ Drive→R2 アップ＋URLをログ／シートへ
   21-⑦（任意）04 に SC用ZIP生成

③ GAS（既存）
   21-① GENERATED
   （画像URL列に R2 URL を載せる／参照可能にする）

④ xlsm 仕上げ（※§5・未確定）
   当面: Cursor／手作業で 06純正＋GENERATED → 03 へ .xlsm
   将来候補: ローカル短い処理（精度を見てから確定）

⑤ 人間
   SC で xlsm UP ＋ 必要なら 04 の ZIP
   21-③ UPLOADED_OK ／ ENABLED=false
```

楽天より段が1つ多い理由: **Amazonは純正 `.xlsm` が必須**だから。画像まわりは楽天と同型にできる。

---

## 4. GAS で担う範囲（実装時の設計）

| 機能 | 内容 | 実装状態 |
|------|------|----------|
| Drive走査 | 02（または指定 subBatch フォルダ）の jpg 一覧 | 未実装 |
| ファイル名検証 | `{SKU}.MAIN.jpg` / `PT0n` 形式 | 未実装 |
| R2 Put | S3互換署名付き PUT（Secretは Script Properties） | **T2実装済**（21-⑥・1枚PoC） |
| URL記録 | `▼Lv4実行ログ(Amazon)` または専用シート／GENERATED連携 | 未実装 |
| ZIP生成 | 04 へ `{親SKU}_MAIN_images_for_SC.zip` | 未実装・推奨（18320対策） |
| メニュー | 21-⑥／21-⑦（名称は実装時確定） | 未実装 |

**触らない**: 楽天CSV・`Yahoo.js`・マスタJAN／在庫一括・純正xlsmのバイナリ直接編集。

### Script Properties（実装時想定・キーはコードに書かない）

```text
APPROVAL_AMAZON_LV4_ENABLED          … 既存（再GENERATED用。画像メニューは別フラグでも可）
AMAZON_DRIVE_IMAGE_FOLDER_ID         … 02 の Folder ID
AMAZON_DRIVE_ZIP_FOLDER_ID           … 04
AMAZON_DRIVE_GENERATED_FOLDER_ID     … 01（既存と共用可）
R2_ACCOUNT_ID / R2_ACCESS_KEY_ID / R2_SECRET_ACCESS_KEY / R2_BUCKET / R2_PUBLIC_BASE
```

公開ベース（現状）: `https://pub-d974bd81c7d84f9bbc65f8479d3f85d4.r2.dev`  
バケット（現状）: `octas-amazon-imag`

---

## 5. xlsm 仕上げ（⑤）— **提案のみ・実装は後確定**

### 5.1 社長懸念（採用）

Amazonバルクの **項目・選択肢・推奨値はテンプレ更新で変わり得る**。  
固定スクリプトだけで毎回塗り続けると、列ずれ・許容値外れで **精度が落ちる**可能性がある。  
よって **Cursor（都度テンプレ／成功DBを見ながら埋める）を当面の正**とし、完全自動ローカル埋めは急がない。

### 5.2 選択肢（後で精度を見て選ぶ）

| 案 | 内容 | 精度リスク | 備考 |
|----|------|------------|------|
| **C0（当面の正）** | Cursor／手作業で PACKAGED → 03へ保存 | 低い（都度確認） | いまの titlefix ルート |
| **C1（提案）** | ローカル短い処理: 06純正＋GENERATED＋URL → 03 | 中〜高（テンプレ変更に弱い） | **精度観察後に採否** |
| **C2（提案）** | C1＋`accepted_values_db`／最新純正の差分チェックを必須化 | 中 | 自動の前提条件が重い |

**確定ポリシー（2026-07-24）**:

- ⑤の自動実装は **まだしない**  
- C1/C2は提案として残す  
- 採否は「SC成功率が手作業／Cursorと同等以上」を見てから  

GASが苦手なのは「Excelが一切できない」ではなく、**純正 `.xlsm`（VBA・行5属性マップ）を安全に量産すること**。スプレッドシート／CSV／ZIPはGAS向き。

---

## 6. 旧 PoC（ローカル Python 一式）との関係

| 方式 | 位置づけ |
|------|----------|
| **Drive起点 GAS**（本節） | **正（進める）** |
| ローカル jpg → Python で R2＋xlsm | **副案**。Drive運用が回れば不要寄り。xlsmのみC1として残す可能性 |

---

## 7. 実装チケット分割（別承認ごと）

| 順 | 内容 | 依存 | 状態（2026-07-24） |
|----|------|------|-------------------|
| T1 | `04` フォルダ01〜06＋Folder ID控え | — | **済**（Property未登録） |
| T2 | GAS: Drive→R2（1SKU PoC）＋ログ | T1・R2キー | **済**（21-⑥・runId `R2T2_20260724_221107_7f9cf7`） |
| T3 | GAS: ZIP生成→04 | T2 | **実装待ち**（手ZIP＝当面の正。§11.0 18320をT2不足証跡。[D_MENU_AMAZON_FACADE_REQUIREMENTS.md](D_MENU_AMAZON_FACADE_REQUIREMENTS.md) §6.4。別実装承認） |
| T4 | GENERATED への画像URL連携 | T2 | 未着手（D要件 U4） |
| T5 | xlsm自動（C1） | 精度レビュー後 | **当面スキップ** |

**C（画像紐付け）**: 本線は **案α**（既存 `★画像AIマッチング` 拡張）。Amazon MAIN は **子SKU行で人間紐付け** → Drive `02` 出力。サブは `REUSE_RAKUTEN`／`AMAZON_ONLY`。将来 **ε**（自動＋修正のみ）はバックログ。詳細 [D_MENU_U2_C_AMAZON_IMAGE_REQUIREMENTS.md](D_MENU_U2_C_AMAZON_IMAGE_REQUIREMENTS.md)。親 [D_MENU_AMAZON_FACADE_REQUIREMENTS.md](D_MENU_AMAZON_FACADE_REQUIREMENTS.md) §7。

コード実装は各Tiで **変更ファイル一覧／概要／リスク**を提示してから。

### 7.1 T2 実装チケット（承認用・コード未着手）

**目的**: Drive `02` の1SKU（例: 既存試験SKUの MAIN 1枚）を R2 に Put し、公開URLをログに残す。

| 項目 | 内容 |
|------|------|
| スコープ内 | Drive App で Folder ID `02` 走査、`{SKU}.MAIN.jpg` 検証、R2 S3互換 PUT、`Logger.log`（runId／sku／http／url） |
| スコープ外 | ZIP量産、GENERATED自動埋め、xlsm編集、楽天`03`書込、全件ループ、本番一括 |
| 想定変更ファイル | 新規 `AmazonDriveImageExport.js`（仮）／`コード.js` メニュー21-⑥のみ／docs・CHANGE_LEDGER |
| Script Properties（実装時） | `AMAZON_DRIVE_IMAGE_FOLDER_ID`＝02、`R2_*`（既存運用と共用可）、トグル `AMAZON_DRIVE_R2_UPLOAD_ENABLED` 既定 **false** |
| 検収 | ENABLED系トグルオフで何もしない／オンで1枚 Put＋URLがブラウザ200／マスタ・楽天CSV差分なし |
| リスク | 秘密鍵漏洩（Propertiesのみ）／誤Folder（楽天03を触らない定数）／18320はZIP本線のためT2単独ではSC成功を断言しない |
| 復元 | Property false＋メニュー削除＋`git revert` |

**承認コメント例**: 「T2のみ承認。T3以降は別」  

**実装ファイル**: `AmazonDriveImageExport.js`／メニュー21-⑥／手順 [LV4_T2_HUMAN_RUN.md](LV4_T2_HUMAN_RUN.md)

---

## 8. リスク

| リスク | 緩和 |
|--------|------|
| r2.dev → SC 18320 | ZIP（04）を本線。R2はステージング |
| DriveのG:パスとGAS | GASは **Folder ID** のみ使う（パス文字列に依存しない） |
| テンプレ陳腐化で自動埋め精度低下 | §5: 自動は後回し。Cursor当面正 |
| 秘密鍵 | Script Properties のみ。Git／チャット禁止 |
| 楽天フォルダ誤操作 | Amazonは `04` のみ。`03` は参照のみ |

---

## 9. 人間チェックリスト（設計段階でできること）

1. ~~`04` 配下に 01〜06~~ → **済**  
2. ~~06 に純正 HPC／FOOD~~ → **済**  
3. ~~02 に試験 MAIN~~ → **済**（HPC／FOOD）  
4. Folder ID 控え → Drive `05/DRIVE_04_FOLDER_IDS.md`  
5. GAS／Pythonコードは **T2承認まで書かない**  

---

## 10. 参照

- 楽天同型: `03.楽天・Yahoo!商品登録（CSV一括UL）\02.楽天アップロード画像保存場所`  
- [LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md](LV4_AMAZON_ORCHESTRATION_REQUIREMENTS.md) §1.4・§11.0  
- [D_MENU_AMAZON_FACADE_REQUIREMENTS.md](D_MENU_AMAZON_FACADE_REQUIREMENTS.md)（本線D・T3保留・C案α）  
- [CURRENT_PHASE.md](../CURRENT_PHASE.md) §0  
- R2: `octas-amazon-imag` / `pub-d974bd81c7d84f9bbc65f8479d3f85d4.r2.dev`

### 更新履歴（抜粋）

| 日付 | 内容 |
|------|------|
| 2026-07-25 | C: 案α本線・MAIN=sheet／02=出口・εバックログ（U2方針）を §7 相当に反映。 |
| 2026-07-24 | T2実装: `AmazonDriveImageExport.js`＋21-⑥。人間手順 LV4_T2_HUMAN_RUN。 |
| 2026-07-24 | T2済・T3保留（D要件§6.4）・C案αを §7 に反映。 |
| 2026-07-24 | Drive起点＋サブ二モード。 |
