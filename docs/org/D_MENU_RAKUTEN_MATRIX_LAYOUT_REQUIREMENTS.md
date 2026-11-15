# 楽天マッチング sheet 画像レイアウト拡張 — 要件

**文書種別**: requirements  
**状態**: 2026-09-05 チャットロック・実装  
**関連**: [D_MENU_RAKUTEN_IMAGE_SKU_AUTOBIND_REQUIREMENTS.md](D_MENU_RAKUTEN_IMAGE_SKU_AUTOBIND_REQUIREMENTS.md)／[D_MENU_CDE_MENU_SPLIT_REQUIREMENTS.md](D_MENU_CDE_MENU_SPLIT_REQUIREMENTS.md)

---

## 1. 目的

C（楽天画像マッチング）生成時に、人間の並び替えを減らしつつ次を自動配置する。

- 他セットMAINの訴求（メイン2〜10）
- サブの全子共通＋訴求順
- 親行＝最少セット子と同じ見せ方

R-Cabinet は既存どおり **同一 Drive fileId は1回のアップ実行で送信1回**（`uploadedUrlMap`）。

---

## 2. 置き場（正）

| 段階 | 場所 | 中身 |
|------|------|------|
| 入力 | Drive **02**（ソース） | 既存 MAIN／サブ |
| 作業面 | **★画像AIマッチング** の青枠（楽天メイン1〜10・サブ1〜10） | GASが `=IMAGE(...fileId...)` を **直接書込** |
| 反映 | D 画像アップ | fileId → R-Cabinet → マスタURL |

**禁止（本線）**: 中間 staging フォルダ／人手の切り取り貼付を前提にしない。手直しは目視救済のみ。

---

## 3. ルール

| ID | 内容 |
|----|------|
| L0 | **子・メイン1**: 現行どおり。ファイル名に子SKU → 空のメイン1へ（既存非上書き） |
| L1 | **子・メイン2〜10**: 同一親の **他子メイン1** を **総個数（セット数）昇順**で埋める（自分除外・空きのみ） |
| L2 | **子・サブ1〜10**: 親のサブを **訴求順で全対象子へ同一 fileId で上書き**（空きのみにしない。親空枠は子も空）。親が全空なら最少セット子から親へ昇格してから配る。L3後に再実行 |
| L3 | **親行**: メイン1〜10・サブ1〜10 ＝ **最少総個数の子** と同じ内容（同じ fileId）。同値なら display 順の先頭子 |
| L4 | Drive ファイルは **複製しない**。IMAGE 数式（fileId）のコピーのみ |
| L5 | トグル: `RAKUTEN_MATRIX_LAYOUT_EXPAND_ENABLED`（未設定=**true**。false で L1〜L3オフ・L0のみ） |
| L0s | **サブ自動**: ファイル名に `_subN` がある画像は **親行のサブNのみ**（子SKU一致でも子サブへ直結しない）。L2で全子共有 |

### セット数の解決

1. マスタ **総個数**列（あれば）  
2. なければバリエーション名から先頭の整数（例: `12袋/600g` → 12）  
3. 取れなければ大きい値扱い（他セット並べの末尾）

### JAN／親レベルサブ

JAN（10〜13桁）が親SKUに含まれ、かつ `_subN` がある画像は **親グループのサブN** に載せ、L2で全子へ配る。  
子SKUを含む `{子SKU}_…_subN.jpg` も **親のサブN** へ載せ、L2で全子同一にする（子行へ直結しない）。

---

## 4. フロー

```text
02
 → C マトリクス生成
    → L0 ファイル名 autobind（＋JANサブ）
    → （既存）余りは AI／候補
    → L1〜L3 レイアウト拡張（Property ON時）
 → 人間目視
 → D 画像アップ（fileId キャッシュ）
```

---

## 5. 受け入れ

- [ ] 最少セット子と親行のメイン／サブ fileId が一致
- [ ] 各子メイン1＝自SKU、メイン2〜10＝他セット昇順（空き埋め）
- [ ] 全子のサブ1〜10が同一 fileId・同一順
- [ ] 1回のアップで同一 fileId は R-Cabinet 送信1回
- [ ] Property false で L1〜L3 が走らない
- [ ] 中間フォルダ／切り取り貼付が本線に無い

---

## 6. 実装メモ

- `コード.js`: `rakutenMatrixLayoutExpand_`／`rakutenMatrixShareSubsForce_`／`rakutenFilenameAutobind`（SUB_TO_PARENT）／`generateAiImageMatrix` から呼出
- 戻し: `RAKUTEN_MATRIX_LAYOUT_EXPAND_ENABLED=false`（L1〜L3）。サブの親寄せはコード差戻し

---

## 7. 更新履歴

| 日付 | 内容 |
|------|------|
| 2026-10-22 | L2を上書き共有に変更。`_subN` は親のみ。L3後に再共有。子ごとの別Drive ID増殖を止める。要 clasp |
| 2026-10-19 | `rakutenMatrixLayoutExpand_` が `コード.js` から消えていたため、同じルールで再掲。要 clasp |
| 2026-09-05 | 初版ロック・実装 |
