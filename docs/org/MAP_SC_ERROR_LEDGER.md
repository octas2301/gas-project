# Amazon MAP — SCエラー記録（派生MD）

**正本は sheet**（`▼設定(Amazonマッピング)`）。本ファイルは調査・履歴用。  
実行反映は **sheet 更新 → `sync_map_sheet_to_column_json.py` → PACKAGED**。

---

## 記録テンプレ（コピーして追記）

```markdown
### YYYY-MM-DD / subBatchId / 親SKU
- **症状**:
- **SCコード／画面**:
- **原因（仮説→確定）**:
- **sheetへの結論**（attrKey / 変更内容）:
- **JSON同期**: sync_map_sheet_to_column_json --profile …（済／未）
- **再PACKAGED／再UL**:
```

---

## ログ

### 2026-08-02 / CK_5beb0cbf67ea_B2 / sanky-4538872281013-oya（UL#3・`…052208`）
- **症状**: 処理13／成功0／**その他エラー13**／失敗0。コードは **100521のみ**（90194/90225なし）。ステータス「成功 (その他のエラー)」
- **サマリ**: `…052208-processing-summary.xlsm`（05）
- **データ確認**: browse=`食品・飲料・お酒 > 缶詰・瓶詰 > 肉の缶詰・瓶詰 (71192051)`（プルダウン一致）。子の入数/サイズ/重量OK。テーマ=`サイズ`
- **原因**: Amazon側の出品情報レビュー（最大48h）。属性MAP不備ではない（七味と同型）
- **sheetへの結論**: **変更不要**。再PACKAGED／再ULは原則不要（追加情報リクエストが出たら対応）
- **次**: SC在庫・親ASINの公開／停止を確認。出たら台帳 SUCCESS

### 2026-08-02 / CK_5beb0cbf67ea_B2 / sanky-4538872281013-oya（UL#2）
- **症状**: 処理13／成功0／失敗13。**90194**＋**90225**（推奨ブラウズノード）。値=`肉の缶詰・瓶詰 (71192051)`（短縮名）
- **原因**: xlsm プルダウン外の短縮表記。正は `データを閲覧する` の **BrowsePath**（例: `食品・飲料・お酒 > 缶詰・瓶詰 > 肉の缶詰・瓶詰`）
- **対策**: Node ID→`shelf_browse_catalog` で BrowsePath 解決して列へ書く（数値IDだけ・短縮名は不可）。ダイアログ手順は成功/失敗分離済
- **再PACKAGED**: BrowsePath＋`(NodeId)` 版 `…052208.xlsm` → UL#3（100521のみ）

### 2026-08-02 / CK_5beb0cbf67ea_B2 / sanky-4538872281013-oya（UL#1 事前）
- **症状**: テーマ `SET_NAME` が GROCERY プルダウン外で全行エラー。その後 入数全6／ユニット160／重量960／ATサイズ空／色その他／形態・温度の黒セル入力
- **原因**: テーマ英語ENUM誤用；親総個数・一人分数量・総重量の継承；`size` を `package_size_name` に誤マップ；MAP RULES（PARSE_*）未実装
- **sheetへの結論**: var_theme=固定`サイズ`；number_items/unit/weight はセット数・サイズ解析；size=AT；色・形態・温度は未出力ノート
- **JSON同期**: `c1_quantity_policy`＋grocery map（2026-08-02）。以降は sheet→JSON 必須
- **再PACKAGED**: `…_20260802_044503.xlsm` UL済（人間）→上記 UL#2 で browse エラー
