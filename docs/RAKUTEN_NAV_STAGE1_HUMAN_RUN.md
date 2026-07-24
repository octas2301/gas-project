# 楽天ジャンル Nav Stage1 — 人間実行メモ（2026-07-24）

**Agent静的確認**: PASS  
**実機検収**: **PASS（2026-07-24）**  
- オフ時: 「Stage1書込は無効です」表示  
- オン時: 成功 **2/2**／runId `navS1_20260724_203812_a61c6b17`  
- シート: `101888`/`101535` とも http200・`ok=TRUE`・マスタ非書込メッセージ確認  
- **忘れず**: Property `RAKUTEN_NAV_GENRE_STAGE1_WRITE_ENABLED` を **`false`** に戻す  

**Agent は clasp push／Property変更をしない。**

---

## 手順（5分）

1. 必要ならローカルで `clasp push`（未反映のときだけ）  
2. GAS → プロジェクトの設定 → Script Properties  
3. **オフ確認**: `RAKUTEN_NAV_GENRE_STAGE1_WRITE_ENABLED` が無い／false → メニュー **Z → 17 → 17-⑥** → 「Stage1書込は無効です」で終了すれば OK  
4. 同 Property を **`true`** にする  
5. スプレッドシート再読込 → **17-⑥** 実行  
6. シート **`▼診断(楽天ジャンルNav)`** で `101888` / `101535` の `ok` 列を確認  
7. **▼商品マスタ** / **AI情報取得data** が変わっていないことを目視  
8. Property を **`false`** に戻す  

合否表の正本: [RAKUTEN_NAV_GENRE_STAGE1.md](RAKUTEN_NAV_GENRE_STAGE1.md) §9.1  

結果（ok件数・runId）をチャットに貼れば、Agentが docs を「検収済」に更新する。
