# 競合専用ストア

仮想テスト: `python tools/competitor_store/run_tests.py`（T0–T22）  
初期化: `python tools/competitor_store/init_store.py`（項目マップ更新＋`Keepaフル`タブ保証。`--full-init` は既存ヒットを消す）
Keepa辞書再生成: `python tools/competitor_store/gen_keepa_official.py`  
90日 dry-run: `python tools/competitor_store/purge.py`  
退避 dry-run: `python tools/competitor_store/backup.py`

出品マスタの `Keepa取得_キャッシュ` は読取のみ。`--apply` と `COMPETITOR_STORE_ENABLED` は人間が付ける。
