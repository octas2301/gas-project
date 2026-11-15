---
name: b5-review-fresh
description: >-
  Relaunches B-⑤ Rakuten MAIN review_loop with --fresh, or mix (theme_new+competitor)
  review with per-product manifests, so the user always gets current review code and
  the correct folder. Use when opening or reopening MAIN or mix review, ports
  8865/8866/8780/8781, stale/outdated screen, catch/UI edits not appearing after reload,
  same-JAN multi product mix confusion, or immediately after changing review_loop /
  review_launch / mix_review_*.py.
---

# B-⑤ レビュー freshen（MAIN＋混在）

## いつ使うか

- レビューを開く・開き直す・最新か聞く
- 戻らない／直ってない／まだ古い／反映されない（レビュー画面）
- キャッチ欄／再生成 UI／「いまやること」を直した直後に画面を見てもらう
- 同じJANで別商品の混在レビューを同時に開く／取り違え疑い
- `b5_main_base_review_*.py` または `sub_image_mix_review_*.py` を変更した同じターン

## 手順（MAIN・必須）

1. `cd` は `tools/set_main_image`
2. 対象親が分かれば `--prefix`（例: `ishi-4906283045119-oya3`）。複数可
3. **必ず `--fresh`**（旧 Listen／旧 `review_loop` を止めてから起動）
4. `--no-wait` で起動し、出力の `URL http://127.0.0.1:PORT/` と `dir=` をそのまま渡す
5. ブラウザへは **「この URL を開き直す」**（古いタブの F5 だけでは足りない、と一言）

```powershell
cd C:\Users\takuy\Desktop\gas-project\tools\set_main_image
python b5_main_base_review_launch.py --fresh --prefix ishi-4906283045119-oya3 --prefix ishi-4906283045119-oya2 --base-port 8865 --no-wait
```

## 手順（混在レビュー・必須）

1. 旧ポート（例 8780/8781）のプロセスを止める
2. **商品ごとに** `--theme-new-dir`（`overview_copy_…`）と `--base-port` を分けて起動
3. マニフェストは `unique_mix_manifest_path`（`mix_manifest_{JAN}_{フォルダ名}.json`）。JANだけの1ファイルに戻さない
4. 渡す前に、画面上部の **商品名と作業フォルダ** が対象と一致することを確認する

```powershell
cd C:\Users\takuy\Desktop\gas-project\tools\set_main_image
python sub_image_mix_review_launch.py --jan 4906283045119 --theme-new-dir "C:\Users\takuy\Desktop\gas-project\tools\set_main_image\_out\overview_copy_4906283045119_20261025_172740" --base-port 8780 --no-wait
python sub_image_mix_review_launch.py --jan 4906283045119 --theme-new-dir "C:\Users\takuy\Desktop\gas-project\tools\set_main_image\_out\overview_copy_4906283045119_20261025_174533" --base-port 8781 --no-wait
```

## やってはいけないこと

- 前ターンの URL を再利用して「開いてください」だけ言う
- `--compose-dir`／`--theme-new-dir` にチャット記憶の古いタイムスタンプを入れる
- 同じJANで `mix_manifest_{JAN}.json` だけを共有して同時レビューする
- レビューコード変更後にプロセス殺しなしで「直った」「再読込すれば見える」と言う
- 人間の F5 だけをコード反映の手段にする

## 補足

- フォルダが最新でも、**動いている Python が古い**と同じ画面バグが残る。起動し直しがその対策。
- MAIN一覧だけ見る: `python b5_main_base_review_launch.py --list-latest`
- 混在の正本: ALIGN §7.10
