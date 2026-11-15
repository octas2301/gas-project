# -*- coding: utf-8 -*-
"""
Keepa 実行中ロジックのガント（現状 vs 案1+3 改善後）。

出力: tools/common_ss_view/out/keepa_logic_gantt.html
"""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

OUT_DIR = Path(__file__).resolve().parent / "out"
HTML_PATH = OUT_DIR / "keepa_logic_gantt.html"
JSON_PATH = OUT_DIR / "keepa_gantt_data.json"

T0 = datetime(2026, 8, 31, 22, 55, 26)
WINDOW_MIN = 35  # 30分超の「続き予約→発火」まで表示
CASE = "pr_20260831_225524"
PARENT = "ハニーマザー"


def ev(iso: str) -> float:
    dt = datetime.strptime(iso, "%Y-%m-%dT%H:%M:%S")
    return round((dt - T0).total_seconds() / 60, 2)


# 経過分（実ログ）
T_TARGET = ev("2026-08-31T22:56:32")  # Keepa開始
T_CKPT0 = ev("2026-08-31T22:56:56")
T_EXPECT = ev("2026-08-31T22:58:55")
T_OVERLAP = ev("2026-08-31T22:59:18")
T_LIVE1 = ev("2026-08-31T22:57:16")
T_LIVE2 = ev("2026-08-31T23:12:13")
T_PICK_BEGIN = ev("2026-08-31T22:55:30")
T_PICK_END = ev("2026-08-31T22:56:02")
T_SLICE_END = ev("2026-08-31T23:25:16")
T_MISSED = ev("2026-08-31T23:25:11")
T_EXIT = ev("2026-08-31T23:25:22")
T_BUDGET = 20.0  # PR_GAS_BUDGET_MS_ = 20分（設計）
T_LIVE3 = ev("2026-08-31T23:27:21")  # timeup 後の見回り
T_RESERVE2 = ev("2026-08-31T23:25:21")  # timeup 後の続き予約
T_EXPECT_CONT = ev("2026-08-31T23:27:20")  # 続き発火予定
T_FIRE2 = ev("2026-08-31T23:27:34")  # 次 GAS 実行 fire
T_GAS_WALL = 30.0  # 1回目 GAS の30分壁目安（課金枠）


def fmt_dur(minutes: float) -> str:
    s = int(round(minutes * 60))
    if s < 60:
        return f"{s}秒"
    m, sec = divmod(s, 60)
    if sec:
        return f"{m}分{sec}秒"
    return f"{m}分"


# 実績ベース（①ログ pr_20260831_225524、T0=22:55:26）
ACTUAL_STATS = [
    {
        "id": "gas_budget",
        "name": "1回の時間上限",
        "design": "20分で区切る（コード上の目安）",
        "actual": f"上限到達後も {fmt_dur(T_EXIT - T_BUDGET)} 走行",
        "actualMin": T_BUDGET,
        "actualEndMin": T_EXIT,
        "clock": "23:15:26 到達 → 23:25:22 まで継続",
        "note": "取込みの途中では止まらない。20分は「線」だけ",
    },
    {
        "id": "gas_run",
        "name": "作業そのもの",
        "design": "20分以内に終了想定",
        "actual": fmt_dur(T_EXIT),
        "actualMin": T_EXIT,
        "clock": "22:55:26 開始 → 23:25:22 timeup",
        "note": "選定〜取込み〜終了まで同一GAS",
    },
    {
        "id": "orchestration",
        "name": "選定（メーカー決め）",
        "design": "—",
        "actual": fmt_dur(T_PICK_END - T_PICK_BEGIN),
        "actualMin": round(T_PICK_END - T_PICK_BEGIN, 2),
        "clock": "22:55:30 → 22:56:02（phase=pick）",
        "note": "取込み前の準備",
    },
    {
        "id": "inflight",
        "name": "「いま作業中」の札",
        "design": "作業と同じ長さ",
        "actual": fmt_dur(T_EXIT),
        "actualMin": T_EXIT,
        "clock": "22:55:26 立てる → 23:25:22 外す",
        "note": "finally まで外れない",
    },
    {
        "id": "wash_lock",
        "name": "メーカーの鍵",
        "design": "取込み中",
        "actual": fmt_dur(T_EXIT - T_TARGET),
        "actualMin": round(T_EXIT - T_TARGET, 2),
        "clock": "22:56:32 → 23:25:22",
        "note": "ハニーマザー固定",
    },
    {
        "id": "keepa_slice",
        "name": "商品データの取込み",
        "design": "1スライス約15分想定",
        "actual": fmt_dur(T_SLICE_END - T_TARGET),
        "actualMin": round(T_SLICE_END - T_TARGET, 2),
        "clock": "22:56:32 開始 → 23:25:16 slice=1完了",
        "note": "1回のAPI塊が28分超",
    },
    {
        "id": "ckpt_trigger",
        "name": "目覚まし（続き予約）",
        "design": "2分後に起こす",
        "actual": f"待機{fmt_dur(T_OVERLAP - T_CKPT0)}→拒否",
        "actualMin": round(T_OVERLAP - T_CKPT0, 2),
        "clock": "22:56:56 予約 → 22:59:18 発火",
        "note": "鳴ったが作業中で無効",
    },
    {
        "id": "expect",
        "name": "起こす予定時刻メモ",
        "design": "発火で消える",
        "actual": fmt_dur(T_MISSED - T_CKPT0) + "残存",
        "actualMin": round(T_MISSED - T_CKPT0, 2),
        "clock": "22:56:56 記録 → 23:25:11 誤判定",
        "note": "overlap 後も消えない",
    },
    {
        "id": "overlap",
        "name": "二重起動の拒否",
        "design": "瞬時",
        "actual": "1回・瞬時",
        "actualMin": 0,
        "clock": "22:59:18 overlap_skip",
        "note": "別GASが即 return",
    },
    {
        "id": "ckpt_trigger2",
        "name": "目覚まし（timeup後）",
        "design": "終了後2分で起こす",
        "actual": fmt_dur(T_FIRE2 - T_EXIT),
        "actualMin": round(T_FIRE2 - T_EXIT, 2),
        "clock": "23:25:21 予約 → 23:27:34 fire",
        "note": "30分チャート外にあった「本当の続き」",
    },
    {
        "id": "liveness",
        "name": "見回り番",
        "design": "15分ごと",
        "actual": "点検2回（+終了後1回）",
        "actualMin": None,
        "clock": "22:57:16 / 23:12:13（周期≈15分）",
        "note": "各回は別GAS・ログ1行（数秒以内）。rearmed=0",
    },
]

LANES = [
    {
        "id": "gas_run",
        "label": "作業そのもの",
        "role": "メーカー洗いの一連の処理を動かしている本体。始まると終わるまで同じ作業が続きます。",
        "order": 0,
    },
    {
        "id": "inflight",
        "label": "「いま作業中」の札",
        "role": "他のタイマーが来ても「いま動いているから触らないで」と伝える札。終わったら外します。",
        "order": 1,
    },
    {
        "id": "wash_lock",
        "label": "メーカーの鍵",
        "role": "同じメーカーを同時に二重で洗わないよう、一時的に鍵をかけます。",
        "order": 2,
    },
    {
        "id": "keepa_slice",
        "label": "商品データの取込み",
        "role": "Keepaから商品情報をまとめて取る作業。時間がかかります。",
        "order": 3,
    },
    {
        "id": "ckpt_trigger",
        "label": "目覚まし（続き予約）",
        "role": "「何分後にもう一度起こして続きをやる」ための予約。途中で落ちたときの保険にも使います。",
        "order": 4,
    },
    {
        "id": "expect",
        "label": "起こす予定時刻",
        "role": "目覚ましが鳴るはずの時刻のメモ。遅れすぎたら異常として記録します。",
        "order": 5,
    },
    {
        "id": "overlap",
        "label": "二重起動の拒否",
        "role": "「作業中の札」があるのに目覚ましが鳴ったとき、二重に動かないよう断る動き。",
        "order": 6,
    },
    {
        "id": "liveness",
        "label": "見回り番",
        "role": "15分ごとに「止まっていないか」を見ます。止まっていたら起こし直します。",
        "order": 7,
    },
    {
        "id": "gas_budget",
        "label": "1回の時間上限",
        "role": "1回の作業はだいたい20分で区切るルール。長くなりすぎないための目安です。",
        "order": 8,
    },
]

# --- 現状（実ログ） ---
def _dur_label(start: float, end: float) -> str:
    return fmt_dur(end - start)


BEFORE = {
    "activeSpans": [
        {
            "lane": "gas_run",
            "start": 0,
            "end": T_EXIT,
            "label": f"作業 {_dur_label(0, T_EXIT)}",
            "menu": "作業そのもの",
            "tone": "primary",
            "actualMin": T_EXIT,
        },
        {
            "lane": "inflight",
            "start": 0,
            "end": T_EXIT,
            "label": f"札 {_dur_label(0, T_EXIT)}",
            "menu": "いま作業中の札",
            "tone": "danger",
            "actualMin": T_EXIT,
        },
        {
            "lane": "wash_lock",
            "start": T_TARGET,
            "end": T_EXIT,
            "label": f"鍵 {_dur_label(T_TARGET, T_EXIT)}",
            "menu": "メーカーの鍵",
            "tone": "warn",
            "actualMin": round(T_EXIT - T_TARGET, 2),
        },
        {
            "lane": "keepa_slice",
            "start": T_TARGET,
            "end": T_SLICE_END,
            "label": f"取込 {_dur_label(T_TARGET, T_SLICE_END)}",
            "menu": "商品データの取込み",
            "tone": "accent",
            "actualMin": round(T_SLICE_END - T_TARGET, 2),
        },
        {
            "lane": "ckpt_trigger",
            "start": T_CKPT0,
            "end": T_OVERLAP,
            "label": f"目覚 {_dur_label(T_CKPT0, T_OVERLAP)}",
            "menu": "目覚まし",
            "tone": "trigger",
            "actualMin": round(T_OVERLAP - T_CKPT0, 2),
        },
        {
            "lane": "expect",
            "start": T_CKPT0,
            "end": T_MISSED,
            "label": f"メモ {_dur_label(T_CKPT0, T_MISSED)}",
            "menu": "起こす予定時刻",
            "tone": "danger",
            "actualMin": round(T_MISSED - T_CKPT0, 2),
        },
        # liveness: 点のみ（実行は数秒以内の別GAS）
        {
            "lane": "liveness",
            "start": T_LIVE1,
            "end": T_LIVE1 + 0.08,
            "label": "見回り①",
            "menu": "見回り番",
            "tone": "liveness",
            "actualMin": 0.08,
        },
        {
            "lane": "liveness",
            "start": T_LIVE2,
            "end": T_LIVE2 + 0.08,
            "label": "見回り②",
            "menu": "見回り番",
            "tone": "liveness",
            "actualMin": 0.08,
        },
        {
            "lane": "ckpt_trigger",
            "start": T_EXIT,
            "end": T_FIRE2,
            "label": f"続き予約→発火 {_dur_label(T_EXIT, T_FIRE2)}",
            "menu": "目覚まし（timeup後）",
            "tone": "trigger",
            "actualMin": round(T_FIRE2 - T_EXIT, 2),
        },
        {
            "lane": "gas_run",
            "start": T_FIRE2,
            "end": min(WINDOW_MIN, T_FIRE2 + 0.5),
            "label": "次run開始",
            "menu": "作業そのもの（232733）",
            "tone": "primary",
        },
    ],
    "budgetLine": T_BUDGET,
    "budgetOverrunMin": round(T_EXIT - T_BUDGET, 2),
    "events": [
        {"t": 0, "lane": "gas_run", "label": "作業開始", "kind": "fire"},
        {"t": T_CKPT0, "lane": "ckpt_trigger", "label": "目覚まし予約", "kind": "arm"},
        {"t": T_EXPECT, "lane": "expect", "label": "起こす予定時刻", "kind": "expect"},
        {"t": T_OVERLAP, "lane": "overlap", "label": "二重起動を拒否", "kind": "conflict"},
        {"t": T_BUDGET, "lane": "gas_budget", "label": f"設計上限20分（+{fmt_dur(T_EXIT - T_BUDGET)}超過）", "kind": "budget"},
        {"t": T_MISSED, "lane": "expect", "label": "「遅れた」と誤記録", "kind": "missed"},
        {"t": T_EXIT, "lane": "gas_run", "label": "時間切れで終了", "kind": "exit"},
        {"t": T_RESERVE2, "lane": "ckpt_trigger", "label": "続き予約（timeup後）", "kind": "arm"},
        {"t": T_EXPECT_CONT, "lane": "ckpt_trigger", "label": "発火予定", "kind": "expect"},
        {"t": T_GAS_WALL, "lane": "gas_budget", "label": "GAS30分壁（目安）", "kind": "budget"},
        {"t": T_FIRE2, "lane": "gas_run", "label": "次run fire", "kind": "fire"},
    ],
    "overlaps": [
        {
            "start": T_CKPT0,
            "end": T_OVERLAP,
            "lanes": ["gas_run", "inflight", "wash_lock", "keepa_slice", "ckpt_trigger", "expect"],
            "label": "重複×6",
            "detail": "取込み中なのに目覚ましも予約され、同時に動いている",
        },
        {
            "start": T_OVERLAP,
            "end": T_MISSED,
            "lanes": ["gas_run", "inflight", "wash_lock", "keepa_slice", "expect"],
            "label": "重複×5",
            "detail": "目覚ましは拒否されたのに、予定時刻メモだけ残り続ける",
        },
    ],
    "callouts": [
        {
            "id": "c_ckpt0",
            "t": T_CKPT0,
            "lane": "ckpt_trigger",
            "side": "top",
            "offsetPx": -50,
            "text": "想定: 2分後に別の作業で続き\n実際: 同じ作業が取込みを継続",
        },
        {
            "id": "c_overlap",
            "t": T_OVERLAP,
            "lane": "overlap",
            "side": "bottom",
            "offsetPx": 10,
            "text": "想定: 目覚ましで続き開始\n実際: 「作業中」のため拒否",
        },
        {
            "id": "c_budget",
            "t": T_BUDGET,
            "lane": "keepa_slice",
            "side": "top",
            "offsetPx": 0,
            "text": f"想定: 20分で区切る\n実際: あと{fmt_dur(T_EXIT - T_BUDGET)}走り続ける",
        },
        {
            "id": "c_missed",
            "t": T_MISSED,
            "lane": "expect",
            "side": "bottom",
            "offsetPx": 0,
            "text": "想定: 鳴ったらメモを消す\n実際: 「遅れた」と誤って記録",
        },
        {
            "id": "c_reserve2",
            "t": T_EXIT,
            "lane": "ckpt_trigger",
            "side": "top",
            "offsetPx": 55,
            "text": "ここで本当の続き予約\n2分後→23:27発火（30分超）",
        },
    ],
}

# --- 案1+3 改善後（イメージ） ---
# ※ 取込み時間は実績どおり ~29分。変わるのは目覚ましの張り方だけ
AFTER = {
    "activeSpans": [
        {
            "lane": "gas_run",
            "start": 0,
            "end": T_EXIT,
            "label": f"作業 {_dur_label(0, T_EXIT)}（同じ）",
            "menu": "作業そのもの",
            "tone": "primary",
            "actualMin": T_EXIT,
        },
        {
            "lane": "inflight",
            "start": 0,
            "end": T_EXIT,
            "label": f"札 {_dur_label(0, T_EXIT)}（同じ）",
            "menu": "いま作業中の札",
            "tone": "danger",
            "actualMin": T_EXIT,
        },
        {
            "lane": "wash_lock",
            "start": T_TARGET,
            "end": T_EXIT,
            "label": f"鍵 {_dur_label(T_TARGET, T_EXIT)}（同じ）",
            "menu": "メーカーの鍵",
            "tone": "warn",
            "actualMin": round(T_EXIT - T_TARGET, 2),
        },
        {
            "lane": "keepa_slice",
            "start": T_TARGET,
            "end": T_SLICE_END,
            "label": f"取込 {_dur_label(T_TARGET, T_SLICE_END)}（同じ）",
            "menu": "商品データの取込み",
            "tone": "accent",
            "actualMin": round(T_SLICE_END - T_TARGET, 2),
        },
        {
            "lane": "ckpt_trigger",
            "start": T_EXIT,
            "end": min(WINDOW_MIN, T_FIRE2),
            "label": f"終了後だけ予約 {_dur_label(T_EXIT, T_FIRE2)}",
            "menu": "目覚まし",
            "tone": "trigger",
            "actualMin": round(T_FIRE2 - T_EXIT, 2),
        },
        {
            "lane": "liveness",
            "start": T_LIVE1,
            "end": T_LIVE1 + 0.08,
            "label": "見回り①",
            "menu": "見回り番",
            "tone": "liveness",
            "actualMin": 0.08,
        },
        {
            "lane": "liveness",
            "start": T_LIVE2,
            "end": T_LIVE2 + 0.08,
            "label": "見回り②",
            "menu": "見回り番",
            "tone": "liveness",
            "actualMin": 0.08,
        },
    ],
    "budgetLine": T_BUDGET,
    "budgetOverrunMin": round(T_EXIT - T_BUDGET, 2),
    "events": [
        {"t": 0, "lane": "gas_run", "label": "作業開始", "kind": "fire"},
        {"t": T_BUDGET, "lane": "gas_budget", "label": f"設計上限（実際は+{fmt_dur(T_EXIT - T_BUDGET)}継続）", "kind": "budget"},
        {"t": T_EXIT, "lane": "gas_run", "label": "終了→続き予約", "kind": "exit"},
        {"t": T_EXIT + 0.05, "lane": "ckpt_trigger", "label": "目覚まし予約", "kind": "arm"},
    ],
    "overlaps": [],  # 自己衝突なし
    "callouts": [
        {
            "id": "a_ckpt0",
            "t": T_CKPT0,
            "lane": "ckpt_trigger",
            "side": "top",
            "offsetPx": 0,
            "tone": "ok",
            "text": "改善: 作業中は目覚ましを張らない\n（灰色のまま＝休眠）",
        },
        {
            "id": "a_overlap",
            "t": T_OVERLAP,
            "lane": "overlap",
            "side": "bottom",
            "offsetPx": 0,
            "tone": "ok",
            "text": "改善: 二重起動の拒否が起きない\n（目覚ましがまだ鳴らない）",
        },
        {
            "id": "a_exit",
            "t": T_EXIT,
            "lane": "ckpt_trigger",
            "side": "top",
            "offsetPx": -40,
            "tone": "ok",
            "text": "改善: 作業が終わってから\n続きの目覚ましを1回だけ予約",
        },
    ],
}

HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Keepa ガント比較 — 現状 vs 改善後</title>
<style>
  :root {
    --bg: #0f1419; --panel: #1a2332; --text: #e7ecf3; --muted: #8b9cb3; --grid: #2a3544;
    --sleep: #252b34; --sleep-border: #333b48;
    --primary: #4d9fff; --danger: #ff6b6b; --warn: #ffb020; --accent: #3dd68c;
    --trigger: #7eb8ff; --budget: #a78bfa; --liveness: #c4b5fd; --ok: #3dd68c;
  }
  * { box-sizing: border-box; }
  body { margin: 0; font-family: "Segoe UI", "Hiragino Sans", sans-serif; background: var(--bg); color: var(--text); }
  header { padding: 18px 24px 12px; border-bottom: 1px solid var(--grid); }
  h1 { margin: 0 0 4px; font-size: 1.25rem; }
  .sub { color: var(--muted); font-size: 0.85rem; line-height: 1.5; }
  .layout { display: grid; grid-template-columns: 280px 1fr; }
  aside { padding: 14px 16px; border-right: 1px solid var(--grid); background: var(--panel); font-size: 0.8rem; overflow-y: auto; max-height: 100vh; position: sticky; top: 0; }
  aside h2 { font-size: 0.7rem; text-transform: uppercase; letter-spacing: .05em; color: var(--muted); margin: 16px 0 8px; }
  aside h2:first-child { margin-top: 0; }
  .role { margin-bottom: 10px; padding: 8px 10px; background: #141c28; border-radius: 6px; border-left: 3px solid var(--primary); }
  .role strong { display: block; font-size: 0.78rem; margin-bottom: 3px; color: var(--text); }
  .role span { color: var(--muted); font-size: 0.74rem; line-height: 1.4; }
  .legend-item { display: flex; align-items: center; gap: 8px; margin-bottom: 6px; }
  .swatch { width: 14px; height: 14px; border-radius: 3px; flex-shrink: 0; }
  .swatch-sleep { background: var(--sleep); border: 1px solid var(--sleep-border); }
  .swatch-overlap { background: repeating-linear-gradient(-45deg, rgba(255,80,80,.2), rgba(255,80,80,.2) 4px, rgba(255,107,107,.45) 4px, rgba(255,107,107,.45) 8px); }
  main { padding: 12px 16px 56px; overflow-x: auto; }
  .section-title {
    margin: 8px 0 10px; padding: 10px 12px; border-radius: 8px; font-size: 0.95rem; font-weight: 600;
  }
  .section-title.before { background: #2a1a1a; border-left: 4px solid var(--danger); }
  .section-title.after { background: #14241c; border-left: 4px solid var(--ok); margin-top: 36px; }
  .section-note { color: var(--muted); font-size: 0.78rem; margin: -4px 0 12px 12px; line-height: 1.45; }
  .chart-area { position: relative; min-width: 1050px; margin-bottom: 8px; }
  .time-axis { display: flex; justify-content: space-between; margin-left: 200px; font-size: 0.68rem; color: var(--muted); padding-bottom: 6px; }
  .overlap-columns { position: absolute; left: 200px; right: 0; top: 26px; bottom: 0; pointer-events: none; z-index: 1; }
  .overlap-col {
    position: absolute; top: 0; bottom: 0;
    background: repeating-linear-gradient(-45deg, transparent, transparent 5px, rgba(255,90,90,.12) 5px, rgba(255,90,90,.12) 10px);
    border-left: 2px solid rgba(255,107,107,.5); border-right: 2px solid rgba(255,107,107,.5);
  }
  .overlap-badge {
    position: absolute; top: 2px; left: 50%; transform: translateX(-50%);
    font-size: 0.6rem; font-weight: 600; color: #ff8a8a;
    background: rgba(20,24,32,.92); padding: 2px 6px; border-radius: 4px;
    border: 1px solid rgba(255,107,107,.4); white-space: nowrap;
  }
  .lanes { position: relative; z-index: 2; }
  .lane-row { display: flex; align-items: center; height: 46px; margin-bottom: 2px; }
  .lane-label { width: 200px; flex-shrink: 0; font-size: 0.74rem; padding-right: 10px; text-align: right; color: var(--muted); line-height: 1.2; }
  .lane-track { flex: 1; height: 34px; position: relative; border-radius: 4px; overflow: visible; }
  .bar-sleep { position: absolute; inset: 0; background: var(--sleep); border: 1px solid var(--sleep-border); border-radius: 4px; }
  .bar-active {
    position: absolute; top: 3px; height: 28px; border-radius: 4px;
    display: flex; align-items: center; padding: 0 6px;
    font-size: 0.62rem; font-weight: 500; white-space: nowrap; overflow: hidden;
    cursor: pointer; z-index: 3;
  }
  .bar-active:hover { filter: brightness(1.12); z-index: 8; }
  .tone-primary { background: var(--primary); color: #0a0e14; }
  .tone-danger { background: var(--danger); color: #fff; }
  .tone-warn { background: var(--warn); color: #0a0e14; }
  .tone-accent { background: var(--accent); color: #0a0e14; }
  .tone-trigger { background: var(--trigger); color: #0a0e14; }
  .tone-budget { background: var(--budget); color: #0a0e14; opacity: .85; }
  .tone-liveness { background: var(--liveness); color: #0a0e14; min-width: 4px !important; }
  .marker { position: absolute; top: 0; bottom: 0; width: 2px; z-index: 6; cursor: pointer; }
  .marker::after {
    content: ""; position: absolute; top: -3px; left: -5px;
    width: 11px; height: 11px; border-radius: 50%; border: 2px solid currentColor; background: var(--bg);
  }
  .callout { position: absolute; z-index: 20; max-width: 190px; font-size: 0.66rem; line-height: 1.35; cursor: pointer; transition: opacity .15s; }
  .callout.hidden { opacity: 0; pointer-events: none; visibility: hidden; }
  .callout-pin {
    position: absolute; z-index: 25; width: 12px; height: 12px; border-radius: 50%;
    background: #ff9f43; border: 2px solid #fff; cursor: pointer; transform: translate(-50%, -50%);
    box-shadow: 0 0 0 0 rgba(255,159,67,.5); transition: box-shadow .15s, opacity .15s;
  }
  .callout-pin.ok { background: var(--ok); }
  .callout-pin.off { opacity: .45; box-shadow: none; background: var(--muted); }
  .callout-pin:hover { box-shadow: 0 0 0 3px rgba(255,159,67,.35); }
  .callout-pin.ok:hover { box-shadow: 0 0 0 3px rgba(61,214,140,.35); }
  .callout-toggle-btn {
    display: block; width: 100%; margin: 8px 0 4px; padding: 7px 10px;
    background: #141c28; border: 1px solid var(--grid); border-radius: 6px;
    color: var(--text); font-size: 0.74rem; cursor: pointer;
  }
  .callout-toggle-btn:hover { background: #1a2434; }
  .callout-hint { font-size: 0.68rem; color: var(--muted); line-height: 1.4; margin-bottom: 8px; }
  .callout-box {
    background: #1e2a3a; border: 1px solid #ff9f43; border-radius: 8px;
    padding: 6px 8px; color: #ffe8cc; box-shadow: 0 4px 14px rgba(0,0,0,.35);
  }
  .callout-box.ok { border-color: var(--ok); color: #d8f5e6; }
  .callout-box strong { display: block; font-size: 0.58rem; margin-bottom: 2px; color: #ff9f43; }
  .callout-box.ok strong { color: var(--ok); }
  .callout-tail { width: 0; height: 0; margin: 0 auto; border-left: 7px solid transparent; border-right: 7px solid transparent; }
  .callout.top .callout-tail { border-top: 8px solid #ff9f43; }
  .callout.top .callout-tail.ok { border-top-color: var(--ok); }
  .callout.bottom .callout-tail { border-bottom: 8px solid #ff9f43; order: -1; }
  .callout.bottom .callout-tail.ok { border-bottom-color: var(--ok); }
  .callout.top { transform: translate(-50%, -100%); margin-top: -6px; }
  .callout.bottom { transform: translate(-50%, 4px); display: flex; flex-direction: column; }
  .lane-overlap-hatch {
    position: absolute; top: 3px; height: 28px; border-radius: 4px; z-index: 4; pointer-events: none;
    background: repeating-linear-gradient(90deg, rgba(255,80,80,.25) 0 3px, transparent 3px 6px);
    border: 1px dashed rgba(255,107,107,.65);
  }
  #tooltip {
    position: fixed; display: none; max-width: 300px; padding: 10px 12px;
    background: #243044; border: 1px solid var(--grid); border-radius: 8px;
    font-size: 0.76rem; line-height: 1.45; z-index: 100; pointer-events: none;
  }
  .diff-box {
    margin-top: 20px; padding: 12px 14px; background: #14241c; border-radius: 8px;
    border: 1px solid rgba(61,214,140,.35); font-size: 0.82rem; line-height: 1.55;
  }
  .stats-table { width: 100%; border-collapse: collapse; font-size: 0.72rem; margin-top: 6px; }
  .stats-table th, .stats-table td { border: 1px solid var(--grid); padding: 5px 6px; text-align: left; vertical-align: top; }
  .stats-table th { background: #141c28; color: var(--muted); font-weight: 600; }
  .stats-table td.actual { color: #ffb4b4; font-weight: 600; }
  .stats-table td.design { color: var(--muted); }
  .budget-vline {
    position: absolute; top: 0; bottom: 0; width: 2px; z-index: 5; pointer-events: none;
    border-left: 2px dashed var(--budget);
  }
  .budget-vline-label {
    position: absolute; top: -2px; transform: translateX(-50%);
    font-size: 0.58rem; color: var(--budget); white-space: nowrap;
    background: rgba(15,20,25,.9); padding: 1px 4px; border-radius: 3px;
  }
  .gas-wall-vline {
    position: absolute; top: 0; bottom: 0; width: 2px; z-index: 5; pointer-events: none;
    border-left: 2px dotted rgba(255,180,100,.55);
  }
  .gas-wall-label {
    position: absolute; top: 12px; transform: translateX(-50%);
    font-size: 0.58rem; color: #ffb464; white-space: nowrap;
    background: rgba(15,20,25,.9); padding: 1px 4px; border-radius: 3px;
  }
  .overrun-band {
    position: absolute; top: 0; bottom: 0; z-index: 0;
    background: rgba(167,139,250,.08);
    border-left: 1px solid rgba(167,139,250,.35);
  }
</style>
</head>
<body>
<header>
  <h1>商品データ取込み中の動き — 現状と改善後の比較</h1>
  <p class="sub">実例: __CASE__ / __PARENT__ / 起点 __T0__（35分）<br/>
  薄グレー＝寝ている／色＝動いている／紫破線＝設計上限20分／橙点線＝GAS30分壁／赤斜線＝ぶつかり</p>
</header>
<div class="layout">
<aside>
  <h2>実測時間一覧（①ログ）</h2>
  <div id="stats"></div>
  <h2>各行の役割</h2>
  <div id="roles"></div>
  <h2>色の意味</h2>
  <div class="legend-item"><span class="swatch swatch-sleep"></span>寝ている（待機）</div>
  <div class="legend-item"><span class="swatch tone-primary"></span>作業そのもの</div>
  <div class="legend-item"><span class="swatch tone-accent"></span>商品データの取込み</div>
  <div class="legend-item"><span class="swatch tone-trigger"></span>目覚まし</div>
  <div class="legend-item"><span class="swatch tone-liveness"></span>見回り番（点）</div>
  <div class="legend-item"><span style="color:var(--budget)">┊</span> 紫破線＝20分上限（設計のみ）</div>
  <div class="legend-item"><span class="swatch swatch-overlap"></span>ぶつかり（重複）</div>
  <h2>吹き出し</h2>
  <p class="callout-hint">橙／緑の丸をクリックで表示・非表示。吹き出し本体をクリックしても消せます。</p>
  <button type="button" class="callout-toggle-btn" id="toggle-all-callouts">すべての吹き出しを非表示</button>
</aside>
<main>
  <div class="section-title before">① いま（実ログ）</div>
  <p class="section-note">作業は<strong>約30分</strong>走る。右端（30分超）に timeup 後の「本当の続き予約→発火」が見える。</p>
  <div class="chart-area" id="chart-before"></div>

  <div class="section-title after">② 案1+3（イメージ）</div>
  <p class="section-note"><strong>取込み時間は同じ約30分。</strong>変わるのは目覚ましの張り方だけ（終了後に1回）。ぶつかりが消える。</p>
  <div class="chart-area" id="chart-after"></div>

  <div class="diff-box">
    <strong>比較の要点</strong><br/>
    ・上下とも作業・取込みの<strong>実測時間は同じ（約30分 / 取込み約29分）</strong><br/>
    ・20分上限はどちらも「設計上の線」で、実際の停止は timeup（約30分）<br/>
    ・上だけ赤斜線（重複）と橙吹き出しあり → 下では解消<br/>
    ・30分超の続き予約（23:25→23:27発火）は上段の右端
  </div>
</main>
</div>
<div id="tooltip"></div>
<script>
const DATA = __JSON__;

function fmtMin(m) {
  const s = Math.round(m * 60);
  return Math.floor(s/60) + ":" + String(s%60).padStart(2,"0");
}
function pct(m, W) { return (m / W) * 100; }
function showTip(e, html) {
  const t = document.getElementById("tooltip");
  t.innerHTML = html; t.style.display = "block";
  t.style.left = (e.clientX + 14) + "px"; t.style.top = (e.clientY + 14) + "px";
}
function hideTip() { document.getElementById("tooltip").style.display = "none"; }

function renderChart(elId, pack) {
  const W = DATA.windowMin;
  const lanes = [...DATA.lanes].sort((a,b) => a.order - b.order);
  const laneH = 48;
  const totalH = lanes.length * laneH + 28;
  const ticks = [];
  for (let m = 0; m <= W; m += 5) ticks.push(m);
  let html = `<div class="time-axis">${ticks.map(m => `<span>${fmtMin(m)}</span>`).join("")}</div>`;
  html += `<div class="overlap-columns" style="height:${totalH}px">`;
  if (pack.budgetLine != null) {
    const bl = pct(pack.budgetLine, W);
    html += `<div class="budget-vline" style="left:${bl}%"><span class="budget-vline-label">設計上限20分</span></div>`;
    if (pack.budgetOverrunMin > 0) {
      html += `<div class="overrun-band" style="left:${bl}%;width:${pct(pack.budgetOverrunMin,W)}%"></div>`;
    }
  }
  if (DATA.gasWallMin != null) {
    const gw = pct(DATA.gasWallMin, W);
    html += `<div class="gas-wall-vline" style="left:${gw}%"><span class="gas-wall-label">GAS30分</span></div>`;
  }
  for (const o of (pack.overlaps || [])) {
    html += `<div class="overlap-col" style="left:${pct(o.start,W)}%;width:${pct(o.end-o.start,W)}%">
      <span class="overlap-badge">${o.label}</span></div>`;
  }
  html += `</div><div class="lanes">`;
  for (const lane of lanes) {
    html += `<div class="lane-row"><div class="lane-label">${lane.label}</div><div class="lane-track">`;
    html += `<div class="bar-sleep" title="寝ている"></div>`;
    for (const s of (pack.activeSpans || []).filter(x => x.lane === lane.id)) {
      if (s.end <= s.start) continue;
      const left = pct(s.start, W), width = Math.max(pct(s.end - s.start, W), 0.35);
      html += `<div class="bar-active tone-${s.tone}" style="left:${left}%;width:${width}%"
        data-tip="<strong>${s.menu}</strong><br/>${s.label}<br/>${fmtMin(s.start)}–${fmtMin(s.end)}${s.actualMin != null ? '<br/>実測:' + s.actualMin + '分' : ''}">${s.label}</div>`;
    }
    for (const o of (pack.overlaps || [])) {
      if ((o.lanes || []).indexOf(lane.id) < 0) continue;
      html += `<div class="lane-overlap-hatch" style="left:${pct(o.start,W)}%;width:${pct(o.end-o.start,W)}%"></div>`;
    }
    for (const e of (pack.events || []).filter(x => x.lane === lane.id)) {
      const col = e.kind === "conflict" ? "var(--danger)" : e.kind === "fire" || e.kind === "exit" ? "var(--primary)" : "var(--muted)";
      html += `<div class="marker" style="left:${pct(e.t,W)}%;color:${col}" data-tip="<strong>${e.label}</strong><br/>T0+${fmtMin(e.t)}"></div>`;
    }
    (pack.callouts || []).filter(x => x.lane === lane.id).forEach((c, i) => {
      const top = c.side === "top" ? "-8px" : "36px";
      const off = c.offsetPx || 0;
      const ok = c.tone === "ok";
      const cid = `${elId}-${c.id || ("c" + i)}`;
      html += `<div class="callout-pin ${ok ? "ok" : ""}" data-callout-id="${cid}"
        style="left:calc(${pct(c.t,W)}% + ${off}px);top:17px" title="クリックで吹き出し表示/非表示"></div>`;
      html += `<div class="callout ${c.side}" data-id="${cid}" style="left:calc(${pct(c.t,W)}% + ${off}px);top:${top}">
        <div class="callout-box ${ok ? "ok" : ""}"><strong>${ok ? "改善" : "想定 ≠ 実際"}</strong>${c.text.replace(/\n/g,"<br/>")}</div>
        <div class="callout-tail ${ok ? "ok" : ""}"></div></div>`;
    });
    html += `</div></div>`;
  }
  html += `</div>`;
  const el = document.getElementById(elId);
  el.innerHTML = html;
  el.querySelectorAll("[data-tip]").forEach(node => {
    node.addEventListener("mousemove", ev => showTip(ev, node.dataset.tip));
    node.addEventListener("mouseleave", hideTip);
  });
  bindCallouts(el);
}

function toggleCallout(el, id) {
  const box = el.querySelector(`.callout[data-id="${id}"]`);
  const pin = el.querySelector(`.callout-pin[data-callout-id="${id}"]`);
  if (!box) return;
  const hidden = box.classList.toggle("hidden");
  if (pin) pin.classList.toggle("off", hidden);
}

function bindCallouts(el) {
  el.querySelectorAll(".callout-pin").forEach(pin => {
    pin.addEventListener("click", e => {
      e.stopPropagation();
      toggleCallout(el, pin.dataset.calloutId);
    });
  });
  el.querySelectorAll(".callout[data-id]").forEach(box => {
    box.addEventListener("click", e => {
      e.stopPropagation();
      const id = box.dataset.id;
      box.classList.add("hidden");
      const pin = el.querySelector(`.callout-pin[data-callout-id="${id}"]`);
      if (pin) pin.classList.add("off");
    });
  });
}

let allCalloutsVisible = true;
document.getElementById("toggle-all-callouts").addEventListener("click", () => {
  allCalloutsVisible = !allCalloutsVisible;
  document.querySelectorAll(".callout").forEach(c => c.classList.toggle("hidden", !allCalloutsVisible));
  document.querySelectorAll(".callout-pin").forEach(p => p.classList.toggle("off", !allCalloutsVisible));
  document.getElementById("toggle-all-callouts").textContent =
    allCalloutsVisible ? "すべての吹き出しを非表示" : "すべての吹き出しを表示";
});

document.getElementById("stats").innerHTML = `
<table class="stats-table">
<tr><th>項目</th><th>設計</th><th>実測</th></tr>
${DATA.actualStats.map(r => `<tr>
  <td>${r.name}</td>
  <td class="design">${r.design}</td>
  <td class="actual">${r.actual}<br/><span style="font-weight:400;color:var(--muted)">${r.clock}</span></td>
</tr>`).join("")}
</table>`;

document.getElementById("roles").innerHTML = DATA.lanes
  .slice().sort((a,b) => a.order - b.order)
  .map(l => `<div class="role"><strong>${l.label}</strong><span>${l.role}</span></div>`).join("");

renderChart("chart-before", DATA.before);
renderChart("chart-after", DATA.after);
</script>
</body>
</html>
"""


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    payload = {
        "t0": T0.strftime("%Y-%m-%d %H:%M:%S"),
        "case": CASE,
        "parent": PARENT,
        "windowMin": WINDOW_MIN,
        "gasWallMin": T_GAS_WALL,
        "lanes": LANES,
        "actualStats": ACTUAL_STATS,
        "before": BEFORE,
        "after": AFTER,
    }
    JSON_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    html = (
        HTML_TEMPLATE.replace("__JSON__", json.dumps(payload, ensure_ascii=False))
        .replace("__CASE__", CASE)
        .replace("__PARENT__", PARENT)
        .replace("__T0__", T0.strftime("%Y-%m-%d %H:%M:%S"))
    )
    HTML_PATH.write_text(html, encoding="utf-8")
    print(f"wrote {HTML_PATH}")


if __name__ == "__main__":
    main()
