# -*- coding: utf-8 -*-
"""仮想: テーマ概要JSON → OpenAI画像（楽天ULなし）。概要が通るかのスモーク用。"""
from __future__ import annotations

import argparse
import json
import logging
import time
from pathlib import Path
from typing import Any, Dict, List, Optional

LOG = logging.getLogger("set_main_image.dry_sub_overview_openai")

DEFAULT_JAN = "4906283045317"
DEFAULT_TRUTH_DIRS = (
    Path(r"G:\マイドライブ\05.画像生成（セットMAIN）\01.amazon白抜きベース"),
    Path(r"G:\My Drive\05.画像生成（セットMAIN）\01.amazon白抜きベース"),
)
DEFAULT_OVERVIEW = Path(__file__).resolve().parent / "_out" / "dry_sub_theme_gate_4906283045317_v2b.json"


def _truth_search_dirs(base: Path):
    """直下と、その一つ下の「処理済み」。B-④が選ぶ正本は処理済みに入る。"""
    if not base.is_dir():
        return []
    out = [base]
    proc = base / "処理済み"
    if proc.is_dir():
        out.append(proc)
    return out


def _iter_truth_images(base: Path):
    for d in _truth_search_dirs(base):
        for p in sorted(d.iterdir()):
            if p.is_file() and p.suffix.lower() in (".png", ".jpg", ".jpeg", ".webp"):
                yield p


def _shared_truth_for_jan(jan: str):
    """同じJANが複数親にあるとき、最後に保存された正本を使う。"""
    j = str(jan or "").strip()
    if not j:
        return None
    try:
        from b5_parent_shared import load_parent_shared
    except Exception:
        return None
    doc = load_parent_shared()
    best = None
    best_at = ""
    for row in (doc.get("byParentSku") or {}).values():
        if not isinstance(row, dict):
            continue
        if str(row.get("jan") or "").strip() != j:
            continue
        rel = str(row.get("packageTruthRel") or "").strip().replace("/", "\\")
        if not rel:
            continue
        updated = str(row.get("updatedAt") or "")
        for base in DEFAULT_TRUTH_DIRS:
            path = base / rel
            if path.is_file() and updated >= best_at:
                best = path
                best_at = updated
    return best


def _exact_named_truth(hint: str) -> Optional[Path]:
    """渡されたファイル名と完全一致するものだけ。oya が oya2 に吸われない。"""
    name = Path(str(hint or "").strip()).name
    if not name or name in (".", ".."):
        return None
    hits = []
    for base in DEFAULT_TRUTH_DIRS:
        for p in _iter_truth_images(base):
            if p.name == name:
                hits.append(p)
    if not hits:
        return None
    processed = [p for p in hits if p.parent.name == "処理済み"]
    if len(processed) == 1:
        return processed[0]
    return hits[0]


def _find_truth(jan: str, hint: str = "") -> Path:
    """正本画像を探す。ファイル名の完全一致 → 保存済み正本 → ファイル名のJAN → ヒント。"""
    j = str(jan or "").strip()
    h = str(hint or "").strip()
    exact = _exact_named_truth(h)
    if exact is not None:
        return exact
    # ファイル名にJANが無い運用向け（食べるおだし等）。完全一致が無いときだけ。
    if not h:
        h = {
            "4906283045119": "かつお",
            "4906283045317": "まぐろ",
            "4906283047410": "ぶり",
        }.get(j, "")
    shared = _shared_truth_for_jan(j)
    if shared is not None:
        return shared
    for d in DEFAULT_TRUTH_DIRS:
        if j:
            for p in _iter_truth_images(d):
                if j in p.name:
                    return p
        if h:
            for p in _iter_truth_images(d):
                if h in p.name:
                    return p
    raise SystemExit(f"正本が見つかりません jan={jan} hint={hint or h} dirs={DEFAULT_TRUTH_DIRS}")


def _build_prompt(product_name: str, theme_id: int, theme_name: str, overview: str) -> str:
    return f"""あなたは食品ECのサブ画像を1枚作る。出口は楽天用の正方形。

【商品】{product_name}
【主テーマ】T{theme_id:02d} {theme_name}
【テーマ概要（最優先・この内容を画像化する）】
{overview}

【パッケージ方針】
- 商品パッケージは必須ではない。概要がシーン・人・料理主役ならパッケージ無し／ごく小さくてよい。
- パッケージを載せる場合のみ、参照の正本画像の色・柄・ラベル文字を改変しない。

【守ること】
- 概要とテーマ名に合わない構図にしない（他テーマの言い換え禁止）。
- 材料に無い数字・ランキング・効能・価格を作らない。
- 文字は最小。宣伝スマイルの偽お客様顔は禁止。
- 写真っぽい商用品質。正方形。

【入力画像】IMAGE_PACKAGE_TRUTH（正本。載せるときだけ厳密に使え）。
"""


def _write_gallery(out_dir: Path, rows: List[Dict[str, Any]], meta: Dict[str, Any]) -> Path:
    cards = []
    for r in rows:
        img = r.get("relPath") or ""
        err = r.get("error") or ""
        body = (
            f"<img src='{img}' alt='T{r['themeId']}'/>"
            if r.get("ok") and img
            else f"<pre class='err'>{err}</pre>"
        )
        cards.append(
            "<article class='card'>"
            f"<div class='meta'><b>T{r['themeId']}</b> {r.get('name','')} "
            f"<span class='badge'>{'OK' if r.get('ok') else 'FAIL'}</span></div>"
            f"<p class='ov'>{r.get('overview','')}</p>"
            f"{body}"
            "</article>"
        )
    html = f"""<!DOCTYPE html>
<html lang="ja"><head><meta charset="UTF-8"/>
<title>overview→OpenAI test</title>
<style>
body{{font-family:Segoe UI,Hiragino Sans,sans-serif;margin:0;background:#111;color:#eee}}
header{{padding:16px 20px;background:#222}}
main{{display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:12px;padding:16px}}
.card{{background:#1c1c1c;border:1px solid #333;border-radius:10px;padding:10px}}
.card img{{width:100%;height:auto;border-radius:6px;margin-top:8px}}
.ov{{font-size:12px;line-height:1.5;color:#ccc}}
.badge{{font-size:11px;padding:2px 6px;border-radius:999px;background:#2e7d32}}
.err{{color:#ff8a80;white-space:pre-wrap;font-size:11px}}
</style></head><body>
<header>
<h1>概要→OpenAI 画像テスト</h1>
<p>JAN {meta.get('jan')}／{meta.get('productName')}／quality={meta.get('quality')}／model={meta.get('model')}／n={len(rows)}</p>
<p>truth={meta.get('truth')}</p>
</header>
<main>{''.join(cards)}</main>
</body></html>
"""
    path = out_dir / "index.html"
    path.write_text(html, encoding="utf-8")
    return path


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    ap = argparse.ArgumentParser(description="概要JSON→OpenAI画像（ULなし）")
    ap.add_argument("--overview-json", type=Path, default=DEFAULT_OVERVIEW)
    ap.add_argument("--jan", default=DEFAULT_JAN)
    ap.add_argument("--truth-hint", default="まぐろ")
    ap.add_argument("--truth", type=Path, default=None)
    ap.add_argument("--quality", default="medium", choices=("low", "medium", "high"))
    ap.add_argument("--limit", type=int, default=0, help="先頭N件のみ（0=全部）")
    ap.add_argument("--theme-ids", default="", help="例: 1,3,11（空=JSONの全hit）")
    ap.add_argument(
        "--out-dir",
        type=Path,
        default=None,
        help="省略時 tools/set_main_image/_out/overview_openai_{jan}_{ts}",
    )
    args = ap.parse_args()

    data = json.loads(Path(args.overview_json).read_text(encoding="utf-8"))
    hits = list(data.get("hitOverviews") or [])
    if not hits:
        raise SystemExit("hitOverviews が空です")
    want: Optional[set] = None
    if str(args.theme_ids or "").strip():
        want = {int(x) for x in str(args.theme_ids).split(",") if x.strip().isdigit()}
        hits = [h for h in hits if int(h["themeId"]) in want]
    if args.limit and args.limit > 0:
        hits = hits[: int(args.limit)]

    mat = data.get("materials") or {}
    product_name = str(mat.get("name") or data.get("jan") or "product")
    truth = Path(args.truth) if args.truth else _find_truth(str(args.jan), str(args.truth_hint))
    if not truth.is_file():
        raise SystemExit(f"正本ファイルなし: {truth}")

    ts = time.strftime("%Y%m%d_%H%M%S")
    out_dir = args.out_dir or (
        Path(__file__).resolve().parent / "_out" / f"overview_openai_{args.jan}_{ts}"
    )
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "prompts").mkdir(exist_ok=True)

    from openai_image import generate_with_references, resolve_openai_image_model

    model = resolve_openai_image_model(None)
    rows: List[Dict[str, Any]] = []
    t0 = time.time()
    for i, h in enumerate(hits):
        tid = int(h["themeId"])
        name = str(h.get("name") or "")
        overview = str(h.get("overview") or "").strip()
        stem = f"T{tid:02d}_{i+1:02d}"
        prompt = _build_prompt(product_name, tid, name, overview)
        (out_dir / "prompts" / f"{stem}.txt").write_text(prompt, encoding="utf-8")
        rec: Dict[str, Any] = {
            "themeId": tid,
            "name": name,
            "overview": overview,
            "ok": False,
            "path": "",
            "relPath": "",
            "error": "",
            "ms": 0,
        }
        LOG.info("generate %s/%s %s …", i + 1, len(hits), stem)
        g0 = time.time()
        try:
            data_b, used = generate_with_references(
                prompt=prompt,
                image_paths=[truth],
                model=model,
                size="1024x1024",
                quality=str(args.quality),
            )
            out_path = out_dir / f"{stem}.jpg"
            out_path.write_bytes(data_b)
            rec["ok"] = True
            rec["path"] = str(out_path)
            rec["relPath"] = out_path.name
            rec["model"] = used
            rec["ms"] = int((time.time() - g0) * 1000)
            LOG.info("OK %s ms=%s model=%s", stem, rec["ms"], used)
        except Exception as e:
            rec["error"] = str(e)
            rec["ms"] = int((time.time() - g0) * 1000)
            LOG.exception("FAIL %s: %s", stem, e)
        rows.append(rec)

    meta = {
        "jan": args.jan,
        "productName": product_name,
        "truth": str(truth),
        "overviewJson": str(args.overview_json),
        "quality": args.quality,
        "model": model,
        "okN": sum(1 for r in rows if r["ok"]),
        "failN": sum(1 for r in rows if not r["ok"]),
        "totalMs": int((time.time() - t0) * 1000),
        "rows": rows,
        "note": "楽天ULなし。概要→OpenAIの仮想テスト。",
    }
    (out_dir / "summary.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    gallery = _write_gallery(out_dir, rows, meta)
    print(json.dumps({k: meta[k] for k in ("jan", "okN", "failN", "totalMs", "truth")}, ensure_ascii=False))
    print("OUT", out_dir)
    print("GALLERY", gallery)
    return 0 if meta["failN"] == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
