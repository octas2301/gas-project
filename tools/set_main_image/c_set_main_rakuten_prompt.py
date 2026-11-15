# -*- coding: utf-8 -*-
"""CセットMAIN（楽天）Cursor指示文 — B-④ 02コピー後の弱い統合用。

GAS `menuCSetMainRakutenCursorPrompt_` と同趣旨。ローカル 02 マップを読んで
compose_set_main --checked-only の指示文を返す（GASは呼ばない）。

既定: focus_parent_sku があるとき **今の親だけ**（--parent-sku）。
全親一括が必要なときだけ scope="all"。
"""
from __future__ import annotations

import json
import logging
import subprocess
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

from b5_set_base_map import MAP_SUFFIX, normalize_text_color
from text_color import resolve_text_color
from work_paths import (
    DEFAULT_MASTER_CSV,
    DEFAULT_RAKUTEN_UPLOAD_DIR,
    DEFAULT_WORK_ROOT,
    SUB_RAKUTEN,
    default_work_root,
)

LOG = logging.getLogger("set_main_image.c_set_main_rakuten_prompt")
TOOL_DIR = Path(__file__).resolve().parent

_PALETTE_JA = {
    "1": "えんじ",
    "2": "青",
    "3": "黒",
    "4": "緑",
    "5": "茶",
}


def _by_child_n(data: Dict[str, Any]) -> int:
    by = data.get("byChildSku") or data.get("byChild") or {}
    return len(by) if isinstance(by, dict) else 0


def _scan_maps(rakuten_dir: Path, work_root: Optional[Path] = None) -> List[Dict[str, Any]]:
    from b5_set_base_map import resolve_text_color_for_parent

    rows: List[Dict[str, Any]] = []
    if not rakuten_dir.is_dir():
        return rows
    root = work_root or rakuten_dir.parent
    for p in sorted(rakuten_dir.glob(f"*{MAP_SUFFIX}")):
        try:
            data = json.loads(p.read_text(encoding="utf-8-sig"))
        except (OSError, json.JSONDecodeError):
            continue
        if not isinstance(data, dict):
            continue
        sku = str(data.get("parentSku") or p.name.replace(MAP_SUFFIX, "")).strip()
        tc, src = resolve_text_color_for_parent(
            sku,
            work_root=root,
            rakuten_dir=rakuten_dir,
            default="1",
            promote_map_to_shared=True,
        )
        adopted = data.get("adopted") or {}
        rows.append(
            {
                "parentSku": sku,
                "mapFile": p.name,
                "textColor": tc or "",
                "textColorSource": src,
                "textColorNameJa": _PALETTE_JA.get(tc, "") if tc else "",
                "adoptedN": len(adopted) if isinstance(adopted, dict) else 0,
                "byChildN": _by_child_n(data),
                "composeDir": str(
                    (data.get("extra") or {}).get("composeDir")
                    or data.get("composeDir")
                    or ""
                ),
            }
        )
    return rows


def build_c_set_main_rakuten_prompt(
    *,
    work_root: Optional[Path] = None,
    focus_parent_sku: str = "",
    fallback_text_color: str = "1",
    master_csv: Optional[Path] = None,
    out_dir: Optional[Path] = None,
    scope: str = "focus",
) -> Dict[str, Any]:
    """指示文とサマリ。ok=True なら prompt を Agent／人間がコピーして実行。

    scope:
      - focus（既定）: focus_parent_sku があればその親だけ
      - all: 02直下の全マップ親を一括（従来）
    """
    root = Path(work_root) if work_root else default_work_root()
    rak = root / SUB_RAKUTEN
    maps_all = _scan_maps(rak, work_root=root)
    focus = str(focus_parent_sku or "").strip()
    scope_n = str(scope or "focus").strip().lower()
    if scope_n not in ("focus", "all"):
        scope_n = "focus"

    use_focus = bool(focus) and scope_n == "focus"
    maps = (
        [m for m in maps_all if m.get("parentSku") == focus]
        if use_focus
        else list(maps_all)
    )

    fb = resolve_text_color(fallback_text_color or "1")
    fb_id = str(fb.get("id") or "1")
    if fb_id == "custom":
        fb_id = "1"
        fb = resolve_text_color("1")

    with_color = [m for m in maps if m.get("textColor")]
    # フォールバック: 対象マップの最頻色（focus ならその親の色が優先）
    counts: Dict[str, int] = {}
    for m in with_color:
        cid = str(m["textColor"])
        counts[cid] = counts.get(cid, 0) + 1
    if counts:
        fb_id = max(counts.items(), key=lambda x: x[1])[0]
        fb = resolve_text_color(fb_id)
    elif use_focus and fallback_text_color:
        fb = resolve_text_color(fallback_text_color)
        fb_id = str(fb.get("id") or "1")
        if fb_id == "custom":
            fb_id = "1"
            fb = resolve_text_color("1")

    summary_lines: List[str] = []
    if use_focus:
        summary_lines.append(
            f"対象: **今の親だけ** {focus}（弱い統合・02直後）／02フォルダ={rak}"
        )
        if not maps:
            summary_lines.append(
                "警告: 02直下にこの親の set_base_map.json がありません。"
                "先に「採用分を02へコピー」するか、処理済みから戻してください。"
            )
        else:
            m0 = maps[0]
            summary_lines.append(
                f"- {m0['parentSku']} 採用{m0['adoptedN']} 子割当{m0['byChildN']} "
                f"textColor={m0['textColor'] or 'なし'} {m0.get('textColorNameJa') or ''}"
            )
        if len(maps_all) > 1:
            others = [m["parentSku"] for m in maps_all if m["parentSku"] != focus]
            summary_lines.append(
                f"参考: 02直下に他親マップ {len(others)} 件あり（今回のコマンド対象外）: "
                + ", ".join(others[:6])
                + ("…" if len(others) > 6 else "")
            )
    else:
        summary_lines.append(
            f"対象: 02直下の全親一括（scope=all）／マップ{len(maps)}件"
            f"（textColorあり {len(with_color)}）／フォルダ={rak}"
        )
        if focus:
            summary_lines.append(f"直前の親SKU（このレビュー）: {focus}")
        for m in maps[:12]:
            col = (
                f"textColor={m['textColor']} {m.get('textColorNameJa') or ''}"
                if m.get("textColor")
                else "マップ色なし→フォールバック"
            )
            mark = " ★今の親" if focus and m["parentSku"] == focus else ""
            summary_lines.append(
                f"- {m['parentSku']} 採用{m['adoptedN']} 子割当{m['byChildN']} {col}{mark}"
            )
        if len(maps) > 12:
            summary_lines.append(f"- …他 {len(maps) - 12} 件")

    master = Path(master_csv) if master_csv else DEFAULT_MASTER_CSV
    dest = Path(out_dir) if out_dir else DEFAULT_RAKUTEN_UPLOAD_DIR
    work = root if work_root else DEFAULT_WORK_ROOT

    parent_flag = f' --parent-sku "{focus}"' if use_focus else ""
    if use_focus:
        head = (
            f"【セットMAIN量産・楽天】親SKU {focus} のみ（このレビューの02直後）。\n"
            "出品CKレ点の**この親の子SKU**について、Canvaグリフで金丸文字を合成してください。\n"
            "N=1含む。人が色やパスを手打ちしない。他親は対象外。\n"
        )
    else:
        head = (
            "【セットMAIN量産・楽天】出品CKレ点の子SKUについて、Canvaグリフで金丸文字を合成してください。\n"
            "複数品番・複数親を一括（--checked-only）。N=1含む。人が色やパスを手打ちしない。\n"
        )

    prompt = (
        head
        + "文字色: 親ごと共有ストア→02 set_base_map の順（MAIN／サブどちらが先でも可）。\n"
        f"共有・マップに無い親だけフォールバック: {fb_id} {fb.get('nameJa')}（{fb.get('hex')}）\n"
        "土台: マップ byChildSku → {親SKU}_単体__{adoptKey}.png。マップ無しは従来の親1枚。\n"
        + "\n".join(summary_lines)
        + "\n"
        "前提: マスタCSVが古ければ tools/c1_hpc_packaged/c1_fetch_inputs.py で再取得。"
        " または Sheets 直読で更新してから実行。\n"
        f"作業ルート: {work}\n"
        f"出力先（必須）: {dest}\n\n"
        "実行コマンド:\n"
        f'cd "{TOOL_DIR}"\n'
        "python compose_set_main.py"
        f' --master-csv "{master}"'
        " --checked-only"
        f"{parent_flag}"
        " --malls rakuten"
        " --rakuten-engine layered"
        f" --text-color {fb_id}"
        f' --work-root "{work}"'
        f' --out-dir "{dest}"\n'
    )

    return {
        "ok": True,
        "prompt": prompt,
        "fallbackTextColor": fb_id,
        "fallbackNameJa": fb.get("nameJa"),
        "mapN": len(maps),
        "mapWithColorN": len(with_color),
        "mapAllN": len(maps_all),
        "focusParentSku": focus,
        "scope": "focus" if use_focus else "all",
        "summaryLines": summary_lines,
        "maps": maps,
        "commandParts": {
            "masterCsv": str(master),
            "workRoot": str(work),
            "outDir": str(dest),
            "textColor": fb_id,
            "parentSku": focus if use_focus else "",
            "checkedOnly": True,
        },
        "note": (
            "弱い統合: 今の親だけの指示が既定。任意でレビューUIから実行可。"
            "全親一括は scope=all。"
        ),
    }


def run_c_set_main_rakuten(
    *,
    work_root: Optional[Path] = None,
    focus_parent_sku: str = "",
    fallback_text_color: str = "1",
    master_csv: Optional[Path] = None,
    out_dir: Optional[Path] = None,
    scope: str = "focus",
    refresh_master_from_sheets: bool = True,
) -> Dict[str, Any]:
    """弱い統合の任意実行: compose_set_main をローカルで起動（GASなし）。"""
    built = build_c_set_main_rakuten_prompt(
        work_root=work_root,
        focus_parent_sku=focus_parent_sku,
        fallback_text_color=fallback_text_color,
        master_csv=master_csv,
        out_dir=out_dir,
        scope=scope,
    )
    parts = built.get("commandParts") or {}
    master = Path(parts.get("masterCsv") or DEFAULT_MASTER_CSV)
    work = Path(parts.get("workRoot") or DEFAULT_WORK_ROOT)
    dest = Path(parts.get("outDir") or DEFAULT_RAKUTEN_UPLOAD_DIR)
    tc = str(parts.get("textColor") or "1")
    parent = str(parts.get("parentSku") or "").strip()

    if built.get("scope") == "focus" and not parent:
        return {
            "ok": False,
            "message": "親SKUが空です（レビューの parentSku を確認）",
            "promptInfo": built,
        }
    if built.get("scope") == "focus" and int(built.get("mapN") or 0) < 1:
        return {
            "ok": False,
            "message": (
                f"02直下に {parent} のマップがありません。"
                "先に「採用分を02へコピー」してください。"
            ),
            "promptInfo": built,
        }

    refresh_note = ""
    if refresh_master_from_sheets:
        try:
            import csv

            from sheets_master import count_true_in_column, fetch_master_rows

            rows, info = fetch_master_rows()
            master.parent.mkdir(parents=True, exist_ok=True)
            with master.open("w", encoding="utf-8-sig", newline="") as f:
                csv.writer(f).writerows(rows)
            ck = count_true_in_column(rows)
            refresh_note = f"master refreshed rows={info.get('rowCount')} trueCk={ck}"
            LOG.info("%s -> %s", refresh_note, master)
        except Exception as e:
            refresh_note = f"master refresh skipped: {e}"
            LOG.warning(refresh_note)

    cmd = [
        sys.executable,
        str(TOOL_DIR / "compose_set_main.py"),
        "--master-csv",
        str(master),
        "--checked-only",
        "--malls",
        "rakuten",
        "--rakuten-engine",
        "layered",
        "--text-color",
        tc,
        "--work-root",
        str(work),
        "--out-dir",
        str(dest),
    ]
    if parent:
        cmd.extend(["--parent-sku", parent])

    LOG.info("run C set main: %s", " ".join(cmd))
    try:
        proc = subprocess.run(
            cmd,
            cwd=str(TOOL_DIR),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=60 * 30,
        )
    except subprocess.TimeoutExpired:
        return {
            "ok": False,
            "message": "C量産がタイムアウト（30分）しました",
            "promptInfo": built,
            "refreshNote": refresh_note,
            "cmd": cmd,
        }
    except Exception as e:
        return {
            "ok": False,
            "message": f"C量産起動失敗: {e}",
            "promptInfo": built,
            "refreshNote": refresh_note,
            "cmd": cmd,
        }

    tail = ((proc.stdout or "") + "\n" + (proc.stderr or "")).strip().splitlines()
    tail = tail[-40:] if len(tail) > 40 else tail
    ok = proc.returncode == 0
    return {
        "ok": ok,
        "message": (
            f"C量産完了 parent={parent or '(all)'} exit={proc.returncode}"
            if ok
            else f"C量産失敗 exit={proc.returncode}"
        ),
        "exitCode": proc.returncode,
        "logTail": tail,
        "promptInfo": built,
        "refreshNote": refresh_note,
        "outDir": str(dest),
        "cmd": cmd,
    }
