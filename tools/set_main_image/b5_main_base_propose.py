# -*- coding: utf-8 -*-
"""
B-⑤ 案F — 材料から最大4案（画像なし）。

GAS／Cursor 指示の第1段。売れる順の試行点数付き JSON を出す。
画像生成は b5_main_base_compose.py（P2 --role rakuten_main_base）。仮想検証は b5_virt_abc.py。
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List

from b5_virt_abc import (
    JAN as DEFAULT_JAN,
    _compact_facts_from_sheets,
    propose_from_materials,
)
from work_paths import DEFAULT_WORK_ROOT, SUB_TEST_OUT

LOG = logging.getLogger("set_main_image.b5_main_base_propose")


def run_propose(jan: str) -> Dict[str, Any]:
    mat = _compact_facts_from_sheets(jan)
    materials = mat.get("materials") or {}
    if not materials:
        materials = {
            "name": "",
            "maker": "",
            "feats": "",
            "kw": "",
            "yahoo": "",
            "rakuten": "",
            "bullets": [],
            "desc": "",
        }
    proposals = propose_from_materials(materials)
    return {
        "jan": jan,
        "error": mat.get("error") or "",
        "materials": materials,
        "proposals": proposals,
        "note": "案F試行。catchは6〜18字・【見出し】優先。語の途中切れ禁止。画像は別コマンド。",
    }


def main(argv: List[str] | None = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    ap = argparse.ArgumentParser(description="B-⑤ 材料→最大4案（画像なし）")
    ap.add_argument("--jan", default=DEFAULT_JAN, help="対象JAN（既定=かつお検証JAN）")
    ap.add_argument(
        "--out",
        default="",
        help="出力JSON。空なら 00.テスト出力/b5_main_base_virt/{jan}_{ts}_propose/b5_proposals.json",
    )
    args = ap.parse_args(argv)
    jan = "".join(ch for ch in str(args.jan or "") if ch.isdigit()) or str(args.jan or "").strip()
    if not jan:
        raise SystemExit("--jan が空")

    payload = run_propose(jan)
    if str(args.out or "").strip():
        out_path = Path(args.out)
        out_path.parent.mkdir(parents=True, exist_ok=True)
    else:
        out_root = (
            Path(DEFAULT_WORK_ROOT)
            / SUB_TEST_OUT
            / "b5_main_base_virt"
            / f"{jan}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_propose"
        )
        out_root.mkdir(parents=True, exist_ok=True)
        out_path = out_root / "b5_proposals.json"

    out_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    LOG.info("wrote %s error=%s n=%s", out_path, payload.get("error"), len(payload.get("proposals") or []))
    for row in payload.get("proposals") or []:
        LOG.info(
            "  #%s %s %s score=%s catch=%s",
            row.get("sellRank"),
            row.get("themeId"),
            row.get("themeName"),
            row.get("sellScore"),
            row.get("catch"),
        )
    print(json.dumps({"ok": True, "out": str(out_path), "n": len(payload.get("proposals") or [])}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
