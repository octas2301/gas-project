# -*- coding: utf-8 -*-
"""①一桁 ②pair ③digits の固定パターンを3枚書き出す。"""
from __future__ import annotations

import json
from pathlib import Path

from PIL import Image

from glyph_assets import compose_badge_with_glyphs, default_glyph_dirs
from work_paths import default_work_root

RULES = Path(__file__).resolve().parent / "layout_rules.json"


def main() -> None:
    root = default_work_root()
    base_path = next(
        (next(p for p in root.iterdir() if p.is_dir() and p.name.startswith("02"))).glob("*.jpg")
    )
    out_dir = root / "00.テスト出力"
    dirs = default_glyph_dirs(root)
    rules = json.loads(RULES.read_text(encoding="utf-8"))
    typo = rules["rakuten"]["badgeTypography"]
    gold = rules["rakuten"]["goldCircle1200"]
    patterns = typo["badgeLayoutPatternsLocked"]
    assert patterns["status"] == "ACCEPTED"
    assert typo["lockedUnitSet1000"]["1digit"]["freezeNumber"] is True
    assert typo["lockedUnitSet1000"]["2digit"]["freezeNumber"] is True
    assert typo["digitCompose2"]["frozen"] is True

    base0 = Image.open(base_path).convert("RGBA")
    if base0.size != (1200, 1200):
        base0 = base0.resize((1200, 1200), Image.Resampling.LANCZOS)

    jobs = [
        ("LOCKED_P1_1fukuro.jpg", 1, "袋", False, "①"),
        ("LOCKED_P2_10kan_PAIR.jpg", 10, "缶", False, "②"),
        ("LOCKED_P3_21kan_DIGITS.jpg", 21, "缶", False, "③"),
    ]
    for name, n, unit, force, label in jobs:
        t = json.loads(json.dumps(typo))
        if force:
            t["lockedUnitSet1000"]["2digit"]["forceDigits"] = True
            t["lockedUnitSet1000"]["2digit"]["preferPair"] = False
        out, meta = compose_badge_with_glyphs(
            base0.copy(),
            set_count=n,
            unit=unit,
            cx=int(gold["cx"]),
            cy=int(gold["cy"]),
            diameter=int(gold["diameter"]),
            typo=t,
            glyph_dirs=dirs,
            canvas=1200,
        )
        path = out_dir / name
        out.convert("RGB").save(path, quality=95)
        print(label, path.name, meta.get("numSource"), meta.get("numberFitMode"), meta.get("numberPos"))


if __name__ == "__main__":
    main()
