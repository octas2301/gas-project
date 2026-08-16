# -*- coding: utf-8 -*-
"""L6: 永谷園 みそ汁 Catalog 1頁(≤20)→門→①候補。モールヒット非書。既定 dry。"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from apply_keepa_full import append_log, sheets_service

SMOKE = Path(__file__).resolve().parents[1] / "spapi_smoke" / "config.local.json"
SEARCH = Path(__file__).resolve().parents[1] / "purchase_research_path3" / "search_catalog_keywords.py"
NAGA = Path(__file__).resolve().parent / "apply_naga_cap20.py"
OUT = Path(__file__).resolve().parents[1] / "purchase_research_path3" / "out"


def main() -> int:
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    ap.add_argument("--keywords", default="永谷園 みそ汁")
    ap.add_argument("--needles", default="永谷園,みそ汁")
    ap.add_argument("--run-id", default="pr_20260816_l6")
    args = ap.parse_args()
    svc = sheets_service(write=True)
    ok = SMOKE.is_file() and SEARCH.is_file() and NAGA.is_file()
    rid = args.run_id + ("dry" if not args.apply else "col")
    line = "runId=%s cfg=%s needles=%s cap=20 %s" % (
        rid,
        SMOKE.is_file(),
        args.needles,
        "PASS" if ok else "FAIL",
    )
    print(line)
    if not args.apply:
        append_log(svc, "L6", line)
        return 0 if ok else 1
    if not ok:
        return 1
    OUT.mkdir(parents=True, exist_ok=True)
    r = subprocess.run(
        [
            sys.executable,
            str(SEARCH),
            "--keywords",
            args.keywords,
            "--max-pages",
            "1",
            "--out-dir",
            str(OUT),
        ],
        cwd=str(SEARCH.parent),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    print(r.stdout[-1500:] if r.stdout else "")
    if r.returncode != 0:
        print(r.stderr[-800:] if r.stderr else "")
        append_log(svc, "L6", "runId=%s FAIL catalog rc=%s" % (rid, r.returncode))
        return 1
    dest = ""
    for ln in (r.stdout or "").splitlines():
        if ln.startswith("report="):
            dest = ln.split("=", 1)[1].strip()
    if not dest:
        append_log(svc, "L6", "runId=%s FAIL no_report" % rid)
        return 1
    r2 = subprocess.run(
        [
            sys.executable,
            str(NAGA),
            "--json",
            dest,
            "--needles",
            args.needles,
            "--maker",
            "永谷園",
            "--run-id",
            args.run_id,
            "--apply",
            "--live",
        ],
        cwd=str(NAGA.parent),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    print(r2.stdout[-2000:] if r2.stdout else "")
    if r2.returncode != 0:
        print(r2.stderr[-800:] if r2.stderr else "")
        append_log(svc, "L6", "runId=%s FAIL naga rc=%s" % (rid, r2.returncode))
        return 1
    append_log(svc, "L6", "runId=%s catalog+naga PASS " % rid + (r2.stdout or "").replace("\n", " ")[-400:])
    print("VERIFY PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
