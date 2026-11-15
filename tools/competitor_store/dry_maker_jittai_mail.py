# -*- coding: utf-8 -*-
"""実態=要確認のメール下書きを標準出力。送信しない。"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from client import sheets_service  # noqa: E402
from dry_maker_sum import as_dicts, read_all  # noqa: E402
from maker_alias import JITTAI_CHECK, amazon_search_url  # noqa: E402
from schema import SHEET_MAKER_SUM  # noqa: E402

COMPETITOR_SS = "1UrdWDBw8NcuOf71Bi-2m8WNQDW2onIkA-zl6mLE7AHs"
SS_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1UrdWDBw8NcuOf71Bi-2m8WNQDW2onIkA-zl6mLE7AHs/edit"
)


def main() -> int:
    svc = sheets_service(write=False, interactive=True)
    if not svc:
        print("no_creds")
        return 1
    _, rows = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_MAKER_SUM))
    checks = [r for r in rows if str(r.get("実態") or "").strip() == JITTAI_CHECK]
    print("To: （担当）")
    print("Subject: 【仕入れ検討①】メーカー実態の要確認 %d件" % len(checks))
    print()
    print("メーカーサマリの「実態」が要確認の行です。")
    print("各行を 製造メーカー／入口ブランド／非仕入先 のいずれかに確定し、")
    print("シートの実態列を更新してください（または maker_alias.MAKER_JITTAI へ）。")
    print("確認用: 列「実態ck」または下の Amazon 検索URL。")
    print()
    print("シート: %s" % SS_URL)
    print("タブ: メーカーサマリ")
    print()
    print("| メーカー | 倉庫SKU | 通過 | 入口推奨 | Amazon検索 |")
    print("|----------|---------|------|----------|------------|")
    for r in checks:
        maker = r.get("メーカー") or ""
        url = str(r.get("実態ck") or "").strip() or amazon_search_url(maker)
        print(
            "| %s | %s | %s | %s | %s |"
            % (
                maker,
                r.get("倉庫SKU") or "",
                r.get("通過数") or "",
                r.get("入口推奨") or "",
                url,
            )
        )
    print()
    print("--- 以上（自動送信なし・下書きのみ）---")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
