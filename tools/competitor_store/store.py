# -*- coding: utf-8 -*-
from __future__ import annotations

import csv
from pathlib import Path

from schema import (
    FIELD_MAP_HEADERS,
    HITS_HEADERS,
    KEEPA_FULL_HEADERS,
    KEEPA_SNAP_HEADERS,
    MAKER_HEADERS,
    META_HEADERS,
    SELLER_HEADERS,
    SHEET_FIELD_MAP,
    SHEET_HITS,
    SHEET_KEEPA,
    SHEET_KEEPA_FULL,
    SHEET_MAKER,
    SHEET_META,
    SHEET_SELLER,
)

LOCAL = Path(__file__).resolve().parent / "local_store"


class LocalStore:
    def __init__(self, root: Path | None = None):
        self.root = root or LOCAL
        self.root.mkdir(parents=True, exist_ok=True)

    def _path(self, name: str) -> Path:
        return self.root / (name + ".csv")

    def ensure_schema(self) -> None:
        specs = {
            SHEET_FIELD_MAP: FIELD_MAP_HEADERS,
            SHEET_HITS: HITS_HEADERS,
            SHEET_KEEPA: KEEPA_SNAP_HEADERS,
            SHEET_KEEPA_FULL: KEEPA_FULL_HEADERS,
            SHEET_META: META_HEADERS,
            SHEET_MAKER: MAKER_HEADERS,
            SHEET_SELLER: SELLER_HEADERS,
        }
        for name, headers in specs.items():
            p = self._path(name)
            if not p.exists() or p.stat().st_size < 4:
                with p.open("w", encoding="utf-8-sig", newline="") as f:
                    csv.writer(f).writerow(headers)
                continue
            try:
                old = self.headers(name)
            except (UnicodeDecodeError, StopIteration, csv.Error):
                with p.open("w", encoding="utf-8-sig", newline="") as f:
                    csv.writer(f).writerow(headers)
                continue
            missing = [h for h in headers if h not in old]
            if not missing:
                continue
            rows = self.rows(name)
            new_headers = old + missing
            self.replace(name, new_headers, rows)

    def headers(self, name: str) -> list[str]:
        with self._path(name).open(encoding="utf-8-sig", newline="") as f:
            return next(csv.reader(f))

    def rows(self, name: str) -> list[dict]:
        with self._path(name).open(encoding="utf-8-sig", newline="") as f:
            return list(csv.DictReader(f))

    def replace(self, name: str, headers: list[str], rows: list[dict]) -> None:
        with self._path(name).open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=headers, extrasaction="ignore")
            w.writeheader()
            w.writerows(rows)

    def append(self, name: str, rec: dict) -> None:
        headers = self.headers(name)
        rows = self.rows(name)
        rows.append({h: rec.get(h, "") for h in headers})
        self.replace(name, headers, rows)

    def row_count(self, name: str) -> int:
        return len(self.rows(name))
