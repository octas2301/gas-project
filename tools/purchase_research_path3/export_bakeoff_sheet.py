# -*- coding: utf-8 -*-
"""競合検索ベイクオフシート → CSV。"""
from pathlib import Path
import csv
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

sid = "1LIWp0qjgvPaZtjsIBmCGqCEEB7AA00nLmBA7iE1MI28"
SHEET = "競合検索ベイクオフ"
OUT = Path(__file__).resolve().parent / "rakuten_yahoo_bakeoff_AE.csv"
base = Path(__file__).resolve().parents[1] / "c1_hpc_packaged" / "secrets"
creds = Credentials.from_authorized_user_file(
    str(base / "token.json"),
    [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
    ],
)
if not creds.valid:
    creds.refresh(Request())
svc = build("sheets", "v4", credentials=creds, cache_discovery=False)
vals = (
    svc.spreadsheets()
    .values()
    .get(spreadsheetId=sid, range="'" + SHEET + "'")
    .execute()
    .get("values")
    or []
)
if not vals:
    raise SystemExit("sheet empty or missing: " + SHEET)
ncols = max(len(r) for r in vals)
header = (vals[0] + [""] * ncols)[:ncols]
rows = []
for r in vals[1:]:
    rec = {}
    for i, h in enumerate(header):
        rec[h] = r[i] if i < len(r) else ""
    rows.append(rec)
with OUT.open("w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=header)
    w.writeheader()
    w.writerows(rows)
print("rows", len(rows), "->", OUT)
