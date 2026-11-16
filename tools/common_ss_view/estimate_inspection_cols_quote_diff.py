# -*- coding: utf-8 -*-
"""見積検品: 判断の右に問屋見積、交渉の前に差分（希望−見積）を入れる。

新列:
  F 判断（端的）
  G メイン画像
  H 問屋見積価格（税抜）
  I 希望価格（税抜）
  J 提示価格（税抜）
  K 差分（希望−見積）  ※マイナス=見積が高い → ▲赤字＋交渉に注意印
  L 値下げ交渉

使い方:
  python tools/common_ss_view/estimate_inspection_cols_quote_diff.py --dry
  python tools/common_ss_view/estimate_inspection_cols_quote_diff.py
"""
from __future__ import annotations

import argparse
import math
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT.parent / "competitor_store"))
from client import sheets_service  # noqa: E402

RES = "1tf7gvkD88yyNz7JWXfNysBcIqSZDlI9dC-l6gOPyLjE"
SANK = "17nXSRjFp8xy-3EaBmp-LIu_tMc9ZOgfkhJoZjO3FLIc"
SHEET = "見積検品"
LIST = "仕入検討品番リスト"

# 新レイアウト（1-based）。F=判断（2026-11-15）の右から画像・価格
COL_JUDGE = 6
COL_IMG = 7
COL_QUOTE = 8
COL_HOPE = 9
COL_OFFER = 10
COL_DIFF = 11
COL_NEGO = 12


def num(v):
    try:
        if v is None or v == "":
            return None
        x = float(str(v).replace(",", "").replace("円", "").replace("▲", "-").strip())
        return x if x == x else None
    except Exception:
        return None


def jan_norm(v) -> str:
    s = str(v or "").strip()
    if not s:
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    if s.replace(".", "", 1).isdigit():
        try:
            s = str(int(float(s)))
        except Exception:
            pass
    if s.isdigit() and len(s) < 13:
        s = s.zfill(13)
    return s


def load_paste_ex(svc) -> dict[str, float]:
    rows = (
        svc.spreadsheets()
        .values()
        .get(
            spreadsheetId=SANK,
            range="'貼付sheet'!A3:F",
            valueRenderOption="UNFORMATTED_VALUE",
        )
        .execute()
        .get("values")
        or []
    )
    out = {}
    for r in rows[1:]:
        while len(r) < 6:
            r.append("")
        j = jan_norm(r[2])
        p = num(r[5])
        if len(j) >= 8 and p and p > 0:
            out[j] = p
    return out


def load_paste_name_jan(svc) -> dict:
    """貼付の商品名 → JAN（検品キーが貼付名のとき用）。"""
    rows = (
        svc.spreadsheets()
        .values()
        .get(
            spreadsheetId=SANK,
            range="'貼付sheet'!A3:F",
            valueRenderOption="UNFORMATTED_VALUE",
        )
        .execute()
        .get("values")
        or []
    )
    out = {}
    for r in rows[1:]:
        while len(r) < 6:
            r.append("")
        j = jan_norm(r[2])
        p = num(r[5])
        if len(j) < 8 or not p:
            continue
        maker = str(r[0] or "").strip()
        name = str(r[3] or "").strip()
        if not name:
            continue
        nb = re.sub(r"\s+", "", name)
        out[nb] = j
        out[(maker, nb)] = j
        # 税込仕入付きキー（FLOOR概算と一致しやすいよう両系統）
        for tax in (0.08, 0.1):
            incl = int(math.floor(p * (1.0 + tax) + 1e-9))
            key = f"{name}／仕入{incl}"
            out[key] = j
            out[re.sub(r"\s+", "", key)] = j
    return out


def load_key_jan(svc) -> dict:
    """商品キー／(メーカー,単品名) → 代表JAN。

    見積候補に加え見送りも見る（門落ち込み検品用）。
    貼付商品名も足す（1商品1行でキーが貼付名のとき）。
    """
    hdr = [
        str(c).strip()
        for c in (
            (
                svc.spreadsheets()
                .values()
                .get(spreadsheetId=RES, range=f"'{LIST}'!4:4")
                .execute()
                .get("values")
                or [[]]
            )[0]
        )
    ]
    idx = {h: i for i, h in enumerate(hdr)}
    rows = (
        svc.spreadsheets()
        .values()
        .get(
            spreadsheetId=RES,
            range=f"'{LIST}'!A5:BO",
            valueRenderOption="UNFORMATTED_VALUE",
        )
        .execute()
        .get("values")
        or []
    )
    out = load_paste_name_jan(svc)
    for r in rows:
        while len(r) < len(hdr):
            r.append("")
        flag = str(r[idx["見積候補"]]).strip()
        if flag not in ("見積候補", "見送り"):
            continue
        name = str(r[idx["単品商品名"]]).strip()
        cost = num(r[idx["1点仕入れ額（税込み）"]])
        j = jan_norm(r[idx["JANコード"]])
        if not j:
            continue
        maker = str(r[idx["メーカー名"]]).strip()
        if name and "要確認" not in name:
            nb = re.sub(r"\s+", "", name)
            out[(maker, nb)] = j
            out[nb] = j
            if cost and cost > 0:
                key = f"{name}／仕入{int(round(cost)) if abs(cost - round(cost)) < 1e-6 else cost}"
                out[key] = j
                out[re.sub(r"\s+", "", key)] = j
    return out


def ensure_columns(svc, sheet_id: int, dry: bool) -> bool:
    """既に問屋見積列が無ければ2列挿入。戻り: 今回挿入したか。"""
    hdr = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=RES, range=f"'{SHEET}'!1:1")
        .execute()
        .get("values")
        or [[]]
    )[0]
    joined = "|".join(str(h) for h in hdr)
    if "問屋見積" in joined or "差分（希望" in joined:
        print("columns already present", hdr[5:12])
        return False
    print("will insert 2 cols at G and before 交渉")
    if dry:
        return True
    # 1) insert at G (index 6) for 問屋見積
    # 2) after that, 交渉 was at I(8) → now J(9); insert at J for 差分
    reqs = [
        {
            "insertDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": 6,
                    "endIndex": 7,
                },
                "inheritFromBefore": True,
            }
        },
        {
            "insertDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": 9,
                    "endIndex": 10,
                },
                "inheritFromBefore": True,
            }
        },
    ]
    svc.spreadsheets().batchUpdate(spreadsheetId=RES, body={"requests": reqs}).execute()
    return True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry", action="store_true")
    args = ap.parse_args()
    svc = sheets_service(write=not args.dry, interactive=False)

    meta = (
        svc.spreadsheets()
        .get(
            spreadsheetId=RES,
            ranges=[f"'{SHEET}'"],
            fields="sheets(properties(sheetId))",
        )
        .execute()
    )
    sid = meta["sheets"][0]["properties"]["sheetId"]
    ensure_columns(svc, sid, dry=args.dry)

    # headers
    headers = [
        "メイン画像",
        "問屋見積価格（税抜）",
        "希望価格（税抜）",
        "提示価格（税抜）",
        "差分（希望−見積）",
        "値下げ交渉",
    ]
    paste = load_paste_ex(svc)
    key_jan = load_key_jan(svc)

    insp = (
        svc.spreadsheets()
        .values()
        .get(
            spreadsheetId=RES,
            range=f"'{SHEET}'!A1:L500",
            valueRenderOption="UNFORMATTED_VALUE",
        )
        .execute()
        .get("values")
        or []
    )
    if not insp:
        raise SystemExit("empty")

    # After insert, read again if we inserted in non-dry
    if not args.dry:
        insp = (
            svc.spreadsheets()
            .values()
            .get(
                spreadsheetId=RES,
                range=f"'{SHEET}'!A1:L500",
                valueRenderOption="UNFORMATTED_VALUE",
            )
            .execute()
            .get("values")
            or []
        )

    quote_vals = []
    diff_vals = []
    nego_vals = []
    report = []
    for row in insp[1:]:
        while len(row) < COL_NEGO:
            row.append("")
        maker = str(row[1] if len(row) > 1 else "")
        key = str(row[2] if len(row) > 2 else "")
        hope = num(row[COL_HOPE - 1]) if len(row) >= COL_HOPE else None
        # if columns not yet shifted in dry preview, hope may still be at G=7
        if hope is None and len(row) >= 7:
            hope = num(row[6])

        base = key.split("／")[0].strip()
        nb = re.sub(r"\s+", "", base)
        j = key_jan.get(key) or key_jan.get(re.sub(r"\s+", "", key))
        if not j:
            j = key_jan.get((maker, nb)) or key_jan.get(nb)
        # fuzzy（メーカー一致＋名前の包含）
        if not j:
            for k, jan in key_jan.items():
                if isinstance(k, tuple) and k[0] == maker and (nb in k[1] or k[1] in nb):
                    j = jan
                    break
        if not j and nb:
            for k, jan in key_jan.items():
                if isinstance(k, str) and not k.startswith("jan") and "／" not in k:
                    if len(k) >= 4 and (nb in k or k in nb):
                        j = jan
                        break

        quote = paste.get(j) if j else None
        # 貼付JANが取れないときだけ、商品キーの仕入税込÷1.08で概算（正は貼付税抜）
        if quote is None and "仕入なし" not in key:
            m = re.search(r"仕入\s*([0-9.]+)", key)
            if m:
                cost_in = float(m.group(1))
                quote = round(cost_in / 1.08, 1)

        diff = None
        if hope is not None and quote is not None:
            diff = round(hope - quote, 1)

        note = ""
        existing_nego = ""
        if len(row) >= COL_NEGO:
            existing_nego = str(row[COL_NEGO - 1] or "").strip()
        # strip old auto note
        if existing_nego.startswith("注意＝見積もりが高い"):
            existing_nego = existing_nego.split("／", 1)[-1].strip() if "／" in existing_nego else ""
            if existing_nego.startswith("注意＝"):
                existing_nego = ""
        if diff is not None and diff < 0:
            note = "注意＝見積もりが高い"
            existing_nego = (note + ("／" + existing_nego if existing_nego else "")).strip()

        quote_vals.append([quote if quote is not None else ""])
        diff_vals.append([diff if diff is not None else ""])
        nego_vals.append([existing_nego])
        report.append((key[:28], j, quote, hope, diff, note))

    print("samples:")
    for t in report[:8]:
        print(t)
    print("... total", len(report))

    if args.dry:
        print("DRY ok")
        return

    # write headers G–L（Fは判断列）
    svc.spreadsheets().values().update(
        spreadsheetId=RES,
        range=f"'{SHEET}'!G1:L1",
        valueInputOption="USER_ENTERED",
        body={"values": [headers]},
    ).execute()
    n = len(quote_vals)
    end = 1 + n
    svc.spreadsheets().values().update(
        spreadsheetId=RES,
        range=f"'{SHEET}'!H2:H{end}",
        valueInputOption="USER_ENTERED",
        body={"values": quote_vals},
    ).execute()
    svc.spreadsheets().values().update(
        spreadsheetId=RES,
        range=f"'{SHEET}'!K2:K{end}",
        valueInputOption="USER_ENTERED",
        body={"values": diff_vals},
    ).execute()
    svc.spreadsheets().values().update(
        spreadsheetId=RES,
        range=f"'{SHEET}'!L2:L{end}",
        valueInputOption="USER_ENTERED",
        body={"values": nego_vals},
    ).execute()

    # number format for J: minus as red ▲
    svc.spreadsheets().batchUpdate(
        spreadsheetId=RES,
        body={
            "requests": [
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": sid,
                            "startRowIndex": 1,
                            "endRowIndex": end,
                            "startColumnIndex": COL_DIFF - 1,
                            "endColumnIndex": COL_DIFF,
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "numberFormat": {
                                    "type": "NUMBER",
                                    "pattern": '#,##0.0;[Red]"▲"#,##0.0',
                                }
                            }
                        },
                        "fields": "userEnteredFormat.numberFormat",
                    }
                }
            ]
        },
    ).execute()
    print("WROTE quote/diff/nego", n)


if __name__ == "__main__":
    main()
