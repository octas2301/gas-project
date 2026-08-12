# -*- coding: utf-8 -*-
"""Sheets 認証（読取=C1 token／書込=token_sheets_rw）。"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

HERE = Path(__file__).resolve().parent
C1 = HERE.parent / "c1_hpc_packaged"


def _c1_paths() -> Tuple[Path, Path]:
    cfg_path = C1 / "config.local.json"
    if not cfg_path.is_file():
        cfg_path = C1 / "config.json"
    cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
    fetch = cfg.get("fetch") if isinstance(cfg.get("fetch"), dict) else {}
    cred = C1 / str(fetch.get("credentials_path") or "secrets/credentials.json")
    token = C1 / str(fetch.get("token_path") or "secrets/token.json")
    if not cred.is_absolute():
        cred = (C1 / cred).resolve()
    if not token.is_absolute():
        token = (C1 / token).resolve()
    return cred, token


def get_creds(*, write: bool = False):
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow

    cred_path, token_ro = _c1_paths()
    if write:
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        token_path = C1 / "secrets" / "token_sheets_rw.json"
    else:
        scopes = None  # use token as-is
        token_path = token_ro

    creds = None
    if token_path.is_file():
        if write:
            creds = Credentials.from_authorized_user_file(str(token_path), scopes)
        else:
            creds = Credentials.from_authorized_user_file(str(token_path))
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    if write and (not creds or not creds.valid):
        flow = InstalledAppFlow.from_client_secrets_file(str(cred_path), scopes)
        creds = flow.run_local_server(port=0)
        token_path.parent.mkdir(parents=True, exist_ok=True)
        token_path.write_text(creds.to_json(), encoding="utf-8")
    return creds


def sheets_service(*, write: bool = False):
    from googleapiclient.discovery import build

    return build("sheets", "v4", credentials=get_creds(write=write), cache_discovery=False)


def read_sheet_rows(
    svc, spreadsheet_id: str, title: str, *, formulas: bool = True
) -> Tuple[List[str], List[Dict[str, Any]]]:
    """行読取。既定は FORMULA（IMAGE式の往復保全）。"""
    rng = f"'{title.replace(chr(39), chr(39)+chr(39))}'!A1:AZ5000"
    kwargs = {
        "spreadsheetId": spreadsheet_id,
        "range": rng,
    }
    if formulas:
        kwargs["valueRenderOption"] = "FORMULA"
    data = svc.spreadsheets().values().get(**kwargs).execute()
    values = data.get("values") or []
    if not values:
        return [], []
    headers = [str(h).strip() for h in values[0]]
    rows: List[Dict[str, Any]] = []
    for raw in values[1:]:
        d = {headers[i]: (raw[i] if i < len(raw) else "") for i in range(len(headers))}
        if any(str(v).strip() for v in d.values()):
            rows.append(d)
    return headers, rows


def ensure_sheet(svc, spreadsheet_id: str, title: str) -> None:
    meta = svc.spreadsheets().get(spreadsheetId=spreadsheet_id, fields="sheets.properties.title").execute()
    titles = [s["properties"]["title"] for s in meta.get("sheets") or []]
    if title in titles:
        return
    svc.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": [{"addSheet": {"properties": {"title": title}}}]},
    ).execute()


def rename_sheet_if_exists(svc, spreadsheet_id: str, old: str, new: str) -> bool:
    meta = svc.spreadsheets().get(
        spreadsheetId=spreadsheet_id, fields="sheets.properties(sheetId,title)"
    ).execute()
    titles = {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta.get("sheets") or []}
    if new in titles:
        return False
    if old not in titles:
        return False
    svc.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {"sheetId": titles[old], "title": new},
                        "fields": "title",
                    }
                }
            ]
        },
    ).execute()
    return True


def ensure_column_capacity(
    svc, spreadsheet_id: str, title: str, min_cols: int
) -> None:
    """シートの列数を min_cols 以上に拡張（ヘッダ追記前に必要）。"""
    meta = (
        svc.spreadsheets()
        .get(
            spreadsheetId=spreadsheet_id,
            fields="sheets(properties(sheetId,title,gridProperties(columnCount,rowCount)))",
        )
        .execute()
    )
    for s in meta.get("sheets") or []:
        props = s.get("properties") or {}
        if props.get("title") != title:
            continue
        sid = int(props["sheetId"])
        gp = props.get("gridProperties") or {}
        cur = int(gp.get("columnCount") or 26)
        if cur >= min_cols:
            return
        need = int(min_cols) - cur
        svc.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={
                "requests": [
                    {
                        "appendDimension": {
                            "sheetId": sid,
                            "dimension": "COLUMNS",
                            "length": need,
                        }
                    }
                ]
            },
        ).execute()
        return
    raise RuntimeError("sheet not found: %s" % title)


def ensure_headers_append(
    svc, spreadsheet_id: str, title: str, required: List[str]
) -> List[str]:
    """足りないヘッダを右端に追記。既存データは消さない。戻り値=最終ヘッダ。"""
    headers, _ = read_sheet_rows(svc, spreadsheet_id, title)
    if not headers:
        write_headers_and_rows(svc, spreadsheet_id, title, required, [], clear=True)
        return list(required)
    missing = [h for h in required if h not in headers]
    if not missing:
        return headers
    ensure_column_capacity(
        svc, spreadsheet_id, title, len(headers) + len(missing)
    )
    q = title.replace("'", "''")

    def col_a1(n: int) -> str:
        s = ""
        while n:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s

    start_col = len(headers) + 1
    svc.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"'{q}'!{col_a1(start_col)}1",
        valueInputOption="USER_ENTERED",
        body={"values": [missing]},
    ).execute()
    return headers + missing


def update_row_fields(
    svc,
    spreadsheet_id: str,
    title: str,
    row_1based: int,
    fields: Dict[str, Any],
) -> int:
    """1行の指定ヘッダ列だけ更新。戻り値=更新セル数。"""
    headers, _ = read_sheet_rows(svc, spreadsheet_id, title)
    if not headers or row_1based < 2:
        return 0
    q = title.replace("'", "''")

    def col_a1(n: int) -> str:
        s = ""
        while n:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s

    data = []
    for name, val in fields.items():
        if name not in headers:
            continue
        col = headers.index(name) + 1
        data.append(
            {
                "range": f"'{q}'!{col_a1(col)}{row_1based}",
                "values": [[val]],
            }
        )
    if not data:
        return 0
    svc.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"valueInputOption": "USER_ENTERED", "data": data},
    ).execute()
    return len(data)


def write_headers_and_rows(
    svc,
    spreadsheet_id: str,
    title: str,
    headers: List[str],
    rows: List[List[Any]],
    *,
    clear: bool = True,
) -> None:
    ensure_sheet(svc, spreadsheet_id, title)
    q = title.replace("'", "''")
    if clear:
        svc.spreadsheets().values().clear(
            spreadsheetId=spreadsheet_id, range=f"'{q}'"
        ).execute()
    body = [headers] + rows
    svc.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"'{q}'!A1",
        valueInputOption="USER_ENTERED",
        body={"values": body},
    ).execute()
    # タイムセール_マスタ: ポイント表示形式＋ヘッダ色分け
    try:
        from sheet_schema import MASTER_SHEET as _MS

        if title == _MS:
            apply_master_point_unit_formats(svc, spreadsheet_id, title)
            apply_master_header_group_colors(svc, spreadsheet_id, title)
            apply_master_human_input_yellow(svc, spreadsheet_id, title)
            apply_master_display_formulas(svc, spreadsheet_id, title)
            apply_master_recovery_validations(svc, spreadsheet_id, title)
    except Exception:
        pass


def sheet_id_by_title(svc, spreadsheet_id: str, title: str) -> Optional[int]:
    meta = svc.spreadsheets().get(
        spreadsheetId=spreadsheet_id, fields="sheets.properties(sheetId,title)"
    ).execute()
    for s in meta.get("sheets") or []:
        props = s.get("properties") or {}
        if props.get("title") == title:
            return int(props["sheetId"])
    return None


def apply_number_format_columns(
    svc,
    spreadsheet_id: str,
    title: str,
    *,
    col_formats: Dict[int, str],
    start_row: int = 1,
    end_row: int = 5000,
) -> None:
    """列インデックス(0始まり)→カスタム表示形式。セル値は数値のまま、表示だけ単位付き。"""
    sid = sheet_id_by_title(svc, spreadsheet_id, title)
    if sid is None:
        raise RuntimeError("sheet not found: %s" % title)
    reqs = []
    for col, pattern in col_formats.items():
        reqs.append(
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sid,
                        "startRowIndex": start_row,
                        "endRowIndex": end_row,
                        "startColumnIndex": col,
                        "endColumnIndex": col + 1,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {"type": "NUMBER", "pattern": pattern}
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat",
                }
            }
        )
    if reqs:
        svc.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body={"requests": reqs}
        ).execute()


def _hex_to_rgb01(hex_color: str) -> Dict[str, float]:
    h = hex_color.lstrip("#")
    return {
        "red": int(h[0:2], 16) / 255.0,
        "green": int(h[2:4], 16) / 255.0,
        "blue": int(h[4:6], 16) / 255.0,
    }


def apply_master_header_group_colors(
    svc, spreadsheet_id: str, title: str
) -> int:
    """
    タイムセール_マスタ 1行目をグループ色で塗る（基本／SC／ポイント／価格戻し／メモ）。
    戻り値: 色付けした列数。
    """
    from sheet_schema import MASTER_HEADER_COLOR_GROUPS

    sid = sheet_id_by_title(svc, spreadsheet_id, title)
    if sid is None:
        raise RuntimeError("sheet not found: %s" % title)
    headers, _ = read_sheet_rows(svc, spreadsheet_id, title)
    if not headers:
        return 0
    name_to_color: Dict[str, str] = {}
    for _label, hex_c, names in MASTER_HEADER_COLOR_GROUPS:
        for n in names:
            name_to_color[str(n)] = hex_c
    reqs = []
    painted = 0
    for i, h in enumerate(headers):
        name = str(h).strip()
        hex_c = name_to_color.get(name)
        if not hex_c:
            continue
        painted += 1
        reqs.append(
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sid,
                        "startRowIndex": 0,
                        "endRowIndex": 1,
                        "startColumnIndex": i,
                        "endColumnIndex": i + 1,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": _hex_to_rgb01(hex_c),
                            "textFormat": {
                                "bold": True,
                                "foregroundColor": {
                                    "red": 0.125,
                                    "green": 0.129,
                                    "blue": 0.141,
                                },
                            },
                        }
                    },
                    "fields": (
                        "userEnteredFormat.backgroundColor,"
                        "userEnteredFormat.textFormat.bold,"
                        "userEnteredFormat.textFormat.foregroundColor"
                    ),
                }
            }
        )
    if reqs:
        svc.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body={"requests": reqs}
        ).execute()
    return painted


def apply_master_point_unit_formats(svc, spreadsheet_id: str, title: str) -> None:
    """
    ポイント%列 → 表示 1%／ポイント円列 → 表示 45円。
    ヘッダ行は対象外。値は数値のまま（文字列の「%」「円」は入れない）。
    """
    headers, _ = read_sheet_rows(svc, spreadsheet_id, title)
    if not headers:
        return
    pct_names = (
        "期間中ポイント%",
        "セール前ポイント%",
        "出品者ポイント現在%",
        "出品者ポイント%",
        "販促ポイント%",
        "減衰段%",
        "減衰中ポイント%",
        "次回減衰後%",
    )
    yen_names = (
        "期間中ポイント円",
        "セール前ポイント円",
        "出品者ポイント現在円",
        "期間中ポイント",
        "セール前ポイント",
        "出品者ポイント",
        "販促ポイント円",
        "実質価格円",
        "目標売価円",
        "最終売価円",
        "現在売価円",
    )
    formats: Dict[int, str] = {}
    for i, h in enumerate(headers):
        name = str(h).strip()
        if name in pct_names or (name.endswith("%") and "ポイント" in name):
            formats[i] = '0"%"'
        elif name in yen_names or (
            "ポイント" in name and name.endswith("円")
        ) or (
            "ポイント" in name
            and "%" not in name
            and name
            in ("期間中ポイント", "セール前ポイント", "出品者ポイント")
        ):
            formats[i] = '0"円"'
    if not formats:
        return
    apply_number_format_columns(svc, spreadsheet_id, title, col_formats=formats)


def apply_master_recovery_validations(
    svc, spreadsheet_id: str, title: str, *, data_rows: int = 200
) -> None:
    """
    減衰期間／減衰間隔だけリスト検証。実質戻しブロック全体の旧検証を先に削除。
    """
    sid = sheet_id_by_title(svc, spreadsheet_id, title)
    if sid is None:
        raise RuntimeError("sheet not found: %s" % title)
    headers, rows = read_sheet_rows(svc, spreadsheet_id, title)
    if not headers:
        return
    idx = {str(h).strip(): i for i, h in enumerate(headers)}
    start = idx.get("目標売価円", idx.get("減衰期間", idx.get("戻し期間")))
    end = idx.get("現在売価円", idx.get("減衰間隔", idx.get("戻し間隔")))
    if start is None:
        return
    if end is None:
        end = start
    end = min(len(headers) - 1, max(end, start) + 15)
    end_row = max(data_rows, len(rows) + 1)  # exclusive end for API = header+data
    reqs: List[Dict[str, Any]] = [
        {
            "setDataValidation": {
                "range": {
                    "sheetId": sid,
                    "startRowIndex": 1,
                    "endRowIndex": end_row + 1,
                    "startColumnIndex": start,
                    "endColumnIndex": end + 1,
                },
                "rule": None,
            }
        }
    ]
    period_opts = ["1か月", "2か月", "3か月", "4か月", "5か月", "6か月"]
    interval_opts = ["1週間", "2週間", "1か月", "2か月"]
    n_rows = max(len(rows), 1)
    period_col = idx.get("減衰期間", idx.get("戻し期間"))
    if period_col is not None:
        c = period_col
        reqs.append(
            {
                "setDataValidation": {
                    "range": {
                        "sheetId": sid,
                        "startRowIndex": 1,
                        "endRowIndex": 1 + n_rows,
                        "startColumnIndex": c,
                        "endColumnIndex": c + 1,
                    },
                    "rule": {
                        "condition": {
                            "type": "ONE_OF_LIST",
                            "values": [{"userEnteredValue": v} for v in period_opts],
                        },
                        "showCustomUi": True,
                        "strict": True,
                    },
                }
            }
        )
    interval_col = idx.get("減衰間隔", idx.get("戻し間隔"))
    if interval_col is not None:
        c = interval_col
        reqs.append(
            {
                "setDataValidation": {
                    "range": {
                        "sheetId": sid,
                        "startRowIndex": 1,
                        "endRowIndex": 1 + n_rows,
                        "startColumnIndex": c,
                        "endColumnIndex": c + 1,
                    },
                    "rule": {
                        "condition": {
                            "type": "ONE_OF_LIST",
                            "values": [{"userEnteredValue": v} for v in interval_opts],
                        },
                        "showCustomUi": True,
                        "strict": True,
                    },
                }
            }
        )
    req_col = idx.get("減衰実行依頼")
    if req_col is not None:
        c = req_col
        reqs.append(
            {
                "setDataValidation": {
                    "range": {
                        "sheetId": sid,
                        "startRowIndex": 1,
                        "endRowIndex": 1 + n_rows,
                        "startColumnIndex": c,
                        "endColumnIndex": c + 1,
                    },
                    "rule": {
                        "condition": {"type": "BOOLEAN"},
                        "showCustomUi": True,
                    },
                }
            }
        )
    svc.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body={"requests": reqs}
    ).execute()


def apply_master_human_input_yellow(
    svc, spreadsheet_id: str, title: str, *, data_rows: int = 500
) -> int:
    """
    人入力列の行2以降を黄色（#FFF2CC）。他列のデータ帯の黄は消さない（ヘッダは別処理）。
    戻り値: 色付けした列数。
    """
    from sheet_schema import MASTER_HUMAN_INPUT_BG, MASTER_HUMAN_INPUT_COLS

    sid = sheet_id_by_title(svc, spreadsheet_id, title)
    if sid is None:
        raise RuntimeError("sheet not found: %s" % title)
    headers, rows = read_sheet_rows(svc, spreadsheet_id, title)
    if not headers:
        return 0
    end_row = max(data_rows, len(rows) + 1)
    bg = _hex_to_rgb01(MASTER_HUMAN_INPUT_BG)
    human = set(MASTER_HUMAN_INPUT_COLS)
    reqs: List[Dict[str, Any]] = []
    painted = 0
    for i, h in enumerate(headers):
        name = str(h).strip()
        if name not in human:
            continue
        painted += 1
        reqs.append(
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sid,
                        "startRowIndex": 1,
                        "endRowIndex": end_row + 1,
                        "startColumnIndex": i,
                        "endColumnIndex": i + 1,
                    },
                    "cell": {
                        "userEnteredFormat": {"backgroundColor": bg}
                    },
                    "fields": "userEnteredFormat.backgroundColor",
                }
            }
        )
    if reqs:
        svc.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body={"requests": reqs}
        ).execute()
    return painted


def _col_index_to_a1(col_zero_based: int) -> str:
    """0-based column index → A1 letter(s)."""
    n = col_zero_based + 1
    letters = []
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters.append(chr(65 + rem))
    return "".join(reversed(letters))


def master_next_pct_formula(
    row: int,
    *,
    c_step: str,
    c_active: str,
    c_pct: str,
    c_before: str = "",
) -> str:
    """
    次回減衰後% のシート数式。f-string は単引用符のみ（\"\" が二重引用で途切れるため）。
    """
    before_expr = (
        f'IF(OR({c_before}{row}="",{c_before}{row}=0),1,{c_before}{row})'
        if c_before
        else "1"
    )
    return (
        f'=IF(OR({c_step}{row}="",AND({c_active}{row}="",{c_pct}{row}="")),"",'
        f'MAX({before_expr},'
        f'IF({c_active}{row}="",{c_pct}{row},{c_active}{row})-{c_step}{row}))'
    )


def apply_master_display_formulas(
    svc, spreadsheet_id: str, title: str, *, data_rows: int = 500
) -> int:
    """
    販促ポイント円・実質価格円に行ごと数式を入れる（入力変更で即再計算）。
    戻り値: 数式を書いたセル数。
    """
    from sheet_schema import (
        EFFECTIVE_PRICE_COL,
        POINT_BEFORE_COL,
        PRICE_TARGET_COL,
        PROMO_POINT_PCT_COL,
        PROMO_POINT_YEN_COL,
        TAPER_ACTIVE_COL,
        TAPER_NEXT_PCT_COL,
        PRICE_RECOVERY_STEP_COL,
    )

    headers, rows = read_sheet_rows(svc, spreadsheet_id, title)
    if not headers:
        return 0
    idx = {str(h).strip(): i for i, h in enumerate(headers)}
    if PRICE_TARGET_COL not in idx and "最終売価円" in idx:
        idx[PRICE_TARGET_COL] = idx["最終売価円"]
    need = (PRICE_TARGET_COL, PROMO_POINT_PCT_COL, PROMO_POINT_YEN_COL, EFFECTIVE_PRICE_COL)
    if any(k not in idx for k in need):
        return 0

    c_target = _col_index_to_a1(idx[PRICE_TARGET_COL])
    c_pct = _col_index_to_a1(idx[PROMO_POINT_PCT_COL])
    c_yen = _col_index_to_a1(idx[PROMO_POINT_YEN_COL])
    end_row = max(2, min(1 + max(data_rows, len(rows)), 2000))

    yen_col: List[List[str]] = []
    eff_col: List[List[str]] = []
    next_col: List[List[str]] = []
    has_next = (
        TAPER_NEXT_PCT_COL in idx
        and PRICE_RECOVERY_STEP_COL in idx
        and TAPER_ACTIVE_COL in idx
    )
    c_step = _col_index_to_a1(idx[PRICE_RECOVERY_STEP_COL]) if has_next else ""
    c_active = _col_index_to_a1(idx[TAPER_ACTIVE_COL]) if has_next else ""
    c_before = _col_index_to_a1(idx[POINT_BEFORE_COL]) if POINT_BEFORE_COL in idx else ""
    for r in range(2, end_row + 1):
        yen_col.append(
            [
                f'=IF(OR({c_target}{r}="",{c_pct}{r}=""),"",'
                f"ROUND({c_target}{r}*{c_pct}{r}/100,0))"
            ]
        )
        eff_col.append(
            [
                f'=IF(OR({c_target}{r}="",{c_yen}{r}=""),"",'
                f"{c_target}{r}-{c_yen}{r})"
            ]
        )
        if has_next:
            next_col.append(
                [
                    master_next_pct_formula(
                        r,
                        c_step=c_step,
                        c_active=c_active,
                        c_pct=c_pct,
                        c_before=c_before,
                    )
                ]
            )

    q = title.replace("'", "''")
    data = [
        {"range": f"'{q}'!{_col_index_to_a1(idx[PROMO_POINT_YEN_COL])}2", "values": yen_col},
        {"range": f"'{q}'!{_col_index_to_a1(idx[EFFECTIVE_PRICE_COL])}2", "values": eff_col},
    ]
    n = len(yen_col) + len(eff_col)
    if has_next and next_col:
        data.append(
            {
                "range": f"'{q}'!{_col_index_to_a1(idx[TAPER_NEXT_PCT_COL])}2",
                "values": next_col,
            }
        )
        n += len(next_col)
    svc.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"valueInputOption": "USER_ENTERED", "data": data},
    ).execute()
    return n
