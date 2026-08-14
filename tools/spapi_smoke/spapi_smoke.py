#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SP-API 疎通スモーク（読取のみ）

1) LWA refresh_token → access_token
2) Catalog Items API で ASIN 1件 GET（出品書込なし）
3) （任意）--poc-category: Product Type Definitions 検索／定義取得＋Catalog分類系

秘密は config.local.json（gitignore）または環境変数。
正本手順: docs/org/D_MENU_SPAPI_SMOKE_HUMAN_RUN.md
P4a: docs/org/LV4_AMAZON_CATEGORY_PT_POC_HUMAN_RUN.md
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional
from urllib.parse import urlencode

LOG = logging.getLogger("spapi_smoke")
SCRIPT_DIR = Path(__file__).resolve().parent

try:
    import requests
except ImportError as e:  # pragma: no cover
    raise SystemExit("requests が必要です: pip install -r requirements.txt") from e

LWA_TOKEN_URL = "https://api.amazon.com/auth/o2/token"


def _utc_stamp() -> str:
    return datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")


def _load_config(path: Path) -> Dict[str, Any]:
    with path.open(encoding="utf-8") as f:
        return json.load(f)


def resolve_cred(cfg: Dict[str, Any], key: str, env_name: str) -> str:
    v = str(os.environ.get(env_name) or cfg.get(key) or "").strip()
    if not v or v.startswith("REPLACE"):
        raise SystemExit(
            "認証情報が未設定です: config.%s または環境変数 %s" % (key, env_name)
        )
    return v


def request_lwa_access_token(client_id: str, client_secret: str, refresh_token: str) -> Dict[str, Any]:
    data = {
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "client_id": client_id,
        "client_secret": client_secret,
    }
    resp = requests.post(
        LWA_TOKEN_URL,
        data=data,
        headers={"Content-Type": "application/x-www-form-urlencoded;charset=UTF-8"},
        timeout=60,
    )
    if resp.status_code != 200:
        raise RuntimeError(
            "LWA token 失敗 HTTP %s body=%s" % (resp.status_code, resp.text[:500])
        )
    body = resp.json()
    if not body.get("access_token"):
        raise RuntimeError("LWA 応答に access_token がありません: %s" % body)
    return body


def _spapi_headers(endpoint: str, access_token: str, user_agent: str) -> Dict[str, str]:
    return {
        "host": endpoint.replace("https://", "").replace("http://", "").split("/")[0],
        "x-amz-access-token": access_token,
        "x-amz-date": datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ"),
        "user-agent": user_agent,
        "accept": "application/json",
    }


def get_catalog_item(
    endpoint: str,
    access_token: str,
    asin: str,
    marketplace_id: str,
    user_agent: str,
    included_data: str = "summaries",
) -> requests.Response:
    path = "/catalog/2022-04-01/items/%s" % asin
    qs = urlencode(
        {
            "marketplaceIds": marketplace_id,
            "includedData": included_data,
        }
    )
    url = "%s%s?%s" % (endpoint.rstrip("/"), path, qs)
    return requests.get(
        url,
        headers=_spapi_headers(endpoint, access_token, user_agent),
        timeout=60,
    )


def search_definitions_product_types(
    endpoint: str,
    access_token: str,
    marketplace_id: str,
    user_agent: str,
    keywords: str = "",
    item_name: str = "",
) -> requests.Response:
    path = "/definitions/2020-09-01/productTypes"
    params: Dict[str, str] = {"marketplaceIds": marketplace_id}
    if keywords:
        params["keywords"] = keywords
    if item_name:
        params["itemName"] = item_name
    url = "%s%s?%s" % (endpoint.rstrip("/"), path, urlencode(params))
    return requests.get(
        url,
        headers=_spapi_headers(endpoint, access_token, user_agent),
        timeout=60,
    )


def get_definitions_product_type(
    endpoint: str,
    access_token: str,
    marketplace_id: str,
    user_agent: str,
    product_type: str,
    locale: str = "ja_JP",
) -> requests.Response:
    path = "/definitions/2020-09-01/productTypes/%s" % product_type
    qs = urlencode(
        {
            "marketplaceIds": marketplace_id,
            "requirements": "LISTING",
            "locale": locale,
        }
    )
    url = "%s%s?%s" % (endpoint.rstrip("/"), path, qs)
    return requests.get(
        url,
        headers=_spapi_headers(endpoint, access_token, user_agent),
        timeout=90,
    )


def _preview(text: str, n: int = 800) -> str:
    """Truncate and strip query/signature fragments from previews (no secrets in reports)."""
    import re

    s = text or ""
    s = re.sub(r"[?&](X-Amz-[^=]+=[^&\s\"]+)", "", s)
    s = re.sub(r"Signature=[^&\s\"]+", "Signature=REDACTED", s)
    return s[:n]


def _extract_product_type_names(body: Any) -> List[str]:
    names: List[str] = []
    if not isinstance(body, dict):
        return names
    for key in ("productTypes", "ProductTypes"):
        arr = body.get(key)
        if isinstance(arr, list):
            for one in arr:
                if isinstance(one, dict):
                    name = one.get("name") or one.get("productType") or one.get("marketplaceId")
                    if name:
                        names.append(str(name))
                elif one:
                    names.append(str(one))
    return names


def run_category_pt_poc(
    endpoint: str,
    access_token: str,
    marketplace_id: str,
    user_agent: str,
    asin: str,
    cfg: Dict[str, Any],
) -> Dict[str, Any]:
    """P4a: Definitions search/get + Catalog classifications (read-only)."""
    keywords = str(cfg.get("keywords") or "").strip()
    item_name = str(cfg.get("item_name") or "").strip()
    forced_pt = str(cfg.get("product_type") or "").strip()
    if not keywords and not item_name:
        keywords = "調味料"

    out: Dict[str, Any] = {
        "keywords": keywords,
        "itemName": item_name,
        "conclusions": {},
    }

    LOG.info("P4a-1 searchDefinitionsProductTypes keywords=%s itemName=%s", keywords, item_name)
    resp_search = search_definitions_product_types(
        endpoint, access_token, marketplace_id, user_agent, keywords, item_name
    )
    search_ok = resp_search.status_code == 200
    search_json: Any = None
    try:
        search_json = resp_search.json() if search_ok else None
    except Exception:
        search_json = None
    pt_names = _extract_product_type_names(search_json) if search_json else []
    out["searchDefinitionsProductTypes"] = {
        "httpStatus": resp_search.status_code,
        "ok": search_ok,
        "productTypeNames": pt_names[:20],
        "bodyPreview": _preview(resp_search.text),
    }
    out["conclusions"]["1_searchDefinitions"] = (
        "OK candidates=%s" % len(pt_names) if search_ok else "FAIL HTTP %s" % resp_search.status_code
    )
    LOG.info("search OK=%s count=%s sample=%s", search_ok, len(pt_names), pt_names[:5])

    # Prefer SEASONING when present (food keywords often return HERB first)
    if forced_pt:
        product_type = forced_pt
    elif "SEASONING" in pt_names:
        product_type = "SEASONING"
    else:
        product_type = pt_names[0] if pt_names else ""
    if product_type:
        LOG.info("P4a-2 getDefinitionsProductType productType=%s", product_type)
        resp_def = get_definitions_product_type(
            endpoint, access_token, marketplace_id, user_agent, product_type
        )
        def_ok = resp_def.status_code == 200
        schema_keys: List[str] = []
        try:
            if def_ok:
                dj = resp_def.json()
                if isinstance(dj, dict):
                    schema_keys = sorted(list(dj.keys()))[:40]
        except Exception:
            pass
        out["getDefinitionsProductType"] = {
            "productType": product_type,
            "httpStatus": resp_def.status_code,
            "ok": def_ok,
            "topLevelKeys": schema_keys,
            "bodyPreview": _preview(resp_def.text, 1200),
        }
        out["conclusions"]["2_getDefinitions"] = (
            "OK productType=%s keys=%s" % (product_type, len(schema_keys))
            if def_ok
            else "FAIL HTTP %s" % resp_def.status_code
        )
    else:
        out["getDefinitionsProductType"] = {"skipped": True, "reason": "no productType candidate"}
        out["conclusions"]["2_getDefinitions"] = "SKIP no productType"

    included = "summaries,productTypes,classifications,attributes"
    LOG.info("P4a-3 Catalog GET asin=%s includedData=%s", asin, included)
    resp_cat = get_catalog_item(
        endpoint, access_token, asin, marketplace_id, user_agent, included_data=included
    )
    cat_ok = resp_cat.status_code == 200
    cat_summary: Dict[str, Any] = {"httpStatus": resp_cat.status_code, "ok": cat_ok}
    if cat_ok:
        try:
            cj = resp_cat.json()
            cat_summary["asin"] = cj.get("asin") or asin
            cat_summary["hasSummaries"] = bool(cj.get("summaries"))
            cat_summary["hasProductTypes"] = bool(cj.get("productTypes"))
            cat_summary["hasClassifications"] = bool(cj.get("classifications"))
            cat_summary["hasAttributes"] = bool(cj.get("attributes"))
            if cj.get("productTypes"):
                cat_summary["productTypesPreview"] = cj.get("productTypes")[:3]
            if cj.get("classifications"):
                cat_summary["classificationsPreview"] = cj.get("classifications")[:3]
            if cj.get("summaries"):
                s0 = cj["summaries"][0] if isinstance(cj["summaries"], list) and cj["summaries"] else {}
                cat_summary["itemName"] = str(
                    (s0 or {}).get("itemName") or (s0 or {}).get("title") or ""
                )[:120]
        except Exception as e:
            cat_summary["parseError"] = str(e)
    cat_summary["bodyPreview"] = _preview(resp_cat.text)
    out["catalogClassifications"] = cat_summary
    out["conclusions"]["3_catalog"] = (
        "OK types=%s class=%s attrs=%s"
        % (
            cat_summary.get("hasProductTypes"),
            cat_summary.get("hasClassifications"),
            cat_summary.get("hasAttributes"),
        )
        if cat_ok
        else "FAIL HTTP %s" % resp_cat.status_code
    )

    out["conclusions"]["4_keepa"] = (
        "MANUAL: メニューAで同ASINを少件取得しカテゴリ系フィールド有無を記録（token消費）"
    )
    out["conclusions"]["5_xlsm_auto_dl"] = (
        "CONCLUSION: SP-API Product Type Definitions は JSON スキーマ。"
        "純正 Seller Central .xlsm テンプレの自動DL API は本経路に無い。"
        "代替は Definitions JSON／既存C1手運用。"
    )
    return out


def run(config_path: Path, poc_category: bool = False) -> int:
    cfg = _load_config(config_path)
    client_id = resolve_cred(cfg, "lwa_client_id", "SPAPI_LWA_CLIENT_ID")
    client_secret = resolve_cred(cfg, "lwa_client_secret", "SPAPI_LWA_CLIENT_SECRET")
    refresh_token = resolve_cred(cfg, "refresh_token", "SPAPI_REFRESH_TOKEN")
    marketplace_id = str(
        os.environ.get("SPAPI_MARKETPLACE_ID")
        or cfg.get("marketplace_id")
        or "A1VC38T7YXB528"
    ).strip()
    endpoint = str(
        os.environ.get("SPAPI_ENDPOINT")
        or cfg.get("endpoint")
        or "https://sellingpartnerapi-fe.amazon.com"
    ).strip()
    asin = str(cfg.get("smoke_asin") or "B07YND44VN").strip()
    user_agent = str(
        cfg.get("user_agent")
        or ("OctasSpapiCategoryPt/1.0 (Language=Python)" if poc_category else "OctasSpapiSmoke/1.0 (Language=Python)")
    ).strip()

    report: Dict[str, Any] = {
        "runId": ("SPAPI_CATEGORY_PT_%s" % _utc_stamp()) if poc_category else ("SPAPI_SMOKE_%s" % _utc_stamp()),
        "version": "spapi-category-pt-v1" if poc_category else "spapi-smoke-v1",
        "marketplaceId": marketplace_id,
        "endpoint": endpoint,
        "asin": asin,
        "pocCategory": poc_category,
        "steps": {},
    }

    LOG.info("Step1 LWA access_token 取得…")
    token_body = request_lwa_access_token(client_id, client_secret, refresh_token)
    expires_in = token_body.get("expires_in")
    report["steps"]["lwa"] = {
        "ok": True,
        "expires_in": expires_in,
        "token_type": token_body.get("token_type"),
        "access_token_len": len(str(token_body.get("access_token") or "")),
    }
    LOG.info("LWA OK expires_in=%s", expires_in)

    access_token = str(token_body["access_token"])

    if poc_category:
        report["steps"]["categoryPt"] = run_category_pt_poc(
            endpoint, access_token, marketplace_id, user_agent, asin, cfg
        )
    else:
        LOG.info("Step2 Catalog GET asin=%s …", asin)
        resp = get_catalog_item(endpoint, access_token, asin, marketplace_id, user_agent)
        report["steps"]["catalog"] = {
            "httpStatus": resp.status_code,
            "ok": resp.status_code == 200,
            "bodyPreview": _preview(resp.text),
        }
        if resp.status_code != 200:
            LOG.error("Catalog 失敗 HTTP %s: %s", resp.status_code, resp.text[:500])
            _write_report(report)
            return 2
        try:
            data = resp.json()
            summaries = data.get("summaries") or []
            title = ""
            if summaries:
                title = str(summaries[0].get("itemName") or summaries[0].get("title") or "")
            LOG.info("Catalog OK asin=%s title=%s", data.get("asin") or asin, title[:80])
        except Exception:
            LOG.info("Catalog OK HTTP 200（JSON解析スキップ）")

    _write_report(report)

    if poc_category:
        cat_pt = report["steps"].get("categoryPt") or {}
        conclusions = cat_pt.get("conclusions") or {}
        LOG.info("P4a conclusions: %s", json.dumps(conclusions, ensure_ascii=False))
        search_ok = bool((cat_pt.get("searchDefinitionsProductTypes") or {}).get("ok"))
        # Catalog拡張が権限で落ちても search 成功なら調査として前進可
        return 0 if search_ok else 2

    return 0


def _write_report(report: Dict[str, Any]) -> Path:
    out_dir = SCRIPT_DIR / "out"
    out_dir.mkdir(parents=True, exist_ok=True)
    report_path = out_dir / ("%s_REPORT.json" % report["runId"])
    with report_path.open("w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
        f.write("\n")
    LOG.info("レポート: %s", report_path)
    return report_path


def main(argv: Optional[list] = None) -> int:
    parser = argparse.ArgumentParser(description="SP-API 読取スモーク（LWA＋Catalog／任意P4a）")
    parser.add_argument(
        "--config",
        default=str(SCRIPT_DIR / "config.local.json"),
        help="config.local.json（gitignore）",
    )
    parser.add_argument(
        "--poc-category",
        action="store_true",
        help="P4a: Product Type Definitions 検索／定義＋Catalog分類系（読取のみ）",
    )
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
    )
    config_path = Path(args.config).expanduser().resolve()
    if not config_path.is_file():
        example = SCRIPT_DIR / "config.example.json"
        raise SystemExit(
            "config がありません: %s\n例: copy %s → config.local.json に秘密を記入"
            % (config_path, example)
        )
    return run(config_path, poc_category=bool(args.poc_category))


if __name__ == "__main__":
    sys.exit(main())
