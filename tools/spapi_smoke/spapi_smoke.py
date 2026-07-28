#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SP-API 疎通スモーク（読取のみ）

1) LWA refresh_token → access_token
2) Catalog Items API で ASIN 1件 GET（出品書込なし）

秘密は config.local.json（gitignore）または環境変数。
正本手順: docs/org/D_MENU_SPAPI_SMOKE_HUMAN_RUN.md
接続: https://developer-docs.amazon.com/sp-api/docs/connecting-to-the-selling-partner-api
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Optional
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


def get_catalog_item(
    endpoint: str,
    access_token: str,
    asin: str,
    marketplace_id: str,
    user_agent: str,
) -> requests.Response:
    path = "/catalog/2022-04-01/items/%s" % asin
    qs = urlencode(
        {
            "marketplaceIds": marketplace_id,
            "includedData": "summaries",
        }
    )
    url = "%s%s?%s" % (endpoint.rstrip("/"), path, qs)
    headers = {
        "host": endpoint.replace("https://", "").replace("http://", "").split("/")[0],
        "x-amz-access-token": access_token,
        "x-amz-date": datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ"),
        "user-agent": user_agent,
        "accept": "application/json",
    }
    return requests.get(url, headers=headers, timeout=60)


def run(config_path: Path) -> int:
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
    user_agent = str(cfg.get("user_agent") or "OctasSpapiSmoke/1.0 (Language=Python)").strip()

    report: Dict[str, Any] = {
        "runId": "SPAPI_SMOKE_%s" % _utc_stamp(),
        "version": "spapi-smoke-v1",
        "marketplaceId": marketplace_id,
        "endpoint": endpoint,
        "asin": asin,
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
    LOG.info("Step2 Catalog GET asin=%s …", asin)
    resp = get_catalog_item(endpoint, access_token, asin, marketplace_id, user_agent)
    report["steps"]["catalog"] = {
        "httpStatus": resp.status_code,
        "ok": resp.status_code == 200,
        "bodyPreview": resp.text[:800],
    }

    out_dir = SCRIPT_DIR / "out"
    out_dir.mkdir(parents=True, exist_ok=True)
    report_path = out_dir / ("%s_REPORT.json" % report["runId"])
    # 秘密をレポートに載せない
    with report_path.open("w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
        f.write("\n")
    LOG.info("レポート: %s", report_path)

    if resp.status_code != 200:
        LOG.error("Catalog 失敗 HTTP %s: %s", resp.status_code, resp.text[:500])
        return 2

    try:
        data = resp.json()
        summaries = (data.get("summaries") or [])
        title = ""
        if summaries:
            title = str(summaries[0].get("itemName") or summaries[0].get("title") or "")
        LOG.info("Catalog OK asin=%s title=%s", data.get("asin") or asin, title[:80])
    except Exception:
        LOG.info("Catalog OK HTTP 200（JSON解析スキップ）")

    return 0


def main(argv: Optional[list] = None) -> int:
    parser = argparse.ArgumentParser(description="SP-API 読取スモーク（LWA＋Catalog）")
    parser.add_argument(
        "--config",
        default=str(SCRIPT_DIR / "config.local.json"),
        help="config.local.json（gitignore）",
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
    return run(config_path)


if __name__ == "__main__":
    sys.exit(main())
