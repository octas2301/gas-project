# -*- coding: utf-8 -*-
"""O2: 通過5件。Keepa offers=20 と SP-API ItemOffers。Keepaフル非書。"""
from __future__ import annotations

import gzip
import json
import sys
from collections import Counter
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from apply_keepa_full import (
    COMPETITOR_SS,
    RESEARCH_SS,
    T_CAND,
    as_dicts,
    append_log,
    keepa_key,
    read_all,
    sheets_service,
)
from schema import SHEET_KEEPA_FULL

SMOKE = Path(__file__).resolve().parents[1] / "spapi_smoke"
sys.path.insert(0, str(SMOKE))


def ok(v) -> bool:
    if v is None or v == -1:
        return False
    try:
        return float(v) >= 0
    except (TypeError, ValueError):
        return False


def stats_counts(p: dict) -> dict:
    st = p.get("stats") if isinstance(p.get("stats"), dict) else {}
    cur = st.get("current") or []
    c11 = cur[11] if len(cur) > 11 else None
    return {
        "totalOfferCount": st.get("totalOfferCount"),
        "countNew": c11,
        "fba": st.get("offerCountFBA"),
        "fbm": st.get("offerCountFBM"),
    }


def pick5(cand: list[dict], full: list[dict]) -> list[str]:
    by = {}
    for r in full:
        a = str(r.get("ASIN") or "").upper()
        if a:
            by[a] = r
    zeros, highs, ones = [], [], []
    for c in cand:
        if str(c.get("門結果") or "") != "通過":
            continue
        a = str(c.get("ASIN") or "").upper()
        row = by.get(a)
        if not row:
            continue
        try:
            p = json.loads(row.get("生JSON") or "{}")
        except json.JSONDecodeError:
            continue
        t = stats_counts(p).get("totalOfferCount")
        try:
            n = int(float(t)) if ok(t) else -1
        except (TypeError, ValueError):
            n = -1
        if n == 0:
            zeros.append(a)
        elif n == 1:
            ones.append(a)
        elif n >= 2:
            highs.append(a)
    out = []
    for bucket in (highs, ones, zeros):
        for a in bucket:
            if a not in out:
                out.append(a)
            if len(out) >= 5:
                return out
    return out[:5]


def fetch_keepa_offers(key: str, asins: list[str]) -> dict:
    q = urlencode(
        {
            "key": key,
            "domain": "5",
            "asin": ",".join(asins),
            "stats": "90",
            "history": "0",
            "offers": "20",
        }
    )
    url = "https://api.keepa.com/product?" + q
    req = Request(url, headers={"User-Agent": "OctasO2/1.0", "Accept-Encoding": "identity"})
    with urlopen(req, timeout=90) as resp:
        raw = resp.read()
    if len(raw) >= 2 and raw[0] == 0x1F and raw[1] == 0x8B:
        raw = gzip.decompress(raw)
    return json.loads(raw.decode("utf-8"))


def count_keepa_offers(p: dict) -> dict:
    offers = p.get("offers") or []
    if not isinstance(offers, list):
        offers = []
    sids = []
    new_sids = []
    used_sids = []
    for o in offers:
        if not isinstance(o, dict):
            continue
        sid = str(o.get("sellerId") or "").strip()
        if not sid:
            continue
        sids.append(sid)
        cond = o.get("condition")
        is_used = o.get("isUsed") is True or cond in (2, 3, 4, 5, 6)
        if is_used:
            used_sids.append(sid)
        else:
            new_sids.append(sid)
    return {
        "n_offers": len(offers),
        "uniq": len(set(sids)),
        "uniq_new": len(set(new_sids)),
        "uniq_used": len(set(used_sids)),
    }


def spapi_session():
    from spapi_smoke import _load_config, _spapi_headers, request_lwa_access_token, resolve_cred

    cfg = _load_config(SMOKE / "config.local.json")
    token = request_lwa_access_token(
        resolve_cred(cfg, "lwa_client_id", "SPAPI_LWA_CLIENT_ID"),
        resolve_cred(cfg, "lwa_client_secret", "SPAPI_LWA_CLIENT_SECRET"),
        resolve_cred(cfg, "refresh_token", "SPAPI_REFRESH_TOKEN"),
    )
    return {
        "access": str(token.get("access_token") or ""),
        "marketplace_id": str(cfg.get("marketplace_id") or "A1VC38T7YXB528"),
        "endpoint": str(cfg.get("endpoint") or "https://sellingpartnerapi-fe.amazon.com"),
        "ua": str(cfg.get("user_agent") or "OctasO2/1.0 (Language=Python)"),
    }


def spapi_item_offers(asin: str, sess: dict) -> dict:
    import requests
    from spapi_smoke import _spapi_headers

    path = "/products/pricing/v0/items/%s/offers" % asin
    qs = urlencode({"MarketplaceId": sess["marketplace_id"], "ItemCondition": "New"})
    url = "%s%s?%s" % (sess["endpoint"].rstrip("/"), path, qs)
    resp = requests.get(url, headers=_spapi_headers(sess["endpoint"], sess["access"], sess["ua"]), timeout=60)
    out = {"http": resp.status_code, "total": None, "err": ""}
    try:
        body = resp.json()
    except Exception:
        out["err"] = (resp.text or "")[:180]
        return out
    if resp.status_code != 200:
        out["err"] = str(body.get("errors") or body)[:180]
        return out
    payload = body.get("payload") or body
    summary = payload.get("Summary") or payload.get("summary") or {}
    tot = summary.get("TotalOfferCount")
    if tot is None:
        tot = summary.get("totalOfferCount")
    out["total"] = tot
    return out


def main() -> int:
    svc = sheets_service(write=True)
    _, cand = as_dicts(read_all(svc, RESEARCH_SS, T_CAND))
    _, full = as_dicts(read_all(svc, COMPETITOR_SS, SHEET_KEEPA_FULL))
    asins = pick5(cand, full)
    print("asins", asins)
    if len(asins) < 5:
        print("need 5 pass asins, got", len(asins))
        return 2
    by = {str(r.get("ASIN") or "").upper(): r for r in full}
    stored = {}
    for a in asins:
        stored[a] = stats_counts(json.loads(by[a].get("生JSON") or "{}"))
        print("stored", a, stored[a])
    key = keepa_key()
    data = fetch_keepa_offers(key, asins)
    print("keepa consumed", data.get("tokensConsumed"), "left", data.get("tokensLeft"))
    products = {str(p.get("asin") or "").upper(): p for p in (data.get("products") or [])}
    sess = spapi_session()
    rows = []
    for a in asins:
        p = products.get(a) or {}
        live_st = stats_counts(p)
        ko = count_keepa_offers(p)
        sp = spapi_item_offers(a, sess)
        rec = {
            "asin": a,
            "stored_total": stored[a].get("totalOfferCount"),
            "stored_c11": stored[a].get("countNew"),
            "live_total": live_st.get("totalOfferCount"),
            "live_c11": live_st.get("countNew"),
            "live_fba": live_st.get("fba"),
            "keepa_uniq_new": ko["uniq_new"],
            "keepa_uniq": ko["uniq"],
            "keepa_n_offers": ko["n_offers"],
            "sp_http": sp["http"],
            "sp_total": sp["total"],
            "sp_err": sp.get("err") or "",
        }
        rows.append(rec)
        print(rec)
    match_keepa = sum(
        1
        for r in rows
        if ok(r["stored_c11"]) and int(float(r["stored_c11"])) == int(r["keepa_uniq_new"] or -1)
    )
    match_sp = sum(
        1
        for r in rows
        if r["sp_total"] is not None and ok(r["stored_c11"]) and int(float(r["stored_c11"])) == int(r["sp_total"])
    )
    line = (
        "runId=pr_20260815_o2 n=5 keepaConsumed=%s left=%s match_storedC11_vs_keepaUniqNew=%s match_c11_vs_sp=%s sp_http=%s"
        % (
            data.get("tokensConsumed"),
            data.get("tokensLeft"),
            match_keepa,
            match_sp,
            Counter(r["sp_http"] for r in rows).most_common(),
        )
    )
    append_log(svc, "O2", line)
    print(line)
    return 0


if __name__ == "__main__":
    sys.exit(main())
