# -*- coding: utf-8 -*-
"""
マスタの競合ASIN／URL／参考画像から「事実参照」画像を解決し、01ベースと同一商品か判定する。

正本ソース（優先順）:
  0. 手置き `06.競合事実参照/{parentSku}/`（任意）
  1. ▼マスタ(参考情報(画像URL)) … 卸サイト等の同一商品写真（物理事実）
  2. 競合店ASINコード → Keepa（`KEEPA_API_KEY` / secrets/keepa_api_key.txt）
  3. 競合AmazonページURL / amazon競合URL → og:image / Keepa
  （ASIN貼り付けKeepa用シートは使わない＝別商品混入リスク）

Vision 一致ゲート未達・取得失敗時は参照なし（フォールバック）。
"""
from __future__ import annotations

import json
import logging
import re
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from master_sets import PARENT_HINTS, CHILD_HINTS, _col, _norm, _read_table
from work_paths import default_work_root

LOG = logging.getLogger("set_main_image.competitor_fact")

COMP_ASIN_HINTS = ("競合店ASINコード", "競合ASIN")
COMP_URL_HINTS = (
    "競合AmazonページURL",
    "amazon競合URL",
    "競合URLAmazon",
)
# 同一商品の実物写真候補（卸サイト等）。デザインコピー禁止・物理事実の照合用。
REF_IMAGE_HINTS = (
    "▼マスタ(参考情報(画像URL))",
    "参考情報(画像URL)",
    "参考画像URL",
)
ASIN_RE = re.compile(r"\b(B0[A-Z0-9]{8})\b", re.I)

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)

# Keepa と同様の厳しさに近いゲート
DEFAULT_OVERALL_MIN = 70
DEFAULT_SHAPE_MIN = 55
DEFAULT_COLOR_MIN = 55


@dataclass
class FactMatchScore:
    shape: int = 0
    color: int = 0
    text: int = 0
    package: int = 0
    capacity: int = 0
    overall: int = 0


@dataclass
class CompetitorFactResult:
    used: bool
    path: Optional[Path] = None
    asin: str = ""
    source_url: str = ""
    source: str = ""  # master_asin | master_url | keepa | og_image | direct_image
    skip_reason: str = ""
    match: Optional[FactMatchScore] = None
    candidates_tried: List[Dict[str, Any]] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        d = asdict(self)
        d["path"] = str(self.path) if self.path else None
        return d


def _keepa_key() -> str:
    import os

    for env in ("KEEPA_API_KEY", "KEEPA_KEY"):
        v = (os.environ.get(env) or "").strip()
        if v:
            return v
    for name in ("keepa_api_key.txt", "KEEPA_API_KEY.txt"):
        p = Path(__file__).resolve().parent / "secrets" / name
        if p.is_file():
            t = p.read_text(encoding="utf-8").strip()
            if t and not t.startswith("#"):
                return t
    return ""


def extract_asin(text: str) -> str:
    s = _norm(text)
    if re.fullmatch(r"B0[A-Z0-9]{8}", s, re.I):
        return s.upper()
    m = ASIN_RE.search(s)
    return m.group(1).upper() if m else ""


def read_competitor_ids_from_master(
    master_csv: Path, parent_sku: str
) -> Dict[str, Any]:
    """親行（子SKU空）優先。無ければ同親の先頭行。"""
    rows, header_i, idx = _read_table(master_csv)
    i_p = _col(idx, PARENT_HINTS)
    i_c = _col(idx, CHILD_HINTS)
    i_asin = _col(idx, COMP_ASIN_HINTS)
    i_url = _col(idx, COMP_URL_HINTS)
    i_ref = _col(idx, REF_IMAGE_HINTS)
    want = _norm(parent_sku)
    parent_row = None
    any_row = None
    for row in rows[header_i + 1 :]:
        if i_p is None or _norm(row[i_p]) != want:
            continue
        any_row = row
        child = _norm(row[i_c]) if i_c is not None else ""
        if not child:
            parent_row = row
            break
    row = parent_row or any_row
    if row is None:
        return {"asins": [], "urls": [], "refImageUrls": [], "note": "parent not found"}

    asins: List[str] = []
    urls: List[str] = []
    ref_imgs: List[str] = []
    if i_asin is not None:
        raw = _norm(row[i_asin])
        for part in re.split(r"[,;/|\s]+", raw):
            a = extract_asin(part)
            if a and a not in asins:
                asins.append(a)
        a2 = extract_asin(raw)
        if a2 and a2 not in asins:
            asins.append(a2)
    if i_url is not None:
        raw_u = _norm(row[i_url])
        if raw_u:
            urls.append(raw_u)
            a3 = extract_asin(raw_u)
            if a3 and a3 not in asins:
                asins.append(a3)
    if i_ref is not None:
        raw_r = _norm(row[i_ref])
        for part in re.split(r"[\s,;]+", raw_r):
            if part.startswith("http"):
                ref_imgs.append(part)
    return {
        "asins": asins,
        "urls": urls,
        "refImageUrls": ref_imgs,
        "asinCol": i_asin,
        "urlCol": i_url,
        "refImageCol": i_ref,
    }


def list_local_fact_images(work_root: Path, parent_sku: str) -> List[Path]:
    """06.競合事実参照/{parentSku}/ または 06直下の手置き画像。"""
    root = work_root / "06.競合事実参照"
    out: List[Path] = []
    for folder in (root / (parent_sku or "_"), root):
        if not folder.is_dir():
            continue
        for p in sorted(folder.iterdir()):
            if p.is_file() and p.suffix.lower() in (
                ".jpg",
                ".jpeg",
                ".png",
                ".webp",
            ):
                if p.parent.name.startswith("_"):
                    continue
                out.append(p)
    return out


def _http_get(url: str, timeout: int = 25) -> Tuple[int, bytes, str]:
    req = urllib.request.Request(
        url, headers={"User-Agent": UA, "Accept": "*/*"}
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        ctype = resp.headers.get("Content-Type") or ""
        return resp.status, resp.read(), ctype


def _is_image_url(url: str) -> bool:
    u = url.lower().split("?")[0]
    return any(u.endswith(ext) for ext in (".jpg", ".jpeg", ".png", ".webp", ".gif"))


def _amazon_image_from_token(tok: str) -> str:
    s = tok.strip()
    if s.startswith("http"):
        return re.sub(r"\._AC_SL\d+_", "._AC_SL1500_", s, flags=re.I)
    if re.search(r"\.(jpe?g|png|gif|webp)$", s, re.I):
        return f"https://m.media-amazon.com/images/I/{s}"
    return f"https://m.media-amazon.com/images/I/{s}._AC_SL1500_.jpg"


def fetch_keepa_image_urls(asin: str) -> List[str]:
    key = _keepa_key()
    if not key:
        return []
    # domain=5 Japan
    q = urllib.parse.urlencode(
        {"key": key, "domain": "5", "asin": asin, "stats": "0"}
    )
    url = f"https://api.keepa.com/product?{q}"
    try:
        code, body, _ = _http_get(url, timeout=40)
        if code != 200:
            LOG.warning("Keepa HTTP %s asin=%s", code, asin)
            return []
        data = json.loads(body.decode("utf-8", errors="replace"))
        products = data.get("products") or []
        if not products:
            return []
        p = products[0]
        urls: List[str] = []
        seen = set()

        def add(u: str) -> None:
            uu = _amazon_image_from_token(u)
            if uu and uu not in seen:
                seen.add(uu)
                urls.append(uu)

        if p.get("image"):
            add(str(p["image"]))
        for img in p.get("images") or []:
            if isinstance(img, dict):
                if img.get("l"):
                    add(str(img["l"]))
                elif img.get("m"):
                    add(str(img["m"]))
            elif isinstance(img, str) and img.strip():
                add(img)
        if p.get("imagesCSV"):
            for part in str(p["imagesCSV"]).split(","):
                if part.strip():
                    add(part.strip())
        LOG.info("Keepa images asin=%s n=%s", asin, len(urls))
        return urls
    except Exception as e:
        LOG.warning("Keepa fetch failed asin=%s: %s", asin, e)
        return []


def fetch_og_image(page_url: str) -> Optional[str]:
    try:
        code, body, _ = _http_get(page_url, timeout=25)
        if code != 200:
            return None
        text = body.decode("utf-8", errors="replace")
        m = re.search(
            r'<meta[^>]+property=["\']og:image["\'][^>]+content=["\']([^"\']+)["\']',
            text,
            re.I,
        )
        if not m:
            m = re.search(
                r'<meta[^>]+content=["\']([^"\']+)["\'][^>]+property=["\']og:image["\']',
                text,
                re.I,
            )
        if m:
            return m.group(1).strip()
    except Exception as e:
        LOG.warning("og:image failed %s: %s", page_url[:60], e)
    return None


def download_image(url: str, dest: Path) -> bool:
    try:
        code, body, ctype = _http_get(url, timeout=30)
        if code != 200 or len(body) < 500:
            LOG.warning("download fail HTTP=%s bytes=%s url=%s", code, len(body), url[:80])
            return False
        if "html" in (ctype or "").lower() and body[:200].lstrip().lower().startswith(
            b"<!doctype"
        ):
            LOG.warning("download got HTML not image url=%s", url[:80])
            return False
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(body)
        return True
    except Exception as e:
        LOG.warning("download exception %s: %s", url[:80], e)
        return False


def match_product_images(
    base_path: Path, candidate_path: Path
) -> Optional[FactMatchScore]:
    """01ベース vs 競合画像。Gemini Vision（テキストモデル）。"""
    from gemini_image import load_api_key, make_client
    import base64

    try:
        client = make_client(load_api_key())
    except SystemExit:
        LOG.warning("Gemini key missing — cannot match competitor fact image")
        return None

    def b64(p: Path) -> Tuple[str, str]:
        data = p.read_bytes()
        mime = "image/png" if p.suffix.lower() == ".png" else "image/jpeg"
        return base64.b64encode(data).decode("ascii"), mime

    b1, m1 = b64(base_path)
    b2, m2 = b64(candidate_path)
    prompt = (
        "2枚の商品画像です。背景は無視し、商品部分のみを比較してください。\n"
        "以下の5項目それぞれについて、一致度を0～100の整数で答えてください。\n"
        "1. shape: 商品の形・シルエット\n"
        "2. color: 色・配色\n"
        "3. text: パッケージやラベルの文字・ロゴ\n"
        "4. package: パッケージ形状（袋/缶/ボトル/箱等）\n"
        "5. capacity: 容量・本数らしき表記\n"
        '回答はJSONのみ。例: {"shape":80,"color":70,"text":60,"package":90,"capacity":50}'
    )
    # 画像モデルではなく通常 Flash 系で比較（新アカウント向けモデルを先に）
    model_candidates = [
        "gemini-flash-latest",
        "gemini-2.5-flash",
        "gemini-2.0-flash",
    ]
    last_err = None
    for mid in model_candidates:
        try:
            resp = client.models.generate_content(
                model=mid,
                contents=[
                    {
                        "role": "user",
                        "parts": [
                            {"text": prompt},
                            {"inline_data": {"mime_type": m1, "data": b1}},
                            {"inline_data": {"mime_type": m2, "data": b2}},
                        ],
                    }
                ],
            )
            text = getattr(resp, "text", None) or ""
            if not text and getattr(resp, "candidates", None):
                # fallback extract
                try:
                    text = resp.candidates[0].content.parts[0].text
                except Exception:
                    text = ""
            score = _parse_score(text)
            if score:
                LOG.info(
                    "fact match model=%s overall=%s shape=%s color=%s",
                    mid,
                    score.overall,
                    score.shape,
                    score.color,
                )
                return score
        except Exception as e:
            last_err = e
            LOG.warning("vision match model=%s failed: %s", mid, e)
    LOG.warning("vision match all models failed last=%s", last_err)
    return None


def _parse_score(text: str) -> Optional[FactMatchScore]:
    if not text:
        return None
    t = re.sub(r"```\w*\n?", "", text).strip()
    start = t.find("{")
    end = t.rfind("}")
    if start < 0 or end <= start:
        return None
    try:
        o = json.loads(t[start : end + 1])
    except json.JSONDecodeError:
        return None

    def n(key: str, default: int = 50) -> int:
        try:
            return max(0, min(100, int(o.get(key, default))))
        except (TypeError, ValueError):
            return default

    shape, color, text_s = n("shape"), n("color"), n("text")
    package, capacity = n("package"), n("capacity")
    overall = round(
        shape * 0.25
        + color * 0.25
        + text_s * 0.2
        + package * 0.15
        + capacity * 0.15
    )
    return FactMatchScore(
        shape=shape,
        color=color,
        text=text_s,
        package=package,
        capacity=capacity,
        overall=overall,
    )


def _passes_gate(score: FactMatchScore) -> bool:
    if score.overall >= DEFAULT_OVERALL_MIN:
        return True
    # Keepa風フロア: shape/color が一定以上なら overall を救済しないが、両方高ければ通す
    if score.shape >= 70 and score.package >= 70 and score.color >= DEFAULT_COLOR_MIN:
        return True
    return False


def resolve_competitor_fact_image(
    *,
    master_csv: Optional[Path],
    parent_sku: str,
    base_product_path: Path,
    work_root: Optional[Path] = None,
    force_refresh: bool = False,
) -> CompetitorFactResult:
    """
    マスタ競合 → 画像取得 → 01ベース一致判定。
    合格時のみ path を返す。
    """
    root = work_root or default_work_root()
    cache_dir = root / "06.競合事実参照" / "_cache" / (parent_sku or "_unknown")
    cache_dir.mkdir(parents=True, exist_ok=True)

    if not master_csv or not Path(master_csv).is_file():
        return CompetitorFactResult(
            used=False, skip_reason="master_csv missing"
        )

    meta = read_competitor_ids_from_master(Path(master_csv), parent_sku)
    asins: List[str] = list(meta.get("asins") or [])
    urls: List[str] = list(meta.get("urls") or [])
    ref_imgs: List[str] = list(meta.get("refImageUrls") or [])

    # 0) 手置き 06.競合事実参照
    tried: List[Dict[str, Any]] = []
    for lp in list_local_fact_images(root, parent_sku):
        score = match_product_images(base_product_path, lp)
        entry = {
            "url": str(lp),
            "asin": "",
            "source": "local_06",
            "path": str(lp),
            "match": asdict(score) if score else None,
        }
        tried.append(entry)
        if score and _passes_gate(score):
            return CompetitorFactResult(
                used=True,
                path=lp,
                asin="",
                source_url=str(lp),
                source="local_06",
                match=score,
                candidates_tried=tried,
            )

    if not asins and not urls and not ref_imgs and not tried:
        return CompetitorFactResult(
            used=False, skip_reason="no competitor ASIN/URL/ref image on master"
        )

    # 候補URL列挙（参考画像URLを最優先）
    candidate_urls: List[Tuple[str, str, str]] = []  # url, asin, source
    for u in ref_imgs:
        candidate_urls.append((u, extract_asin(u), "master_ref_image"))
    for u in urls:
        if _is_image_url(u):
            candidate_urls.append((u, extract_asin(u), "direct_image"))
        else:
            a = extract_asin(u)
            og = fetch_og_image(u)
            if og:
                candidate_urls.append((og, a, "og_image"))
            if a:
                for ku in fetch_keepa_image_urls(a)[:3]:
                    candidate_urls.append((ku, a, "keepa"))
    for a in asins:
        for ku in fetch_keepa_image_urls(a)[:3]:
            candidate_urls.append((ku, a, "keepa"))
        # Keepa無し時の保険: 商品ページ og:image
        if not _keepa_key():
            page = f"https://www.amazon.co.jp/dp/{a}"
            og = fetch_og_image(page)
            if og:
                candidate_urls.append((og, a, "og_image"))

    # 重複除去
    seen_u = set()
    uniq: List[Tuple[str, str, str]] = []
    for u, a, src in candidate_urls:
        if u in seen_u:
            continue
        seen_u.add(u)
        uniq.append((u, a, src))

    for i, (u, a, src) in enumerate(uniq[:6]):
        dest = cache_dir / f"fact_{a or 'url'}_{i}.jpg"
        if force_refresh or not dest.is_file() or dest.stat().st_size < 500:
            ok = download_image(u, dest)
            if not ok:
                tried.append({"url": u, "asin": a, "source": src, "ok": False})
                continue
        # プレースホルダGIF等を除外
        if dest.stat().st_size < 1000:
            tried.append(
                {
                    "url": u,
                    "asin": a,
                    "source": src,
                    "ok": False,
                    "reason": "too_small",
                }
            )
            continue
        score = match_product_images(base_product_path, dest)
        entry = {
            "url": u,
            "asin": a,
            "source": src,
            "path": str(dest),
            "match": asdict(score) if score else None,
        }
        tried.append(entry)
        if score and _passes_gate(score):
            LOG.info(
                "competitor FACT accepted asin=%s source=%s overall=%s file=%s",
                a,
                src,
                score.overall,
                dest.name,
            )
            return CompetitorFactResult(
                used=True,
                path=dest,
                asin=a,
                source_url=u,
                source=src,
                match=score,
                candidates_tried=tried,
            )
        LOG.info(
            "competitor FACT rejected asin=%s overall=%s (different or low match)",
            a,
            score.overall if score else None,
        )

    reason = "no candidate passed same-product gate"
    if not uniq and not tried:
        reason = "no image URLs resolved from ASIN/URL (Keepa key? network?)"
    elif not uniq and tried:
        reason = "local candidates rejected by match gate"
    return CompetitorFactResult(
        used=False,
        skip_reason=reason,
        candidates_tried=tried,
        asin=asins[0] if asins else "",
    )
