# -*- coding: utf-8 -*-
"""
B-⑤ 作業用ファイル名（人が見る比較フォルダ）と 02 出口名。

作業用（日本語OK）:
  フォルダ  {商品名省略}_{YYYYMMDD_HHMMSS}
  比較文件  {商品名省略}_{themeId}_{キャッチ【】内}.jpg

出口（後続C）:
  02.楽天ベース/{親SKU}_単体.png                 … 互換1枚
  02.楽天ベース/{親SKU}_単体__{adoptKey}.png     … P3b 複数土台
  楽天UL は C が {子SKU}_rakuten.jpg（半角）— 本モジュールでは書かない
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Optional

# 正本ファイル名から落とす接頭（省略名のため）
_STRIP_PREFIXES = (
    "たべるおだし",
    "食べるおだし",
    "たべるお出汁",
    "食べるお出汁",
    "タベルオダシ",
)


def sanitize_fs_name(s: str, *, max_len: int = 40) -> str:
    """Windows 禁則を除去。空白は _ 。空なら 'item'。"""
    t = str(s or "").strip()
    t = re.sub(r'[\\/:*?"<>|]+', "", t)
    t = re.sub(r"\s+", "_", t)
    t = t.strip("._")
    if not t:
        return "item"
    if len(t) > max_len:
        t = t[:max_len].rstrip("._")
    return t or "item"


def derive_short_name(
    *,
    package_truth: Path | str,
    explicit: str = "",
    materials_name: str = "",
) -> str:
    """商品名省略。JSON shortName > 正本stemから接頭除去 > 材料名 > stem全体。"""
    ex = str(explicit or "").strip()
    if ex:
        return sanitize_fs_name(ex, max_len=24)

    stem = Path(package_truth).stem if package_truth else ""
    stem_n = re.sub(r"\s+", "", stem)
    for pref in _STRIP_PREFIXES:
        if stem_n.startswith(pref):
            rest = stem_n[len(pref) :].strip(" _-・")
            if rest:
                return sanitize_fs_name(rest, max_len=24)

    mat = re.sub(r"\s+", "", str(materials_name or ""))
    for pref in _STRIP_PREFIXES:
        if mat.startswith(pref):
            rest = mat[len(pref) :].strip(" _-・")
            if rest:
                # 内容量など後ろを落とす
                rest = re.split(r"[（(0-9]", rest, maxsplit=1)[0]
                if rest:
                    return sanitize_fs_name(rest, max_len=24)

    if stem_n:
        return sanitize_fs_name(stem_n, max_len=24)
    return "item"


def catch_head_ja(catch: str, *, max_len: int = 18) -> str:
    """【見出し】優先。無ければ冒頭。括弧まわりの空白は詰める。"""
    t = str(catch or "").strip()
    m = re.search(r"【([^】]+)】", t)
    head = (m.group(1) if m else t).strip()
    head = re.sub(r"[。．.！!？?]+$", "", head)
    # 空白除去のみ（括弧は残す。[] 内の ] 誤解釈を避ける）
    head = re.sub(r"\s+", "", head)
    return sanitize_fs_name(head, max_len=max_len)


def _sales_trim_ja(sales: str, *, sales_max: int = 16, soft_max: int = 20) -> str:
    """セールスを短く。硬切りで「食卓に贅」のような途中切れを出さない。

    §7.6ter: 目安 8〜16字。境界が取れるときだけ soft_max まで許容。
    """
    t = str(sales or "").strip()
    t = re.sub(r'[\\/:*?"<>|]+', "", t)
    t = re.sub(r"\s+", "", t)
    t = re.sub(r"[。．.！!？?]+$", "", t)
    if not t:
        return ""
    weak_end = set("にのをはがもへてでとやしり")
    if len(t) <= sales_max:
        return t
    if len(t) <= soft_max and t[-1] not in weak_end:
        return t

    # 「〜だけで、〜」: 前半が導入だけのときは後半を優先（意味が残る）
    if "、" in t or "，" in t:
        sep = "、" if "、" in t else "，"
        left, right = t.split(sep, 1)
        left, right = left.strip(), right.strip()
        if right and len(right) >= 4:
            if left.endswith(("だけで", "だけ", "するだけ")) and len(right) <= soft_max:
                return _sales_trim_ja(right, sales_max=sales_max, soft_max=soft_max)
            if 4 <= len(left) <= soft_max:
                return left

    for sep in ("。", "．", "!", "！", "?", "？"):
        if sep in t:
            cand = t.split(sep, 1)[0].strip()
            if 4 <= len(cand) <= soft_max:
                return cand

    window = t[:soft_max]
    # 意味の終わりを優先（長いマークから）
    for mark in ("だけで", "するだけ", "します", "ました", "です", "ます", "だけ"):
        idx = window.rfind(mark)
        if idx >= 3:
            cand = window[: idx + len(mark)]
            if 4 <= len(cand) <= soft_max:
                return cand

    # soft 内の助詞・読点で切る（sales_max 以下を優先、無ければ soft）
    # 「…に」「…の」「…で」だけの終わりは途中切れに見えるのでスキップ
    for limit in (sales_max, soft_max):
        w = t[:limit]
        for i in range(len(w) - 1, 3, -1):
            if w[i] in "でにをはがもへ、，":
                cand = w[: i + (0 if w[i] in "、，" else 1)]
                cand = re.sub(r"[、，]+$", "", cand)
                if len(cand) < 4:
                    continue
                if cand[-1] in weak_end:
                    continue
                return cand

    # 最後の手段: sales_max で切り、終わりの助詞・開き括弧を落とす
    cand = t[:sales_max]
    while len(cand) >= 4 and cand[-1] in weak_end:
        cand = cand[:-1]
    cand = re.sub(r"[、，（(【「『［\[]+$", "", cand)
    return cand if len(cand) >= 4 else ""


def split_catch_for_ui(catch: str) -> tuple[str, str, str]:
    """UI用: (キャッチ大, セールストーク小, 2文目以降).

    2文目以降は proposals の catch 全文から復元（再生成時に保持）。
    """
    full = str(catch or "").strip()
    if not full:
        return "", "", ""
    head, sales = catch_sales_pair_ja(full)
    extra = ""
    m = re.search(r"【([^】]+)】", full)
    if m:
        rest = full[m.end() :].strip()
        rest = re.sub(r"^[。．.、，,　\s]+", "", rest)
        if sales and rest.startswith(sales):
            tail = rest[len(sales) :].lstrip("。．.、，,　 ")
            if tail.startswith("。"):
                tail = tail[1:].strip()
            extra = tail
        elif rest and sales and sales not in rest:
            extra = rest
    return head, sales, extra


def merge_catch_for_regen(
    catch_head: str,
    sales_talk: str,
    *,
    original_catch: str = "",
    extra: str = "",
) -> str:
    """キャッチ大＋セールストーク小（＋任意の2文目）を proposals.catch 形式へ。"""
    head = re.sub(r"\s+", "", str(catch_head or "").strip())
    head = re.sub(r"^[【\[]+|[】\]]+$", "", head)
    sales = re.sub(r"\s+", "", str(sales_talk or "").strip())
    sales = re.sub(r"[。．.！!？?]+$", "", sales)
    tail = re.sub(r"\s+", "", str(extra or "").strip())
    if not tail and original_catch:
        _, _, tail = split_catch_for_ui(original_catch)
    if not head and original_catch:
        head, _, _ = split_catch_for_ui(original_catch)
    if not sales and original_catch:
        _, sales, _ = split_catch_for_ui(original_catch)
    if not head:
        return str(original_catch or "").strip()
    out = head
    if sales:
        out += sales
        if tail:
            out += "。" + tail.lstrip("。．.")
        elif not sales.endswith("。"):
            out += "。"
    elif tail:
        out += tail if tail.startswith("。") else "。" + tail
    return out


_PKG_WORD = r"(?:袋|缶|瓶|箱|パック|パッケージ)"


def _appeal_body(rest: str) -> str:
    """包材の「〜袋で」は訴求ではない。後ろの文を使う。後ろが無ければ前を使う。"""
    t = re.sub(r"\s+", "", str(rest or ""))
    t = re.sub(r"^[、，。．.]+|[。．.！!？?]+$", "", t)
    if not t:
        return ""
    m = re.search(rf"{_PKG_WORD}[での](.+)", t)
    if m:
        after = re.sub(r"^[、，。．.]+", "", m.group(1))
        if len(after) >= 4:
            return after
    m2 = re.search(rf"^(.*?){_PKG_WORD}で$", t)
    if m2:
        before = re.sub(r"[との]+$", "", m2.group(1))
        if len(before) >= 4:
            return before
    return t


def catch_sales_pair_ja(
    catch: str,
    *,
    catch_max: int = 18,
    sales_max: int = 16,
    used: Optional[set] = None,
) -> tuple[str, str]:
    """画像用: 大＝【】キャッチ、小＝その案だけの訴求。

    包材の「袋で」で切った共通文は出さない。同じ小文が続いた案は小を空にする。
    Returns:
        (catch_image, sales_talk) — sales が取れなければ空文字。
    """
    t = str(catch or "").strip()
    if not t:
        return "", ""
    head = catch_head_ja(t, max_len=catch_max)
    m = re.search(r"【([^】]+)】", t)
    if m:
        rest = t[m.end() :].strip()
    else:
        # 見出し無し: 冒頭をキャッチ、残りを補足
        rest = t[len(head) :] if head and t.startswith(head) else ""
        rest = rest.lstrip("。．.、，,　 ").strip()
    rest = re.sub(r"\s+", "", rest)
    rest = re.sub(r"[。．.！!？?]+$", "", rest)
    # キャッチと同一・ほぼ同一なら補足なし
    if not rest or rest == head:
        return head, ""
    if head and (rest.startswith(head) or head.startswith(rest)):
        return head, ""
    sales = _sales_trim_ja(_appeal_body(rest), sales_max=sales_max)
    if len(sales) < 4:
        return head, ""
    if used is not None:
        if sales in used:
            return head, ""
        used.add(sales)
    return head, sales


def pick_filename(short_name: str, theme_id: str, catch: str) -> str:
    """比較用: かつお_T4_忙しい日の一品に.jpg"""
    sn = sanitize_fs_name(short_name, max_len=24)
    tid = sanitize_fs_name(theme_id, max_len=8)
    ch = catch_head_ja(catch)
    return f"{sn}_{tid}_{ch}.jpg"


def compose_dir_name(short_name: str, when: Optional[str] = None) -> str:
    """かつお_20260829_154935"""
    from datetime import datetime

    sn = sanitize_fs_name(short_name, max_len=24)
    ts = when or datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{sn}_{ts}"


def export02_filename(
    parent_sku: str,
    *,
    compose_dir: str = "",
    collide: bool = False,
    adopt_key: str = "",
) -> str:
    """02 出口。adopt_key あり→テーマ／版別。衝突時だけ __{composeDir}。"""
    sku = str(parent_sku or "").strip()
    if not sku:
        raise ValueError("parent_sku が空")
    key = sanitize_fs_name(str(adopt_key or "").strip(), max_len=24) if adopt_key else ""
    if key:
        base = f"{sku}_単体__{key}.png"
        if collide and compose_dir:
            safe = sanitize_fs_name(compose_dir, max_len=40)
            return f"{sku}_単体__{key}__{safe}.png"
        return base
    if collide and compose_dir:
        safe = sanitize_fs_name(compose_dir, max_len=48)
        return f"{sku}_単体__{safe}.png"
    return f"{sku}_単体.png"
