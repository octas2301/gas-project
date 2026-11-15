# -*- coding: utf-8 -*-
"""単品商品名・JAN パターン判定（SINGLE_NAME 要件 §4–§5 と同型）。"""
from __future__ import annotations

from typing import Any


def _norm_jans(raw: Any) -> list[str]:
    if raw is None or raw == "":
        return []
    if isinstance(raw, (list, tuple)):
        items = [str(x or "").strip() for x in raw]
    else:
        items = [p.strip() for p in str(raw).replace(",", "|").split("|")]
    out: list[str] = []
    for s in items:
        if s and s not in out:
            out.append(s)
    return out


def corpus_assist_allowed(
    *,
    ai_confidence: str = "high",
    existence: str = "ok",
    sale_form: str = "①",
    human_corpus: bool = False,
    after_existence_ng_retry: bool = False,
) -> bool:
    if human_corpus:
        return True
    if str(ai_confidence or "").lower() in ("low", "mid", "低い", "中"):
        return True
    if after_existence_ng_retry or str(existence).lower() == "ng":
        return True
    if str(sale_form or "").startswith("②"):
        return True
    return False


def _result(
    pattern: str,
    *,
    draft: str,
    single_trust: str,
    jan_trust: str,
    cands: list[str],
    need_human: bool,
    corpus_ok: bool,
    write: bool | None = None,
    for_estimate: bool | None = None,
) -> dict:
    if write is None:
        write = bool(draft) and single_trust not in ("要確認", "保留")
    if for_estimate is None:
        for_estimate = write and single_trust not in ("要確認", "保留")
    return {
        "pattern": pattern,
        "single_name": draft if write or single_trust == "要確認" else draft,
        "single_trust": single_trust,
        "jan_trust": jan_trust,
        "jan_candidates": cands,
        "write_single_name": bool(write),
        "for_estimate": bool(for_estimate),
        "need_human": need_human,
        "corpus_assist_ok": corpus_ok,
        "auto_pick_jan": False,
    }


def classify_single_name(inp: dict) -> dict:
    sale = str(inp.get("sale_form") or "①").strip()
    qty = int(inp.get("set_qty") or 1)
    draft = str(inp.get("draft_name") or "").strip()
    existence = str(inp.get("existence") or "ok").strip().lower()
    ai_conf = str(inp.get("ai_confidence") or "high").strip().lower()
    title_match = str(inp.get("title_match") or "high").strip().lower()
    hit_asin = str(inp.get("hit_asin") or "").strip().upper()
    row_asin = str(inp.get("row_asin") or "").strip().upper()
    keepa_jans = _norm_jans(inp.get("keepa_jans"))
    search_jans = _norm_jans(inp.get("search_jans"))
    corpus_ok = corpus_assist_allowed(
        ai_confidence=ai_conf,
        existence=existence,
        sale_form=sale,
        human_corpus=bool(inp.get("human_corpus")),
        after_existence_ng_retry=bool(inp.get("after_existence_ng_retry")),
    )

    if (
        sale.startswith("③")
        or sale.startswith("④")
        or "ミックス" in sale
        or "詰め合わせ" in sale
    ):
        return _result(
            "P-MIX",
            draft="",
            single_trust="要確認",
            jan_trust="触らない",
            cands=keepa_jans or search_jans,
            need_human=True,
            corpus_ok=False,
            write=False,
            for_estimate=False,
        )

    if existence in ("unknown", "不明"):
        return _result(
            "D2",
            draft=draft,
            single_trust="保留",
            jan_trust="触らない",
            cands=keepa_jans,
            need_human=False,
            corpus_ok=corpus_ok,
            write=False,
            for_estimate=False,
        )

    if existence == "ng":
        return _result(
            "D1",
            draft="",
            single_trust="要確認",
            jan_trust="触らない",
            cands=keepa_jans,
            need_human=True,
            corpus_ok=True,
            write=False,
            for_estimate=False,
        )

    # existence OK
    if hit_asin and row_asin and hit_asin != row_asin:
        return _result(
            "F3",
            draft=draft,
            single_trust="要確認",
            jan_trust="low",
            cands=search_jans or keepa_jans,
            need_human=True,
            corpus_ok=corpus_ok,
            write=False,
            for_estimate=False,
        )

    if sale.startswith("①") and title_match == "low":
        return _result(
            "F2",
            draft=draft,
            single_trust="要確認",
            jan_trust="low",
            cands=keepa_jans or search_jans,
            need_human=True,
            corpus_ok=corpus_ok,
            write=False,
            for_estimate=False,
        )

    if ai_conf in ("low", "低い"):
        return _result(
            "E1",
            draft=draft,
            single_trust="要確認",
            jan_trust="low",
            cands=keepa_jans or search_jans,
            need_human=True,
            corpus_ok=True,
            write=False,
            for_estimate=False,
        )

    f1 = sale.startswith("②") and qty > 1 and title_match == "low"
    nk, ns = len(keepa_jans), len(search_jans)

    def finish(pattern: str, single_trust: str, jan_trust: str, cands: list[str], need_human: bool):
        if f1:
            return _result(
                "F1",
                draft=draft,
                single_trust="mid",
                jan_trust=jan_trust,
                cands=cands,
                need_human=need_human,
                corpus_ok=True,
                write=True,
                for_estimate=True,
            )
        return _result(
            pattern,
            draft=draft,
            single_trust=single_trust,
            jan_trust=jan_trust,
            cands=cands,
            need_human=need_human,
            corpus_ok=corpus_ok,
            write=True,
            for_estimate=True,
        )

    if nk == 1:
        kj = keepa_jans[0]
        if ns == 0:
            return finish("A3", "high", "mid", keepa_jans, False)
        if ns == 1 and search_jans[0] == kj:
            return finish("A1", "high", "high", keepa_jans, False)
        if kj in search_jans:
            return finish("A1", "high", "high", keepa_jans, False)
        return finish("A2", "mid", "low", list(dict.fromkeys(keepa_jans + search_jans)), False)

    if nk == 0:
        if ns == 1:
            return finish("B1", "high", "候補", search_jans, True)
        if ns > 1:
            return finish("B2", "mid", "複数", search_jans, True)
        return finish("B3", "high", "欠測", [], False)

    # nk > 1
    if ns == 1:
        return finish("C1", "high", "候補", search_jans, False)
    return finish(
        "C2",
        "mid",
        "複数",
        list(dict.fromkeys(keepa_jans + search_jans)),
        True,
    )
