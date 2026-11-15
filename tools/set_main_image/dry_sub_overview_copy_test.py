# -*- coding: utf-8 -*-
"""仮想: 概要＋訴求メッセージ分離 →（任意）競合参照つき OpenAI 画像。

B-⑤ テーマ別新規本線。確認フォルダへ出すのは、確認ページで採用したときだけ。
`--skip-generate` で既存 T*.jpg のみ（再生成せず export／完走テスト用）。
"""
from __future__ import annotations

import argparse
import json
import logging
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

from dry_sub_overview_openai import _find_truth
from dry_sub_theme_gate import _blob
from photo_realism_rules import photo_realism_block_ja
from theme_compose_rules import (
    accent_text_block as _accent_text_block,
    default_need_package as _default_need_package,
    package_policy_block as _package_policy_block,
    theme_extra_block as _theme_extra_block,
)
from sub_image_intent import TEXT_MODELS, _extract_json_obj

LOG = logging.getLogger("set_main_image.dry_sub_overview_copy_test")
DEFAULT_OVERVIEW = (
    Path(__file__).resolve().parent / "_out" / "dry_sub_theme_gate_4906283045317_v2b.json"
)
DEFAULT_JAN = "4906283045317"
DEFAULT_IMAGE_THEMES = (1, 3, 11)


def parent_sku_from_truth_hint(hint: str) -> str:
    """正本ファイル名から親SKU（ishi-…-oya3_ミックス.png → ishi-…-oya3）。"""
    stem = Path(str(hint or "").replace("\\", "/")).name
    stem = Path(stem).stem
    for suf in ("_単体", "_ミックス", "_mix", "_single"):
        if stem.endswith(suf):
            return stem[: -len(suf)].strip()
    return stem.strip()


def _materials_blob(data: Dict[str, Any]) -> str:
    return _blob(data.get("materials") or {})[:2200]


def _propose_copy_gemini(
    materials_blob: str, hits: Sequence[Dict[str, Any]]
) -> Dict[str, Any]:
    from gemini_image import load_api_key, make_client

    themes = []
    for h in hits:
        themes.append(
            {
                "themeId": int(h["themeId"]),
                "name": h.get("name") or "",
                "overview": h.get("overview") or "",
            }
        )
    prompt = (
        "食品ECサブ画像用の『訴求メッセージ』だけを提案する。概要（シーン）は既にある。\n"
        "各テーマで画像内に載せる短い日本語コピーを出す。\n"
        "ルール:\n"
        "- catch: 主キャッチ 1本。だいたい12〜28字。必須。\n"
        "- subs: 補足 0〜2本。各だいたい8〜20字。\n"
        "- 材料にある語・言い回しだけ。無い数字・賞・効能・価格は作るな。\n"
        "- 「〜を訴求します」など地の文は禁止。コピーそのものを書け。\n"
        "- 概要のシーン説明をそのままコピーにするな。人が一瞬で読む文言にせよ。\n"
        "- テーマ名とずれるキャッチは禁止。\n"
        'JSONのみ: {"items":[{"themeId":1,"catch":"…","subs":["…"],"note":"material|imagine"}]}\n\n'
        "商品材料:\n"
        + materials_blob
        + "\n\nテーマ＋概要:\n"
        + json.dumps(themes, ensure_ascii=False)
    )
    client = make_client(load_api_key())
    t0 = time.time()
    last_err = ""
    for mid in TEXT_MODELS:
        try:
            resp = client.models.generate_content(model=mid, contents=prompt)
            text = getattr(resp, "text", None) or ""
            if not text and getattr(resp, "candidates", None):
                try:
                    text = resp.candidates[0].content.parts[0].text
                except Exception:
                    text = ""
            obj = _extract_json_obj(text)
            if not obj:
                raise RuntimeError("JSON parse failed: " + text[:200])
            items = obj.get("items") if isinstance(obj, dict) else obj
            if not isinstance(items, list):
                raise RuntimeError("items missing")
            by_id: Dict[int, Dict[str, Any]] = {}
            for it in items:
                try:
                    tid = int(it.get("themeId"))
                except (TypeError, ValueError, AttributeError):
                    continue
                catch = str((it or {}).get("catch") or "").strip()
                subs_raw = (it or {}).get("subs") or []
                subs = [str(s).strip() for s in subs_raw if str(s).strip()][:2]
                if catch:
                    by_id[tid] = {
                        "catch": catch[:40],
                        "subs": [s[:28] for s in subs],
                        "note": str((it or {}).get("note") or ""),
                    }
            if not by_id:
                raise RuntimeError("empty copy")
            return {
                "ok": True,
                "model": mid,
                "ms": int((time.time() - t0) * 1000),
                "byId": by_id,
                "error": "",
            }
        except Exception as e:
            last_err = str(e)
            LOG.warning("copy gemini fail %s: %s", mid, e)
    return {
        "ok": False,
        "model": "",
        "ms": int((time.time() - t0) * 1000),
        "byId": {},
        "error": last_err,
    }


def _fallback_copy(overview: str, name: str) -> Dict[str, Any]:
    # 最低限。本番材料があるときは Gemini が本線。
    short = (overview or name or "訴求")[:24]
    return {"catch": short, "subs": [], "note": "fallback"}


def _write_messages_html(out_dir: Path, rows: List[Dict[str, Any]], meta: Dict[str, Any]) -> Path:
    cards = []
    for r in rows:
        subs = r.get("subs") or []
        subs_h = "".join(f"<li>{s}</li>" for s in subs) or "<li>（補足なし）</li>"
        cards.append(
            "<article class='card'>"
            f"<div class='tid'>T{r['themeId']} <span class='name'>{r.get('name','')}</span></div>"
            f"<p class='ov'><b>概要</b><br/>{r.get('overview','')}</p>"
            f"<p class='catch'><b>訴求 catch</b><br/><span class='big'>{r.get('catch','')}</span></p>"
            f"<p class='subs'><b>訴求 subs</b></p><ul>{subs_h}</ul>"
            f"<p class='note'>note={r.get('note','')}／catch {len(r.get('catch') or '')}字</p>"
            "</article>"
        )
    html = f"""<!DOCTYPE html>
<html lang="ja"><head><meta charset="UTF-8"/>
<title>概要＋訴求メッセージ分離（仮想）</title>
<style>
body{{font-family:Segoe UI,Hiragino Sans,sans-serif;margin:0;background:#f7f7f7;color:#222}}
header{{background:#4a148c;color:#fff;padding:16px 20px}}
.warn{{background:#fff8e1;padding:10px 20px;font-size:13px}}
main{{max-width:920px;margin:0 auto;padding:16px}}
.card{{background:#fff;border:1px solid #e0e0e0;border-radius:10px;padding:14px;margin-bottom:12px}}
.big{{font-size:20px;font-weight:700;color:#4a148c}}
.ov{{font-size:13px;line-height:1.55;color:#444}}
.name{{color:#666;font-size:13px;font-weight:400}}
</style></head><body>
<header><h1>概要 ＋ 訴求メッセージ（分離テスト）</h1>
<p>JAN {meta.get('jan')}／Gemini copy ok={meta.get('copyOk')}／model={meta.get('copyModel')}</p></header>
<div class="warn">概要＝シーン指針。訴求＝画像内に載せる短い文言。このあと代表テーマで競合参照つき画像を試します。</div>
<main>{''.join(cards)}</main></body></html>
"""
    path = out_dir / "messages.html"
    path.write_text(html, encoding="utf-8")
    return path


def _download_comp_refs(jan: str, dest_dir: Path, max_n: int = 4) -> List[Path]:
    from b3_comp_catalog import download_url, url_cache_name
    from sub_image_b3_curate import read_accepted_from_b3

    dest_dir.mkdir(parents=True, exist_ok=True)
    accepted = read_accepted_from_b3(jans=[jan])
    items = accepted.get(jan) or []
    paths: List[Path] = []
    for it in items[: max(1, max_n * 2)]:
        url = str(it.get("url") or "")
        if not url.startswith("http"):
            continue
        dest = dest_dir / url_cache_name(url)
        try:
            download_url(url, dest)
            if dest.is_file():
                paths.append(dest)
        except Exception as e:
            LOG.warning("comp dl fail: %s", e)
        if len(paths) >= max_n:
            break
    return paths


HUMAN_PHOTO_STRICT_JA = """
【人物実写強化（人物が写るとき必須）】
- ストック広告／美容広告の顔にするな。ドキュメンタリー寄りの生活スナップにせよ。
- 肌の平滑しすぎ・均一ツヤ・完璧な歯並び・宣伝スマイル・アイドル照明は禁止。
- 正面の美顔クローズアップより、横顔・後ろ姿・手元＋肩・ぼけた顔を優先してよい。
- 非対称な室内光（窓・スタンド1灯）。顔全体を均一に明るくするな。
- 毛髪・服のシワ・手元の自然な影を残す。プラスチック人形感は不合格。
"""

# テーマ別新規（概要＋訴求）専用。競合本線の人物流用とは別系統。
# 参考:
# - https://hana-art.com/realai-prompt/ （国籍・年代・表情／光／レンズ／服装具体）
# - https://www.techno-edge.net/article/2025/08/09/4524.html
#   （LLMで顔パーツを自然文で詳細化／肌・髪の不完全さ／hyperrealistic・8K回避）
THEME_NEW_PERSON_REALISM_JA = """
【人物リアル感（テーマ別新規・共通／必須）】
- 完璧な左右対称・陶器肌・美容レタッチは禁止。自然な不完全さを残せ。
- 肌: skin texture／fine pores／subtle imperfections（かすかなムラ・色差）。産毛・ほくろはごく自然にあってよい。プラスチック平滑・均一ツヤ禁止。
- 髪: stray hair（アホ毛・ほつれ）を少し。均一ウィッグ感禁止。
- 表情: natural smile／candid（作り笑顔・宣伝スマイル禁止）。口角の盛りすぎ・完璧な歯のテカリ禁止。
- 光: soft side light（柔らかい横〜斜め前の自然光）。均一正面フラッシュ・美容広告ライト禁止。接地影・顔の片側に薄い陰影を残せ。
- カメラ: DSLR／50–85mm・f/1.8前後の浅い被写界深度。documentary／snapshot の生活写真。スタジオポートレートの作り込みすぎ禁止。
- 禁止ワード相当の演出: hyperrealistic／photorealistic／8K／masterpiece 風のテカテカCGIにしない（PHOTO_REALISM の HARD BAN と同趣旨）。
- 人物は競合写真から切り貼りしない。詳細どおり新規生成する。
"""

THEME_NEW_PERSON_BY_ID: Dict[int, str] = {
    1: """
【人物（テーマ別新規・詳細指定／必須）】
- 日本人女性、20〜30代。自然で明るい表情（安堵・ほっとした感じ。苦悩顔・暗い表情は禁止）。
- 顔の概略（特定の有名人ではない一般人）: やわらかい卵型〜丸みのある顎、自然な東アジア系の目鼻立ち。濃すぎない眉、小さめの鼻、ナチュラルな唇。厚化粧禁止。
- 場面: 夜〜夕方の自宅で、小腹が空いたあとにヘルシーなおやつを楽しんでいる生活シーン。
- 構図: 上半身〜腰上。背景（テーブル・カップ等）も少し入れる。顔だけアップ禁止。構図は固定しない。
- 器: 手のひらサイズ前後の小皿。大皿・大きな石皿は禁止（おつまみ量に合った小ささ）。
- 服装: ベージュやグレー系のやわらかいニット／部屋着風。布のシワは自然に。
- 悩みの「痛み」は画像内の訴求文字で伝え、人物の表情は明るい解決後〜楽しんでいる側にする。
"""
    + THEME_NEW_PERSON_REALISM_JA,
    2: """
【人物（テーマ別新規・詳細指定／必須）】
- 日本人女性、20〜30代。満足げ・寛ぎの自然な表情（笑顔気味でよい。作り笑顔の盛りすぎは禁止）。
- 顔の概略（特定の有名人ではない一般人）: やわらかい卵型〜丸みのある顎、自然な東アジア系の目鼻立ち。濃すぎない眉、小さめの鼻、ナチュラルな唇。厚化粧禁止。
- 場面: 一日の終わりの自宅で、お酒と一緒に晩酌を楽しむ生活シーン。
- 構図: 上半身〜腰上。横顔寄りでも正面でも可。背景（グラス・小皿等）も少し入れる。顔だけアップ禁止。構図は固定しない。
- 器: 手のひらサイズ前後の小皿／酒の肴皿。大皿・大きな石皿・顔幅超えのトレーは禁止。
- 服装: 落ち着いたベージュ／グレー系の部屋着やカーディガン。布のシワは自然に。
"""
    + THEME_NEW_PERSON_REALISM_JA,
}


def _theme_new_person_block(theme_id: int) -> str:
    return THEME_NEW_PERSON_BY_ID.get(int(theme_id), "")


def _sub_picture_extras(
    refs: Dict[str, Any],
    mix_plan: Optional[Dict[str, Any]],
    *,
    need_package: bool,
    out_dir: Path,
    theme_id: int,
) -> Dict[str, Any]:
    """P3 §4.5.10: WRAP／袋の指示材料。切戻し時は空寄り。"""
    from b4_sub_picture_rules import bag_prompt_and_verdict, picture_gate_enabled
    from b5_food_look_read import read_wrap_looks_for_paths, wrap_look_block_ja

    use_wrap = bool(refs.get("useWrap") and refs.get("wrap"))
    wrap_path = Path(refs["wrap"]) if use_wrap else None
    wrap_items: List[Dict[str, Any]] = []
    if wrap_path and wrap_path.is_file():
        wrap_items.append({"path": str(wrap_path), "label": "WRAP"})
    if mix_plan:
        for i, it in enumerate(mix_plan.get("identities") or []):
            wp = it.get("wrap") if isinstance(it, dict) else None
            if wp and Path(str(wp)).is_file():
                wrap_items.append(
                    {"path": str(wp), "label": f"WRAP_{i+1}", "pathObj": Path(str(wp))}
                )
                use_wrap = True
    wrap_look = ""
    if picture_gate_enabled() and wrap_items:
        looks = read_wrap_looks_for_paths(wrap_items)
        wrap_look = wrap_look_block_ja(looks)
        try:
            import re as _re

            from b5_slot_lang import (
                piece_size_hint_from_jan,
                wrap_piece_relative_size_block,
            )

            jan_hint = ""
            m = _re.search(r"theme_new_run_(\d+)", str(out_dir))
            if m:
                jan_hint = piece_size_hint_from_jan(m.group(1), max_n=3)
            rel = wrap_piece_relative_size_block(
                has_wrap=True, piece_size_hint=jan_hint
            )
            if rel:
                wrap_look = (wrap_look + "\n" + rel).strip() + "\n"
        except Exception:
            pass
    bags: List[str] = []
    if need_package and mix_plan:
        for it in mix_plan.get("identities") or []:
            b = it.get("bag") if isinstance(it, dict) else None
            if b and Path(str(b)).is_file():
                bags.append(str(b))
    bag_info: Dict[str, Any] = {"rows": [], "block": "", "verdictPath": ""}
    if picture_gate_enabled() and need_package and len(bags) >= 2:
        bag_info = bag_prompt_and_verdict(
            bags, out_dir=Path(out_dir) / "比較", theme_id=f"T{int(theme_id):02d}"
        )
    mix_parts: List[str] = []
    if mix_plan:
        for it in mix_plan.get("identities") or []:
            lab = str((it or {}).get("label") or (it or {}).get("name") or "").strip()
            if lab:
                mix_parts.append(lab)
    return {
        "has_wrap": bool(use_wrap and wrap_items),
        "wrap_items": wrap_items,
        "wrap_look_block": wrap_look,
        "bag_aspect_block": str(bag_info.get("block") or ""),
        "bag_verdict_path": str(bag_info.get("verdictPath") or ""),
        "mix_parts": mix_parts,
        "mix_part_bags": bool(need_package and bags),
    }


def _build_image_prompt(
    *,
    product_name: str,
    theme_id: int,
    theme_name: str,
    overview: str,
    catch: str,
    subs: Sequence[str],
    n_comp: int,
    photo_strict_human: bool = False,
    base_color: str = "beige",
    need_package: bool = False,
    text_color: str = "1",
    human_feedback: str = "",
    has_food_unit_canon: bool = False,
    has_food_unit_secondary: bool = False,
    has_food_unit_sizzle: bool = False,
    has_food_unit_wrap: bool = False,
    has_piece_set: bool = False,
    has_surface_shape_s1: bool = False,
    food_mode: str = "identity",
    food_look_block: str = "",
    wrap_look_block: str = "",
    bag_aspect_block: str = "",
    mix_block: str = "",
    mix_parts: Optional[Sequence[str]] = None,
    mix_part_bags: bool = False,
) -> str:
    from tonmana_palette import tonmana_block_ja
    from b4_sub_picture_rules import (
        build_sub_food_front,
        picture_gate_enabled,
        prepare_overview_for_prompt,
    )

    has_wrap = bool(has_food_unit_wrap) and picture_gate_enabled()
    has_piece_body = bool(has_piece_set)
    has_s1 = bool(has_surface_shape_s1) and not has_piece_body
    # S1も「一片ロック場面」扱い（IDENTITY＝配置のみ）へ寄せる
    has_piece = has_piece_body or has_s1
    overview_use = prepare_overview_for_prompt(str(overview or ""), has_wrap=has_wrap)
    sub_lines = "\n".join(f"- {s}" for s in subs) if subs else "- （補足なし）"
    person_new = _theme_new_person_block(theme_id)
    # テーマ別新規で詳細人物があるときは、旧「隠す寄り」厳格ブロックは使わない
    human_block = "" if person_new else (HUMAN_PHOTO_STRICT_JA if photo_strict_human else "")
    if person_new:
        layout_lines = (
            "- 全面の広告ストック写真にするな。競合の説明カード／フレーム／バッジ等のパーツ感は参考にしてよい。\n"
            "- 競合の構図・人物・文字配置の完全複製は禁止（流用可・全く同じはNG）。\n"
            "- 人物は詳細指定どおり新規生成。競合人物の切り貼り・顔のトレースは禁止。\n"
            "- 皿・グラス・テーブル・背景も人物と同水準の実写（PHOTO_REALISM の物・背景条項）。CGI小物禁止。"
        )
        comp_line = (
            f"- IMAGE_COMP_*: 競合サブ参照 {n_comp} 枚。"
            "外側パネル・レイアウト・文字カードの形は参考流用してよい。"
            "構図・人物・文字の完全複製は禁止。"
            "食材の色・厚み・1片サイズの正本にしない。"
        )
    else:
        layout_lines = (
            "- 全面の広告ストック写真にするな。競合の説明カード／フレーム／バッジ等のパーツ感を残せ。\n"
            "- パーツ流用は可。構図・人物・文字配置の完全複製は禁止（全く同じはNG）。\n"
            "- AI独自の再構成は面積のおよそ半分まで。残りは競合パーツ寄り。\n"
            "- 皿・グラス・テーブル・背景も実写。CGI小物・のっぺり背景禁止。"
        )
        comp_line = (
            f"- IMAGE_COMP_*: 競合サブ参照 {n_comp} 枚。"
            "外側パネル・レイアウト・文字カードの形を優先して流用せよ。"
            "ただし構図・人物・文字の完全複製は禁止。"
            "食材の色・厚み・1片サイズの正本にしない。"
        )
    pkg_block = _package_policy_block(bool(need_package), theme_id)
    extra = _theme_extra_block(
        theme_id,
        has_food_unit_canon=has_food_unit_canon,
        has_food_unit_secondary=has_food_unit_secondary,
    )
    accent = _accent_text_block(text_color)
    feedback_block = ""
    if (human_feedback or "").strip():
        feedback_block = (
            "\n【人間フィードバック（再生成・最優先で反映）】\n"
            f"{human_feedback.strip()}\n"
            "上記以外のPACKAGE_LOCK・食材ユニット正本・ベース色は維持すること。\n"
        )
    food_in = ""
    if has_s1:
        food_in += (
            "- IMAGE_FOOD_UNIT_SURFACE_*（あるとき）: 面アップ。"
            "筋・色・血合い色・表面感だけ。外形はSURFACEから取るな。\n"
            "- IMAGE_FOOD_UNIT_SHAPE_*（あるとき）: 同ファミリー別SKU一片。"
            "細長さ・反り・丸まり・ねじれの形だけ。色・筋・縁質感はSURFACEが正。"
            "自SKU一片本体は今回添付していない。\n"
        )
    elif has_piece_body:
        food_in += (
            "- IMAGE_FOOD_UNIT_SURFACE_*（あるとき）: 面アップ／面クロップ。"
            "筋・色・血合い色・表面感だけ。外形はSURFACEから取るな。\n"
            "- IMAGE_FOOD_UNIT_PIECE_*: 一片セット。開封後個品の色・縁・厚み・筋・血合いの正。"
            "言語化も同番号のみ。白背景の単一片標本禁止。白い小皿に適量盛り。\n"
            "- IMAGE_FOOD_UNIT_SHAPE_*（あるとき）: 同ファミリー別SKU一片。"
            "細長さ・反り・丸まり・ねじれの形だけ。色・筋・縁質感はSURFACE／PIECE側が正。\n"
        )
    if has_food_unit_canon:
        if has_piece:
            food_in += (
                "- IMAGE_FOOD_UNIT_CANON: IDENTITY盛り写真。器・量感・重なり・交差・"
                "ばらけ幅・山の高さの正。個品の色・縁・厚みの正にしない。"
                f"（mode={food_mode}）\n"
            )
        else:
            food_in += (
                "- IMAGE_FOOD_UNIT_CANON: IDENTITY主。色・厚み・1片サイズの正。"
                f"（本テーマ mode={food_mode}）\n"
            )
    if has_food_unit_sizzle:
        food_in += (
            "- IMAGE_FOOD_UNIT_SIZZLE: SIZZLE主。炙り・とろけ等の食欲状態の正。"
            "IDENTITYと違ってよい。状態の創作禁止（見本に無い炙りを足すな）。\n"
        )
    if has_food_unit_secondary:
        food_in += (
            "- IMAGE_FOOD_UNIT_SECONDARY_*: IDENTITY副。盛り・形のみ。"
            "色は IDENTITY 主優先（副の照明差に引っ張られるな）。\n"
        )
    if has_wrap:
        food_in += (
            "- IMAGE_FOOD_UNIT_WRAP: 個包装専用。大袋・背景には使うな。"
            "皿の横か後ろに1〜2個。カメラ最近傍は開封後の個品。\n"
        )
    body = f"""あなたは食品ECのサブ画像を1枚作るデザイナーです（楽天向け正方形）。

【商品】{product_name}
【主テーマ】T{theme_id:02d} {theme_name}

【テーマ概要＝構図・演出の指針】
{overview_use}
※概要の文章そのものを画像内に大きく書くな。絵の指示として使え。

【訴求メッセージ＝画像内に必ず短く載せる】
主キャッチ: {catch}
補足:
{sub_lines}
※上記は説明カード／吹き出し／帯などに読めない大きさで載せよ。キャッチが主役の文字。
{accent}
【入力】
- IMAGE_PACKAGE_TRUTH: 正本（載せる場合のみ画面に出す。載せない場合は色参照のみ）。
{food_in}{comp_line}
{feedback_block}
{pkg_block}
{extra}
【トンマナ／ベース色（全サブで揃える制約）】
{tonmana_block_ja(base_color)}
※背景・カード・余白のみ。商品パッケージ色は変えるな。MAINシリーズと同じベース色系統を崩すな。
{person_new}
【本物感・レイアウト】
{layout_lines}
{photo_realism_block_ja()}
{human_block}
【禁止】偽物のランキング数字、価格、効能の発明。宣伝スマイルの偽お客様顔。正本パッケージの色替え。競合の完全複製。不要なパッケージの常時表示。常温珍味本体からの湯気。ギフトで袋の直接手渡し。人物の手・腕・肩が袋の輪郭から生える／袋を貫通する／袋の天面から出る構図。手前の袋の上に人物部位を重ねない（人物と袋は奥行きで分ける）。マップと無関係な強調色。テーマごとに別物の食材に見える盛り。

【出力】正方形の商用サブ画像。訴求メッセージが読め、テーマが一瞬で分かること。
"""
    if not has_food_unit_canon and not has_wrap and not has_piece:
        extra_m = str(mix_block or "").strip()
        return (extra_m + "\n\n" + body) if extra_m else body
    front = build_sub_food_front(
        has_identity=bool(has_food_unit_canon),
        has_sizzle=bool(has_food_unit_sizzle),
        has_secondary=bool(has_food_unit_secondary),
        has_wrap=bool(has_wrap),
        has_piece_set=bool(has_piece),
        food_mode=str(food_mode or "identity"),
        mix_parts=list(mix_parts or []),
        mix_part_bags=bool(mix_part_bags),
        food_look_block=str(food_look_block or ""),
        wrap_look_block=str(wrap_look_block or ""),
        bag_aspect_block=str(bag_aspect_block or ""),
        mix_block=str(mix_block or ""),
    )
    return (front + "\n\n" + body) if front else body


def _resolve_sub_piece_set(
    jan: str, *, max_n: int = 3
) -> Tuple[List[Path], str, List[Path], List[Path]]:
    """一片セット写真＋ロック／言語化／盛り言語（SUB本線）。無ければ空。

    第3戻り値: SHAPEドナー。第4戻り値: SURFACE面アップ（B5_SURFACE_SHAPE_SPLIT=1 時）。
    S1（仮想C）: SURFACE/SHAPEがあるとき自SKU一片本体は返さず（添付しない）。
    切戻し: B5_S1_OMIT_OWN_PIECE=0／B5_SURFACE_SHAPE_SPLIT=0。
    """
    from food_unit_canon import (
        identity_plating_lock_prompt_block,
        load_identity_plating_lang_block,
        load_piece_lang_prompt_block,
        piece_set_lock_prompt_block,
        resolve_piece_set_paths,
        theme_new_dir_for_jan,
    )

    pieces_all = resolve_piece_set_paths(str(jan or ""), max_n=int(max_n or 3))
    if not pieces_all:
        return [], "", [], []
    n = len(pieces_all)
    shape_paths: List[Path] = []
    surface_paths: List[Path] = []
    omit_own = False
    try:
        from b5_piece_deepen import (
            s1_omit_own_piece_enabled,
            surface_shape_split_enabled,
        )
        from b5_piece_truth_rules import (
            resolve_shape_donor_paths,
            resolve_surface_ref_paths,
            shape_donor_attach_prompt_block,
            surface_attach_prompt_block,
        )

        if surface_shape_split_enabled():
            tn = theme_new_dir_for_jan(str(jan or ""))
            smeta = resolve_surface_ref_paths(tn, max_n=int(max_n or 3))
            if smeta.get("ok"):
                surface_paths = [
                    Path(p) for p in (smeta.get("paths") or []) if Path(p).is_file()
                ]
            try:
                from b5_generic_plating import (
                    generic_plating_core_enabled,
                    prefer_face_crop_surfaces,
                )

                if generic_plating_core_enabled() and surface_paths:
                    surface_paths = prefer_face_crop_surfaces(
                        tn, surface_paths, max_n=int(max_n or 3)
                    )
            except Exception as ge:
                LOG.warning("sub generic face crop skip: %s", ge)
            meta = resolve_shape_donor_paths(tn, max_n=int(max_n or 3))
            if meta.get("ok"):
                shape_paths = [
                    Path(p) for p in (meta.get("paths") or []) if Path(p).is_file()
                ]
            omit_own = bool(s1_omit_own_piece_enabled()) and bool(
                surface_paths or shape_paths
            )
    except Exception as e:
        LOG.warning("sub surface/shape resolve fail: %s", e)

    pieces = [] if omit_own else list(pieces_all)
    parts: List[str] = []
    if omit_own:
        from b5_piece_deepen import surface_shape_split_enabled  # noqa: F401
        from b5_slot_lang import (
            build_shape_lang_prompt_from_donors,
            build_surface_lang_prompt_block,
            s1_synthesis_prompt_block,
        )
        from b5_piece_truth_rules import resolve_shape_donor_paths

        parts.extend(
            [
                s1_synthesis_prompt_block(),
                identity_plating_lock_prompt_block(),
                load_identity_plating_lang_block(str(jan)),
                "【SUB・S1／一片本体なし】白背景の単一片標本禁止。"
                "個品の色・筋は SURFACE＋SURFACE言語化。形は SHAPE＋SHAPE言語化。"
                "IDENTITY盛りは配置・量感のみ。自SKU一片本体は添付していない。",
            ]
        )
        if surface_paths:
            from b5_piece_truth_rules import surface_attach_prompt_block

            sb0 = surface_attach_prompt_block(max_n=len(surface_paths))
            if sb0:
                parts.append(sb0)
            s_lang = build_surface_lang_prompt_block(
                theme_new_dir_for_jan(str(jan)), surface_paths, ensure=True
            )
            if s_lang:
                parts.append(s_lang)
        if shape_paths:
            from b5_piece_truth_rules import shape_donor_attach_prompt_block

            sb = shape_donor_attach_prompt_block(max_n=len(shape_paths))
            if sb:
                parts.append(sb)
            try:
                dmeta = resolve_shape_donor_paths(
                    theme_new_dir_for_jan(str(jan)), max_n=len(shape_paths)
                )
                donor_dirs = [
                    Path(x) for x in (dmeta.get("donorDirs") or []) if Path(x).is_dir()
                ]
            except Exception:
                donor_dirs = []
            sh_lang = build_shape_lang_prompt_from_donors(
                donor_dirs, max_n=len(shape_paths)
            )
            if sh_lang:
                parts.append(sh_lang)
        LOG.info(
            "sub S1 omit own piece jan=%s surfaceN=%s shapeN=%s",
            jan,
            len(surface_paths),
            len(shape_paths),
        )
    else:
        parts.extend(
            [
                piece_set_lock_prompt_block(max_n=n),
                load_piece_lang_prompt_block(str(jan), max_n=n),
                identity_plating_lock_prompt_block(),
                load_identity_plating_lang_block(str(jan)),
                "【SUB・一片あり】白背景の単一片標本禁止。"
                "IDENTITY盛り写真の重なり・密度・ばらけ・山の高さを小皿へ移植せよ。"
                "個品の色・筋は SURFACE（あれば）＋一片セット。盛り型だけ別物にするな。",
            ]
        )
        if surface_paths:
            from b5_piece_truth_rules import surface_attach_prompt_block

            sb0 = surface_attach_prompt_block(max_n=len(surface_paths))
            if sb0:
                parts.append(sb0)
        if shape_paths:
            from b5_piece_truth_rules import shape_donor_attach_prompt_block

            sb = shape_donor_attach_prompt_block(max_n=len(shape_paths))
            if sb:
                parts.append(sb)
    block = "\n".join(p for p in parts if p).strip()
    try:
        from b5_generic_plating import apply_core_to_prompt_text, generic_plating_core_enabled

        if generic_plating_core_enabled() and block:
            block = apply_core_to_prompt_text(block, caller="sub")
            LOG.info("sub generic plating core applied jan=%s", jan)
    except Exception as ge:
        LOG.warning("sub generic plating core skip: %s", ge)
    return pieces, block, shape_paths, surface_paths


def _food_look_for_canon(canon: Optional[Path]) -> Tuple[str, List[Dict[str, Any]]]:
    """主見本があるときだけ、見え方の文章と照合対象を返す。"""
    if not canon or not Path(canon).is_file():
        return "", []
    from b5_food_look_read import look_block_ja, read_looks_for_paths

    items = [{"path": str(canon), "label": "主見本"}]
    looks = read_looks_for_paths(items)
    block = look_block_ja(looks)
    LOG.info("sub food look chars=%s", len(block))
    return block, items


def _mix_plan(jan: str, name_base: str, parent_sku: str = "") -> Optional[Dict[str, Any]]:
    """セット構成1が空でない親だけ。1種類しか取れなければ単品に戻す。"""
    nb = str(name_base or "").strip()
    if not nb and not str(parent_sku or "").strip():
        return None
    try:
        from b5_virt_abc import _jan_digits, _name_base_key, _sheet_cell
        from master_sets import _find_header_row
        from sheets_master import fetch_master_rows
        from food_unit_canon import combo_parts_for_mix_name, resolve_main_food_refs
        from dry_b5_propose_virt import (
            OFFICIAL_URL_CACHE_PATH,
            dishes_from_materials_,
            dishes_from_official_page_,
        )
    except Exception as e:
        LOG.warning("sub mix import fail: %s", e)
        return None
    try:
        rows, _info = fetch_master_rows()
        _hi, idx = _find_header_row(rows)
    except Exception as e:
        LOG.warning("sub mix sheet fail: %s", e)
        return None
    jan_i = idx.get("JANコード") or idx.get("JAN")
    parent_i = idx.get("親SKU")
    child_i = idx.get("子SKU")
    name_i = idx.get("商品名ベース")
    combo_i = idx.get("セット構成1_親SKU")
    cat_i = next((idx[k] for k in idx if "カテゴリー" in k and "食品" in k), None)
    want_jan = _jan_digits(jan)
    want_nb = _name_base_key(nb)
    want_sku = str(parent_sku or "").strip()
    hit = None
    for r in rows[_hi + 1 :]:
        if child_i is not None and _sheet_cell(r, child_i):
            continue
        if want_sku and _sheet_cell(r, parent_i) != want_sku:
            continue
        if want_nb and _name_base_key(_sheet_cell(r, name_i)) != want_nb:
            continue
        if not want_sku and not want_nb:
            if _jan_digits(_sheet_cell(r, jan_i)) != want_jan:
                continue
        hit = r
        break
    if hit is None or combo_i is None or not _sheet_cell(hit, combo_i):
        LOG.info("sub mix skip reason=single")
        return None
    sku = _sheet_cell(hit, parent_i)
    # 正: マスタセット構成1〜5（商品名中黒では分けない）
    parts = combo_parts_for_mix_name(
        name_base=_sheet_cell(hit, name_i) or nb,
        parent_sku=sku,
    )
    if len(parts) < 2:
        LOG.info("sub mix fallback single parts=%s parent=%s", len(parts), sku)
        return None
    is_food = bool(cat_i is not None and "食品" in _sheet_cell(hit, cat_i))
    identities: List[Dict[str, Any]] = []
    missing: List[str] = []
    for part in parts:
        refs = resolve_main_food_refs(str(part.get("jan") or ""), "T4")
        ident = refs.get("identity")
        if not ident or not Path(ident).is_file():
            missing.append(str(part.get("nameBase") or part.get("jan") or ""))
            continue
        identities.append(
            {
                "nameBase": str(part.get("nameBase") or ""),
                "jan": str(part.get("jan") or ""),
                "path": str(ident),
                "bag": str(part.get("packageTruth") or ""),
            }
        )
    dishes: List[str] = []
    if is_food:
        from b5_virt_abc import _compact_facts_from_sheets

        blobs = []
        mix_mat = _compact_facts_from_sheets(want_jan, nb, sku).get("materials") or {}
        blobs.append(mix_mat)
        for part in parts:
            blobs.append(
                _compact_facts_from_sheets(
                    str(part.get("jan") or ""),
                    str(part.get("nameBase") or ""),
                    str(part.get("parentSku") or ""),
                ).get("materials")
                or {}
            )
        seen = set()
        for mat in blobs:
            for row in dishes_from_materials_(mat):
                name = str(row.get("dishName") or "").strip()
                if name and name not in seen:
                    seen.add(name)
                    dishes.append(name)
        url = ""
        if OFFICIAL_URL_CACHE_PATH.is_file():
            import json as _json

            cache = _json.loads(OFFICIAL_URL_CACHE_PATH.read_text(encoding="utf-8"))
            url = str(((cache.get("byKey") or {}).get(f"jan:{want_jan}") or {}).get("url") or "")
        if url:
            for row in dishes_from_official_page_(url, mix_mat if isinstance(mix_mat, dict) else {}):
                name = str(row.get("dishName") or "").strip()
                if name and name not in seen:
                    seen.add(name)
                    dishes.append(name)
    names = [str(p.get("nameBase") or "") for p in parts]
    LOG.info("sub mix parts=%s food=%s dishes=%s missing=%s", len(parts), is_food, len(dishes), missing)
    return {
        "names": names,
        "identities": identities,
        "missing": missing,
        "dishes": dishes,
        "isFood": is_food,
    }


def _piece_mode(theme_id: int) -> str:
    """個品の見本片を出すか。compare=素材の比較、specimen=見本が主役、dish=料理の上だけ、none=出さない。"""
    from sub_image_mix_manifest import _theme_meta

    name = str(_theme_meta(int(theme_id)).get("themeName") or "")
    if "素材" in name:
        return "compare"
    if "セット内容" in name or "シズル" in name or "味の" in name:
        return "specimen"
    if "食べ方" in name or "アレンジ" in name or "日常" in name:
        return "dish"
    return "none"


def _mix_prompt_block(plan: Optional[Dict[str, Any]], *, theme_id: int, need_package: bool) -> str:
    if not plan:
        return ""
    names = [n for n in (plan.get("names") or []) if n]
    if len(names) < 2:
        return ""
    n = len(names)
    joined = "、".join(names)
    from sub_image_mix_manifest import _theme_meta

    mode = _piece_mode(theme_id)
    theme_name = str(_theme_meta(int(theme_id)).get("themeName") or "")
    lines = ["【ミックス・必須】", f"構成は{n}種類: {joined}。"]
    if mode == "compare":
        lines += [
            "このテーマだけ、左右に種類ごとの特徴を書いてよい。1枚の文字に、各種の特徴を全部書く。",
            "キャッチと補足を合わせて、各種の特徴が読める。片方の名前だけは不合格。",
            "個品の見本を主役にする。種類は全部出す。見本の色・厚み・1片のまま。",
            "見本片の横に、名前と内容量だけを短く書く。",
        ]
    elif mode == "specimen":
        lines += [
            "個品を主役にする。種類は全部出す。見本の色・厚み・1片のまま。",
            "見本片の横に、名前と内容量だけを短く書く。",
            "左右に特徴を書き分ける長い説明は作らない。",
            "種類が揃う、という締めの定型文は足さない。",
            "画像に載せる文は、下のキャッチと補足だけ。",
        ]
    elif mode == "dish":
        lines += [
            "別皿の見本片と、名前と内容量のラベルは置かない。",
            "個品は料理の上だけに出す。種類は全部。見本の色・厚み・1片のまま。",
            "種類が揃う、という締めの定型文は足さない。",
            "画像に載せる文は、下のキャッチと補足だけ。",
        ]
    else:
        lines += [
            "個品の見本片、名前、内容量は足さない。",
            "主役は、このテーマの役割だけ。",
            "種類が揃う、という締めの定型文は足さない。",
            "画像に載せる文は、下のキャッチと補足だけ。",
        ]
    if need_package:
        lines.append("袋を出す。各袋の縦横比と、袋どうしの大きさの対比は正本のまま。片方だけ小さくしない。")
    else:
        lines.append("このテーマは袋を出さない。")
    if plan.get("isFood") and mode == "dish":
        dishes = list(plan.get("dishes") or [])
        if not dishes:
            lines.append("食べ方の候補が無い。料理名は空。商品そのものを、このテーマの場に置く。")
        elif "アレンジ" in theme_name:
            picks = "、".join(dishes[:3])
            lines.append(f"料理の候補から、別々の料理を最大3つ。使う候補は「{picks}」。新しい料理名は作らない。")
            lines.append("各料理の上の個品は、ミックスの全部。選んだ食べ方と別の食べ物は足さない。")
        else:
            pick = dishes[(int(theme_id) - 1) % len(dishes)]
            lines.append(f"料理の候補から、このテーマは「{pick}」だけ。新しい料理名は作らない。")
            lines.append("皿の上の個品は、ミックスの全部。選んだ食べ方と別の食べ物は足さない。")
    elif plan.get("isFood") and mode == "none":
        lines.append("料理の見本皿も、個品の見本片も置かない。")
    return "\n".join(lines)


def _generate_with_food_check(
    *,
    out_path: Path,
    prompt_kwargs: Dict[str, Any],
    image_paths: Sequence[Path],
    model: str,
    quality: str,
    look_items: Sequence[Dict[str, Any]],
    wrap_items: Optional[Sequence[Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    """1枚生成し、個片／WRAP合否で最大1回作り直す（P3 §4.5.10）。"""
    from b4_sub_picture_rules import picture_gate_enabled, wrap_gate_failed
    from b5_food_look_read import check_drawn_foods, check_drawn_wraps, read_wrap_looks_for_paths
    from openai_image import generate_with_references

    tries = out_path.parent / "tries"
    tries.mkdir(parents=True, exist_ok=True)
    feedback = str(prompt_kwargs.get("human_feedback") or "")
    wrap_list = [w for w in (wrap_items or []) if Path(str((w or {}).get("path") or "")).is_file()]
    if _piece_mode(int(prompt_kwargs.get("theme_id") or 0)) == "none":
        drop = {Path(str(it.get("path") or "")).resolve() for it in look_items if it.get("path")}
        image_paths = [
            p
            for p in image_paths
            if Path(p).resolve() not in drop
            and "FOOD_UNIT" not in Path(p).name.upper()
            and not Path(p).name.lower().startswith("piece_set_")
        ]
        look_items = []
        prompt_kwargs = dict(prompt_kwargs)
        prompt_kwargs["has_piece_set"] = False
    chosen: Optional[Path] = None
    used = ""
    piece: Dict[str, Any] = {"ok": True, "skipped": True}
    wrap_chk: Dict[str, Any] = {"ok": True, "skipped": True}
    last_err = ""
    attempt_done = 0
    wrap_looks = read_wrap_looks_for_paths(wrap_list) if (
        picture_gate_enabled() and wrap_list
    ) else []
    for attempt in (1, 2):
        kwargs = dict(prompt_kwargs)
        kwargs["human_feedback"] = feedback
        prompt = _build_image_prompt(**kwargs)
        try:
            data_b, used = generate_with_references(
                prompt=prompt,
                image_paths=list(image_paths),
                model=model,
                size="1024x1024",
                quality=quality,
            )
        except Exception as e:
            last_err = str(e)
            LOG.exception("sub gen fail %s try=%s", out_path.name, attempt)
            break
        attempt_done = attempt
        try_path = tries / f"{out_path.stem}_try{attempt}.jpg"
        try_path.write_bytes(data_b)
        (tries / f"{out_path.stem}_try{attempt}.txt").write_text(prompt, encoding="utf-8")
        if picture_gate_enabled() and wrap_list:
            try:
                from b4_wrap_scene_refine import (
                    refine_scene_wrap_to_file,
                    wrap_scene_refine_enabled,
                )

                if wrap_scene_refine_enabled():
                    w0 = Path(str(wrap_list[0].get("path") or ""))
                    if w0.is_file():
                        rpath = tries / f"{out_path.stem}_try{attempt}_wrap_refine.jpg"
                        refine_meta = refine_scene_wrap_to_file(
                            try_path,
                            w0,
                            rpath,
                            looks=wrap_looks,
                            work_dir=tries / f"{out_path.stem}_try{attempt}_wrap_refine_work",
                        )
                        if (
                            refine_meta.get("ok")
                            and not refine_meta.get("skipped")
                            and rpath.is_file()
                        ):
                            try_path.write_bytes(rpath.read_bytes())
                        (tries / f"{out_path.stem}_try{attempt}_wrap_refine.json").write_text(
                            json.dumps(refine_meta, ensure_ascii=False, indent=2),
                            encoding="utf-8",
                        )
            except Exception as e:
                LOG.warning("sub wrap refine skip: %s", e)
        chosen = try_path
        if look_items:
            piece = check_drawn_foods(try_path, look_items)
        else:
            piece = {"ok": True, "skipped": True}
        LOG.info(
            "sub food piece %s try=%s ok=%s skipped=%s",
            out_path.name,
            attempt,
            piece.get("ok"),
            piece.get("skipped"),
        )
        notes: List[str] = []
        note_piece = str(piece.get("retryNote") or "").strip()
        piece_fail = (
            bool(look_items)
            and (not piece.get("skipped"))
            and (not piece.get("ok"))
            and bool(note_piece)
        )
        if piece_fail:
            notes.append(note_piece)
        if picture_gate_enabled() and wrap_list:
            wrap_chk = check_drawn_wraps(try_path, wrap_list, wrap_looks)
            LOG.info(
                "sub wrap %s try=%s ok=%s front=%s skipped=%s",
                out_path.name,
                attempt,
                wrap_chk.get("ok"),
                wrap_chk.get("frontOk"),
                wrap_chk.get("skipped"),
            )
            if wrap_gate_failed(wrap_chk):
                wn = str(wrap_chk.get("retryNote") or "").strip()
                if wn:
                    notes.append(wn)
        if (not notes) or attempt == 2:
            break
        feedback = (feedback + "\n" + "\n".join(notes)).strip()
    prompts = out_path.parent / "prompts"
    prompts.mkdir(exist_ok=True)
    last_txt = tries / f"{out_path.stem}_try{attempt_done}.txt"
    if last_txt.is_file():
        (prompts / f"{out_path.stem}.txt").write_text(
            last_txt.read_text(encoding="utf-8"), encoding="utf-8"
        )
    if chosen is None:
        return {
            "ok": False,
            "model": used,
            "attempts": attempt_done,
            "foodPieceOk": False,
            "foodPieceSkipped": True,
            "wrapOk": None,
            "error": last_err or "generate failed",
        }
    out_path.write_bytes(chosen.read_bytes())
    return {
        "ok": True,
        "model": used,
        "attempts": attempt_done,
        "foodPieceOk": bool(piece.get("ok")),
        "foodPieceSkipped": bool(piece.get("skipped")),
        "wrapOk": None if wrap_chk.get("skipped") else bool(wrap_chk.get("ok")),
        "wrapFrontOk": wrap_chk.get("frontOk"),
        "error": "",
    }


def _theme_id_from_stem(stem: str) -> Optional[int]:
    import re

    m = re.match(r"^T(\d+)(?:__r\d+)?$", str(stem or "").strip(), re.I)
    if not m:
        return None
    return int(m.group(1))


def _next_theme_out_path(out_dir: Path, theme_id: int) -> Path:
    """旧 Txx.jpg は残し、次世代 __rN を返す。"""
    import re

    from b5_set_base_map import next_regen_filename

    tid = int(theme_id)
    base = out_dir / f"T{tid:02d}.jpg"
    cands = []
    if base.is_file():
        cands.append(base)
    for p in out_dir.glob(f"T{tid:02d}__r*.jpg"):
        cands.append(p)
    if not cands:
        return base

    def _gen_key(p: Path) -> int:
        m = re.search(r"__r(\d+)$", p.stem, re.I)
        return int(m.group(1)) if m else 1

    latest = max(cands, key=_gen_key)
    name = next_regen_filename(latest.name)
    dest = out_dir / name
    guard = 0
    while dest.is_file():
        name = next_regen_filename(name)
        dest = out_dir / name
        guard += 1
        if guard > 40:
            raise RuntimeError(f"regen name exhausted: T{tid:02d}")
    return dest


def run_theme_new_regen(
    *,
    theme_new_dir: Path,
    queue_path: Path,
    quality: str = "medium",
    photo_strict_human: bool = False,
    base_color: str = "beige",
) -> Dict[str, Any]:
    """
    混在レビュー段階B: チェック済みテーマだけ再生成（旧残し・__rN）。
    messages.json / 01_refs / FOOD_UNIT_CANON を再利用。
    """
    from food_unit_canon import prepare_sub_food_canon, resolve_food_refs_for_theme
    from openai_image import resolve_openai_image_model

    root = Path(theme_new_dir)
    msg_path = root / "messages.json"
    if not msg_path.is_file():
        raise SystemExit(f"messages.json がありません: {msg_path}")
    msg = json.loads(msg_path.read_text(encoding="utf-8"))
    summary_path = root / "summary.json"
    summary = (
        json.loads(summary_path.read_text(encoding="utf-8"))
        if summary_path.is_file()
        else {}
    )
    queue = json.loads(Path(queue_path).read_text(encoding="utf-8-sig"))
    items = [x for x in (queue.get("items") or []) if x.get("checked")]
    if not items:
        LOG.info("theme_new regen: 対象0")
        return {"ok": True, "n": 0, "rows": []}

    product_name = str(msg.get("productName") or summary.get("productName") or "")
    text_color = str(msg.get("textColor") or summary.get("textColor") or "1")
    jan = str(msg.get("jan") or summary.get("jan") or "")
    by_tid = {int(r["themeId"]): r for r in (msg.get("rows") or []) if r.get("themeId")}
    truth_s = str(summary.get("truth") or "")
    truth = Path(truth_s) if truth_s else None
    if not truth or not truth.is_file():
        truth = _find_truth(jan, "")
    comps_dir = root / "01_refs"
    comps = [
        p
        for p in sorted(comps_dir.iterdir())
        if p.is_file()
        and p.suffix.lower() in (".jpg", ".jpeg", ".png", ".webp")
        and not p.name.upper().startswith("FOOD_UNIT_CANON")
        and not p.name.upper().startswith("FOOD_UNIT_SECONDARY")
        and not p.name.upper().startswith("FOOD_UNIT_SIZZLE")
        and not p.name.upper().startswith("PACKAGE_TRUTH")
    ][:8]
    food_meta = prepare_sub_food_canon(jan, root, truth=Path(truth))
    if food_meta.get("blocked"):
        raise SystemExit(f"JAN={jan} {food_meta.get('reasonJa')}")
    canon_path = Path(str(food_meta.get("path") or ""))
    look_block, look_items = _food_look_for_canon(canon_path if canon_path.is_file() else None)
    piece_paths, piece_block, shape_paths, surface_paths = _resolve_sub_piece_set(
        jan, max_n=3
    )
    mix_plan = _mix_plan(jan, product_name)
    if mix_plan and mix_plan.get("identities"):
        from b5_food_look_read import look_block_ja, read_looks_for_paths

        mix_items = [
            {"path": it["path"], "label": it["nameBase"] or "構成"}
            for it in mix_plan["identities"]
        ]
        look_items = mix_items
        look_block = look_block_ja(read_looks_for_paths(mix_items))
    model = resolve_openai_image_model(None)
    (root / "prompts").mkdir(exist_ok=True)
    rows: List[Dict[str, Any]] = []
    for it in items:
        stem = str(it.get("stem") or "").strip()
        tid = _theme_id_from_stem(stem)
        if tid is None:
            try:
                tid = int(it.get("themeId") or 0) or None
            except (TypeError, ValueError):
                tid = None
        if not tid:
            LOG.warning("theme_new regen skip bad stem=%s", stem)
            continue
        r = by_tid.get(tid)
        if not r:
            LOG.warning("theme_new regen no message row T%02d", tid)
            continue
        feedback = str(it.get("comment") or "").strip()
        need_pkg = _default_need_package(tid, r.get("needPackage"))
        refs = resolve_food_refs_for_theme(root, tid, need_package=need_pkg)
        use_ident = bool(refs.get("identity"))
        use_sizzle = bool(refs.get("sizzle"))
        use_sec = bool(refs.get("secondary"))
        mode_ok = _piece_mode(tid) != "none"
        use_piece = bool(piece_paths) and mode_ok
        use_s1 = mode_ok and bool(surface_paths or shape_paths) and not use_piece
        use_food = use_piece or use_s1
        pic = _sub_picture_extras(
            refs, mix_plan, need_package=need_pkg, out_dir=root, theme_id=tid
        )
        image_paths: List[Path] = [Path(truth)]
        if use_ident and Path(refs["identity"]).resolve() != Path(truth).resolve():
            image_paths.append(Path(refs["identity"]))
        if use_sizzle:
            image_paths.append(Path(refs["sizzle"]))
        if use_sec:
            image_paths.extend(list(refs["secondary"]))
        if use_food:
            if surface_paths:
                image_paths.extend(surface_paths)
            if piece_paths:
                image_paths.extend(piece_paths)
            if shape_paths:
                image_paths.extend(shape_paths)
        for wi in pic.get("wrap_items") or []:
            wp = Path(str(wi.get("path") or ""))
            if wp.is_file():
                image_paths.append(wp)
        image_paths.extend(comps)
        if mix_plan and mix_plan.get("missing"):
            rows.append(
                {
                    "themeId": tid,
                    "stem": stem,
                    "out": "",
                    "feedback": feedback,
                    "ok": False,
                    "ms": 0,
                    "error": "構成の見本が無い: " + "、".join(mix_plan["missing"]),
                    "foodMode": refs.get("mode"),
                }
            )
            continue
        if mix_plan:
            for it in mix_plan.get("identities") or []:
                image_paths.append(Path(it["path"]))
                if need_pkg and it.get("bag") and Path(it["bag"]).is_file():
                    image_paths.append(Path(it["bag"]))
        look_for = look_block if (use_ident or mix_plan or use_food) else ""
        if use_food and piece_block:
            look_for = (look_for + "\n\n" + piece_block).strip() if look_for else piece_block
        # 一片あり時の照合正は一片写真。S1はSURFACE（IDENTITY盛りは配置のみ）
        look_chk: List[Dict[str, Any]] = []
        if use_piece and piece_paths:
            look_chk = [
                {"path": str(p), "label": f"一片{i + 1}"}
                for i, p in enumerate(piece_paths)
            ]
        elif use_s1 and surface_paths:
            look_chk = [
                {"path": str(p), "label": f"SURFACE{i + 1}"}
                for i, p in enumerate(surface_paths)
            ]
        elif use_ident or mix_plan:
            look_chk = list(look_items)
        out_path = _next_theme_out_path(root, tid)
        rec: Dict[str, Any] = {
            "themeId": tid,
            "stem": stem,
            "out": str(out_path),
            "feedback": feedback,
            "ok": False,
            "ms": 0,
            "error": "",
            "foodMode": refs.get("mode"),
            "pieceSetN": len(piece_paths) if use_piece else 0,
            "shapeDonorN": len(shape_paths) if use_food else 0,
            "surfaceN": len(surface_paths) if use_food else 0,
        }
        g0 = time.time()
        made = _generate_with_food_check(
            out_path=out_path,
            prompt_kwargs={
                "product_name": product_name,
                "theme_id": tid,
                "theme_name": str(r.get("name") or ""),
                "overview": str(r.get("overview") or ""),
                "catch": str(r.get("catch") or ""),
                "subs": list(r.get("subs") or []),
                "n_comp": len(comps),
                "photo_strict_human": bool(photo_strict_human),
                "base_color": str(base_color or "beige"),
                "need_package": need_pkg,
                "text_color": text_color,
                "human_feedback": feedback,
                "has_food_unit_canon": use_ident,
                "has_food_unit_secondary": use_sec,
                "has_food_unit_sizzle": use_sizzle,
                "has_food_unit_wrap": bool(pic.get("has_wrap")),
                "has_piece_set": use_piece,
                "has_surface_shape_s1": use_s1,
                "food_mode": str(refs.get("mode") or "identity"),
                "food_look_block": look_for,
                "wrap_look_block": str(pic.get("wrap_look_block") or ""),
                "bag_aspect_block": str(pic.get("bag_aspect_block") or ""),
                "mix_block": _mix_prompt_block(mix_plan, theme_id=tid, need_package=need_pkg),
                "mix_parts": list(pic.get("mix_parts") or []),
                "mix_part_bags": bool(pic.get("mix_part_bags")),
            },
            image_paths=image_paths,
            model=model,
            quality=str(quality),
            look_items=look_chk,
            wrap_items=list(pic.get("wrap_items") or []),
        )
        rec["ok"] = bool(made.get("ok"))
        rec["model"] = made.get("model") or ""
        rec["ms"] = int((time.time() - g0) * 1000)
        rec["error"] = made.get("error") or ""
        rec["foodPieceOk"] = made.get("foodPieceOk")
        rec["attempts"] = made.get("attempts")
        if rec["ok"]:
            LOG.info("theme_new regen OK %s -> %s ms=%s", stem, out_path.name, rec["ms"])
        rows.append(rec)

    regen_summary = {
        "at": time.strftime("%Y-%m-%dT%H:%M:%S"),
        "n": len(rows),
        "okN": sum(1 for x in rows if x.get("ok")),
        "foodCanon": food_meta,
        "rows": rows,
    }
    (root / "regen_summary.json").write_text(
        json.dumps(regen_summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    ok = len(rows) == 0 or regen_summary["okN"] == len(rows)
    return {"ok": ok, **regen_summary}


def _write_gallery(
    out_dir: Path, rows: List[Dict[str, Any]], meta: Dict[str, Any]
) -> Path:
    cards = []
    for r in rows:
        img = (
            f"<img src='{r['relPath']}'/>"
            if r.get("ok") and r.get("relPath")
            else f"<pre class='err'>{r.get('error','')}</pre>"
        )
        subs = r.get("subs") or []
        cards.append(
            "<article class='card'>"
            f"<h3>T{r['themeId']} {r.get('name','')}</h3>"
            f"<p class='catch'>{r.get('catch','')}</p>"
            f"<p class='ov'>{r.get('overview','')}</p>"
            f"<p class='subs'>{' ／ '.join(subs)}</p>"
            f"{img}"
            f"<p class='meta'>comps={r.get('nComp')} ms={r.get('ms')}</p>"
            "</article>"
        )
    html = f"""<!DOCTYPE html>
<html lang="ja"><head><meta charset="UTF-8"/>
<title>overview+copy OpenAI test</title>
<style>
body{{font-family:Segoe UI,Hiragino Sans,sans-serif;margin:0;background:#111;color:#eee}}
header{{padding:16px 20px;background:#222}}
main{{display:grid;grid-template-columns:repeat(auto-fill,minmax(340px,1fr));gap:12px;padding:16px}}
.card{{background:#1b1b1b;border:1px solid #333;border-radius:10px;padding:12px}}
.card img{{width:100%;border-radius:6px;margin-top:8px}}
.catch{{font-size:18px;font-weight:700;color:#ce93d8}}
.ov,.subs{{font-size:12px;color:#bbb;line-height:1.45}}
.err{{color:#ff8a80;font-size:11px;white-space:pre-wrap}}
</style></head><body>
<header><h1>概要＋訴求＋競合参照 → OpenAI</h1>
<p>JAN {meta.get('jan')} ok={meta.get('okN')}/{meta.get('n')} comps={meta.get('compN')}</p></header>
<main>{''.join(cards)}</main></body></html>
"""
    path = out_dir / "gallery.html"
    path.write_text(html, encoding="utf-8")
    return path


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    ap = argparse.ArgumentParser(description="概要＋訴求分離の仮想テスト")
    ap.add_argument("--overview-json", type=Path, default=DEFAULT_OVERVIEW)
    ap.add_argument("--jan", default=DEFAULT_JAN)
    ap.add_argument(
        "--name-base",
        default="",
        help="商品名ベース。指定時は試験用JSONを読まず、その親行の出品材料から概要を作る",
    )
    ap.add_argument("--parent-sku", default="", help="親SKU。商品名ベースと併用可")
    ap.add_argument("--truth-hint", default="", help="正本ファイル名ヒント（空=JAN表／商品名）。まぐろ既定は禁止")
    ap.add_argument(
        "--image-themes",
        default="1,3,11",
        help="画像化する themeId（空=画像なし）",
    )
    ap.add_argument("--quality", default="medium", choices=("low", "medium", "high"))
    ap.add_argument("--max-comps", type=int, default=3)
    ap.add_argument("--out-dir", type=Path, default=None)
    ap.add_argument(
        "--messages-json",
        type=Path,
        default=None,
        help="既存の messages.json を再利用（訴求再生成をスキップ）",
    )
    ap.add_argument(
        "--photo-strict-human",
        action="store_true",
        help="人物実写強化ブロックを追加（ストック美顔禁止）",
    )
    ap.add_argument(
        "--reuse-comps-dir",
        type=Path,
        default=None,
        help="既存 01_refs を再利用（競合再DLをスキップ）",
    )
    ap.add_argument("--base-color", default="beige", choices=("beige", "warm_white", "soft_gray"))
    ap.add_argument(
        "--text-color",
        default="auto",
        help="強調色=MAIN textColor（1..5）。auto=02.楽天ベース set_base_map から解決（既定）",
    )
    ap.add_argument(
        "--regen-from",
        type=Path,
        default=None,
        help="テーマ別新規 out ディレクトリ。--regen-queue と併用で段階B再生成",
    )
    ap.add_argument(
        "--regen-queue",
        type=Path,
        default=None,
        help="再生成キューJSON（checked+comment+stem）",
    )
    ap.add_argument(
        "--force-food-canon",
        action="store_true",
        help="FOOD_UNIT_CANON を再選定する（既存があっても上書き）",
    )
    ap.add_argument(
        "--skip-generate",
        action="store_true",
        help="OpenAI生成をせず既存 T*.jpg を前提に完了（修正・最終テスト用）",
    )
    ap.add_argument(
        "--auto-export",
        action=argparse.BooleanOptionalAction,
        default=False,
        help="計画だけ残す。確認フォルダへは出さない（採用時に出す）",
    )
    ap.add_argument(
        "--export-out-dir",
        type=Path,
        default=None,
        help="--auto-export 先（既定: 02.楽天アップロード画像保存場所）",
    )
    args = ap.parse_args()

    from b5_set_base_map import (
        normalize_text_color,
        resolve_text_color_for_jan,
        save_shared_text_color,
    )

    # 文字色 auto の前に親SKUを確定（単品・ミックスとも。無いと ishi-{JAN}-oya に落ちる）
    if not str(args.parent_sku or "").strip():
        inferred = parent_sku_from_truth_hint(str(args.truth_hint or ""))
        if inferred:
            args.parent_sku = inferred
            LOG.info("parent-sku from truth-hint -> %s", inferred)

    _tc_raw = str(args.text_color or "auto").strip().lower()
    _ps = str(getattr(args, "parent_sku", None) or "").strip()
    if _tc_raw in ("", "auto", "map", "from-map"):
        args.text_color = resolve_text_color_for_jan(
            str(args.jan or ""),
            parent_sku=_ps or None,
            default="1",
        )
        LOG.info("text-color auto jan=%s parent=%s -> %s", args.jan, _ps or "(fallback)", args.text_color)
    else:
        args.text_color = normalize_text_color(args.text_color, default="1")
        # 明示指定＝サブ先でも共有へ書く
        if not _ps and args.jan:
            _ps = f"ishi-{str(args.jan).strip()}-oya"
        if _ps:
            try:
                save_shared_text_color(_ps, args.text_color)
            except Exception:
                pass

    if args.regen_from:
        q = args.regen_queue or (Path(args.regen_from) / "regen_queue.json")
        if not Path(q).is_file():
            raise SystemExit(f"regen queue がありません: {q}")
        result = run_theme_new_regen(
            theme_new_dir=Path(args.regen_from),
            queue_path=Path(q),
            quality=str(args.quality),
            photo_strict_human=bool(args.photo_strict_human),
            base_color=str(args.base_color or "beige"),
        )
        print(json.dumps(result, ensure_ascii=False))
        return 0 if result.get("ok") else 2

    name_base = str(args.name_base or "").strip()
    parent_sku = str(args.parent_sku or "").strip()
    reuse_msg = bool(args.messages_json and Path(args.messages_json).is_file())
    if reuse_msg:
        data = {"hitOverviews": [{"themeId": 1, "name": "", "overview": ""}], "materials": {}}
    elif name_base or parent_sku:
        from dry_sub_theme_gate import run_gate

        data = run_gate(
            jan=str(args.jan),
            name_base=name_base,
            parent_sku=parent_sku,
            gemini=True,
            prompt_ver="v2",
        )
        gate_ok = bool((data.get("gemini") or {}).get("ok"))
        if not data.get("hitOverviews") or not gate_ok:
            raise SystemExit(
                "商品名ベースの場面概要が作れません: "
                + str(data.get("source") or "")
                + " "
                + str((data.get("gemini") or {}).get("error") or data.get("sheetError") or "")
            )
        LOG.info(
            "name-base materials name=%s source=%s themes=%s",
            str((data.get("materials") or {}).get("name") or ""),
            data.get("source"),
            data.get("proposeIds"),
        )
    else:
        data = json.loads(Path(args.overview_json).read_text(encoding="utf-8"))
    hits = list(data.get("hitOverviews") or [])
    if not hits:
        raise SystemExit("hitOverviews 空")

    ts = time.strftime("%Y%m%d_%H%M%S")
    out_dir = Path(args.out_dir) if args.out_dir else (
        Path(__file__).resolve().parent
        / "_out"
        / f"overview_copy_{args.jan}_{ts}{'_humstrict' if args.photo_strict_human else ''}"
    )
    out_dir.mkdir(parents=True, exist_ok=True)

    if args.skip_generate:
        existing = sorted(out_dir.glob("T*.jpg"))
        if not existing:
            raise SystemExit(f"--skip-generate: 既存画像がありません: {out_dir}/T*.jpg")
        export_info: Dict[str, Any] = {"skipped": True}
        if args.auto_export:
            from b5_skip_generate_finish import export_theme_new_to_rakuten
            from work_paths import DEFAULT_RAKUTEN_UPLOAD_DIR

            dest = Path(args.export_out_dir) if args.export_out_dir else Path(DEFAULT_RAKUTEN_UPLOAD_DIR)
            export_info = export_theme_new_to_rakuten(
                jan=str(args.jan), theme_new_dir=out_dir, out_dir=dest
            )
        summary_skip = {
            "ok": True,
            "mode": "skip_generate",
            "jan": args.jan,
            "outDir": str(out_dir),
            "imageN": len(existing),
            "export": export_info,
        }
        print(json.dumps(summary_skip, ensure_ascii=False, indent=2))
        if args.auto_export:
            print("確認フォルダへは出しません。採用したあとだけ出します。")
        return 0

    product_name = str((data.get("materials") or {}).get("name") or args.jan)
    gem: Dict[str, Any] = {"ok": True, "model": "reused", "ms": 0, "error": "", "byId": {}}
    if args.messages_json and Path(args.messages_json).is_file():
        msg_meta = json.loads(Path(args.messages_json).read_text(encoding="utf-8"))
        rows = list(msg_meta.get("rows") or [])
        product_name = str(msg_meta.get("productName") or product_name)
        LOG.info("reuse messages %s n=%s", args.messages_json, len(rows))
        (out_dir / "messages.json").write_text(
            json.dumps(msg_meta, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        msg_html = _write_messages_html(out_dir, rows, msg_meta)
    else:
        mat_blob = _materials_blob(data)
        gem = _propose_copy_gemini(mat_blob, hits)
        rows = []
        for h in hits:
            tid = int(h["themeId"])
            cp = (gem.get("byId") or {}).get(tid) or _fallback_copy(
                str(h.get("overview") or ""), str(h.get("name") or "")
            )
            rows.append(
                {
                    "themeId": tid,
                    "name": h.get("name") or "",
                    "overview": h.get("overview") or "",
                    "catch": cp.get("catch") or "",
                    "subs": cp.get("subs") or [],
                    "note": cp.get("note") or "",
                    "needPackage": _default_need_package(tid, cp.get("needPackage")),
                }
            )
        msg_meta = {
            "jan": args.jan,
            "productName": product_name,
            "textColor": str(args.text_color or "1"),
            "copyOk": gem.get("ok"),
            "copyModel": gem.get("model"),
            "copyMs": gem.get("ms"),
            "copyError": gem.get("error") or "",
            "rows": rows,
        }
        (out_dir / "messages.json").write_text(
            json.dumps(msg_meta, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        msg_html = _write_messages_html(out_dir, rows, msg_meta)
    LOG.info("messages -> %s", msg_html)

    image_ids = [
        int(x)
        for x in str(args.image_themes or "").replace(" ", "").split(",")
        if x.isdigit()
    ]
    if not image_ids:
        print("MESSAGES_ONLY", msg_html)
        return 0 if gem.get("ok") else 2

    truth = _find_truth(str(args.jan), str(args.truth_hint))
    if args.reuse_comps_dir and Path(args.reuse_comps_dir).is_dir():
        comps = [
            p
            for p in sorted(Path(args.reuse_comps_dir).iterdir())
            if p.is_file() and p.suffix.lower() in (".jpg", ".jpeg", ".png", ".webp")
        ][: int(args.max_comps)]
        LOG.info("reuse comps dir=%s n=%s", args.reuse_comps_dir, len(comps))
    else:
        comps = _download_comp_refs(str(args.jan), out_dir / "01_refs", max_n=int(args.max_comps))
    from food_unit_canon import prepare_sub_food_canon, resolve_food_refs_for_theme
    from openai_image import resolve_openai_image_model

    if args.force_food_canon:
        LOG.info("force-food-canon は人間確定を機械選択で上書きしない")
    food_meta = prepare_sub_food_canon(str(args.jan), out_dir, truth=Path(truth))
    if food_meta.get("blocked"):
        raise SystemExit(f"JAN={args.jan} {food_meta.get('reasonJa')}")
    food_canon = Path(str(food_meta.get("path") or ""))
    look_block, look_items = _food_look_for_canon(food_canon if food_canon.is_file() else None)
    piece_paths, piece_block, shape_paths, surface_paths = _resolve_sub_piece_set(
        str(args.jan), max_n=3
    )
    mix_plan = _mix_plan(str(args.jan), name_base, parent_sku)
    if mix_plan and mix_plan.get("identities"):
        from b5_food_look_read import look_block_ja, read_looks_for_paths

        mix_items = [
            {"path": it["path"], "label": it["nameBase"] or "構成"}
            for it in mix_plan["identities"]
        ]
        look_items = mix_items
        look_block = look_block_ja(read_looks_for_paths(mix_items))
    LOG.info(
        "truth=%s comps=%s photoStrictHuman=%s textColor=%s foodCanon=%s meta=%s pieceSetN=%s shapeDonorN=%s surfaceN=%s",
        truth,
        len(comps),
        args.photo_strict_human,
        args.text_color,
        food_canon,
        food_meta.get("confirmedBy") or food_meta.get("source"),
        len(piece_paths),
        len(shape_paths),
        len(surface_paths),
    )

    model = resolve_openai_image_model(None)
    by_tid = {int(r["themeId"]): r for r in rows}
    img_rows: List[Dict[str, Any]] = []
    for tid in image_ids:
        r = by_tid.get(tid)
        if not r:
            img_rows.append(
                {
                    "themeId": tid,
                    "ok": False,
                    "error": "no message row",
                    "name": "",
                    "overview": "",
                    "catch": "",
                    "subs": [],
                    "nComp": len(comps),
                    "ms": 0,
                }
            )
            continue
        need_pkg = _default_need_package(tid, r.get("needPackage"))
        refs = resolve_food_refs_for_theme(out_dir, tid, need_package=need_pkg)
        use_ident = bool(refs.get("identity"))
        use_sizzle = bool(refs.get("sizzle"))
        use_sec = bool(refs.get("secondary"))
        mode_ok = _piece_mode(tid) != "none"
        use_piece = bool(piece_paths) and mode_ok
        use_s1 = mode_ok and bool(surface_paths or shape_paths) and not use_piece
        use_food = use_piece or use_s1
        pic = _sub_picture_extras(
            refs, mix_plan, need_package=need_pkg, out_dir=out_dir, theme_id=tid
        )
        stem = f"T{tid:02d}"
        image_paths = [truth]
        if use_ident and Path(refs["identity"]).resolve() != Path(truth).resolve():
            image_paths.append(Path(refs["identity"]))
        if use_sizzle:
            image_paths.append(Path(refs["sizzle"]))
        if use_sec:
            image_paths.extend(list(refs["secondary"]))
        if use_food:
            if surface_paths:
                image_paths.extend(surface_paths)
            if piece_paths:
                image_paths.extend(piece_paths)
            if shape_paths:
                image_paths.extend(shape_paths)
        for wi in pic.get("wrap_items") or []:
            wp = Path(str(wi.get("path") or ""))
            if wp.is_file():
                image_paths.append(wp)
        image_paths.extend(comps)
        if mix_plan and mix_plan.get("missing"):
            img_rows.append(
                {
                    "themeId": tid,
                    "name": r.get("name") or "",
                    "overview": r.get("overview") or "",
                    "catch": r.get("catch") or "",
                    "subs": r.get("subs") or [],
                    "ok": False,
                    "relPath": "",
                    "error": "構成の見本が無い: " + "、".join(mix_plan["missing"]),
                    "nComp": len(comps),
                    "ms": 0,
                    "foodMode": refs.get("mode"),
                }
            )
            continue
        if mix_plan:
            for it in mix_plan.get("identities") or []:
                image_paths.append(Path(it["path"]))
                if need_pkg and it.get("bag") and Path(it["bag"]).is_file():
                    image_paths.append(Path(it["bag"]))
        look_for = look_block if (use_ident or mix_plan or use_food) else ""
        if use_food and piece_block:
            look_for = (look_for + "\n\n" + piece_block).strip() if look_for else piece_block
        look_chk: List[Dict[str, Any]] = []
        if use_piece and piece_paths:
            look_chk = [
                {"path": str(p), "label": f"一片{i + 1}"}
                for i, p in enumerate(piece_paths)
            ]
        elif use_s1 and surface_paths:
            look_chk = [
                {"path": str(p), "label": f"SURFACE{i + 1}"}
                for i, p in enumerate(surface_paths)
            ]
        elif use_ident or mix_plan:
            look_chk = list(look_items)
        rec: Dict[str, Any] = {
            "themeId": tid,
            "name": r.get("name") or "",
            "overview": r.get("overview") or "",
            "catch": r.get("catch") or "",
            "subs": r.get("subs") or [],
            "ok": False,
            "relPath": "",
            "error": "",
            "nComp": len(comps),
            "ms": 0,
            "foodMode": refs.get("mode"),
            "pieceSetN": len(piece_paths) if use_piece else 0,
            "shapeDonorN": len(shape_paths) if use_food else 0,
            "surfaceN": len(surface_paths) if use_food else 0,
        }
        g0 = time.time()
        out_path = out_dir / f"{stem}.jpg"
        made = _generate_with_food_check(
            out_path=out_path,
            prompt_kwargs={
                "product_name": product_name,
                "theme_id": tid,
                "theme_name": str(r.get("name") or ""),
                "overview": str(r.get("overview") or ""),
                "catch": str(r.get("catch") or ""),
                "subs": list(r.get("subs") or []),
                "n_comp": len(comps),
                "photo_strict_human": bool(args.photo_strict_human),
                "base_color": str(args.base_color or "beige"),
                "need_package": need_pkg,
                "text_color": str(args.text_color or "1"),
                "has_food_unit_canon": use_ident,
                "has_food_unit_secondary": use_sec,
                "has_food_unit_sizzle": use_sizzle,
                "has_food_unit_wrap": bool(pic.get("has_wrap")),
                "has_piece_set": use_piece,
                "has_surface_shape_s1": use_s1,
                "food_mode": str(refs.get("mode") or "identity"),
                "food_look_block": look_for,
                "wrap_look_block": str(pic.get("wrap_look_block") or ""),
                "bag_aspect_block": str(pic.get("bag_aspect_block") or ""),
                "mix_block": _mix_prompt_block(mix_plan, theme_id=tid, need_package=need_pkg),
                "mix_parts": list(pic.get("mix_parts") or []),
                "mix_part_bags": bool(pic.get("mix_part_bags")),
            },
            image_paths=image_paths,
            model=model,
            quality=str(args.quality),
            look_items=look_chk,
            wrap_items=list(pic.get("wrap_items") or []),
        )
        rec["ok"] = bool(made.get("ok"))
        rec["relPath"] = out_path.name if rec["ok"] else ""
        rec["model"] = made.get("model") or ""
        rec["ms"] = int((time.time() - g0) * 1000)
        rec["error"] = made.get("error") or ""
        rec["foodPieceOk"] = made.get("foodPieceOk")
        rec["attempts"] = made.get("attempts")
        if rec["ok"]:
            LOG.info("OK %s ms=%s attempts=%s foodPiece=%s", stem, rec["ms"], rec["attempts"], rec["foodPieceOk"])
        img_rows.append(rec)

    summary = {
        "jan": args.jan,
        "productName": product_name,
        "textColor": str(args.text_color or "1"),
        "truth": str(truth),
        "foodCanon": food_meta,
        "compN": len(comps),
        "n": len(img_rows),
        "okN": sum(1 for x in img_rows if x.get("ok")),
        "messagesHtml": str(msg_html),
        "copyOk": gem.get("ok"),
        "rows": img_rows,
    }
    (out_dir / "summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    gallery = _write_gallery(out_dir, img_rows, summary)
    export_info: Dict[str, Any] = {"skipped": True}
    if args.auto_export and summary["okN"] > 0:
        from b5_skip_generate_finish import export_theme_new_to_rakuten
        from work_paths import DEFAULT_RAKUTEN_UPLOAD_DIR

        dest = Path(args.export_out_dir) if args.export_out_dir else Path(DEFAULT_RAKUTEN_UPLOAD_DIR)
        export_info = export_theme_new_to_rakuten(
            jan=str(args.jan), theme_new_dir=out_dir, out_dir=dest
        )
        summary["export"] = export_info
        (out_dir / "summary.json").write_text(
            json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    print(
        json.dumps(
            {
                "copyOk": gem.get("ok"),
                "okN": summary["okN"],
                "compN": len(comps),
                "messagesHtml": str(msg_html),
                "gallery": str(gallery),
                "out": str(out_dir),
                "export": export_info,
            },
            ensure_ascii=False,
        )
    )
    if args.auto_export:
        print("確認フォルダへは出しません。採用したあとだけ出します。")
    return 0 if (gem.get("ok") and summary["okN"] == len(img_rows)) else 2


if __name__ == "__main__":
    raise SystemExit(main())
