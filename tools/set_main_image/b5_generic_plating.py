# -*- coding: utf-8 -*-
"""一片見本・盛り（MAIN/SUB）共通レイヤ（C1〜C7 / T4/T8核）。

正本: docs/org/PLATING_AND_PIECE_GENERIC_LOGIC.md
汎用性の定義: Agent Store GENERICITY_DEFINITION.md（型 vs 一件焼き）

切戻し: B5_GENERIC_PLATING_CORE=0
個別適用プロファイル: B5_PLATING_PROFILE=generic|food_flake（既定 food_flake＝削り系）
"""
from __future__ import annotations

import base64
import json
import os
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

# --- C2: 危険文（縁の規則アーティファクトを増やす言い回し）---
_DANGER_EDGE_PATTERNS = [
    r"おつまみのさきいかのように[^\n]*",
    r"細い長い繊維が縁から飛び出し[^\n]*",
    r"ちりぢり・糸状ほつれが不規則[^\n]*",
    r"縁と切れ端は糸状の魚肉繊維がほつれ・毛羽立ち・裂けていること[^\n]*",
    r"／人間修正:[^\n]*",
    r"細かなギザつきや凹凸[^\n]*",
    r"細かなギザギザやほつれ[^\n]*",
    r"不揃いなギザつき[^\n]*",
    r"全体的に細かな凹凸やギザつき[^\n]*",
]


def generic_plating_core_enabled() -> bool:
    raw = str(os.environ.get("B5_GENERIC_PLATING_CORE") or "1").strip().lower()
    return raw not in ("0", "false", "off", "no")


def plating_profile() -> str:
    raw = str(os.environ.get("B5_PLATING_PROFILE") or "food_flake").strip().lower()
    return raw or "food_flake"


def sanitize_danger_prompt(text: str) -> str:
    """C2: NGを増やす指示文を落とす（SKU非焼付）。"""
    raw = str(text or "")
    if not raw:
        return ""
    for pat in _DANGER_EDGE_PATTERNS:
        raw = re.sub(pat, "", raw)
    raw = re.sub(
        r"(-\s*縁:\s*)([^\n]+)",
        r"\1滑らか〜弱い不規則の丸み。等間隔ギザ・鋸歯・規則的とげ列は禁止。",
        raw,
    )
    return re.sub(r"\n{3,}", "\n\n", raw).strip()


def plating_core_prompt_block(*, caller: str = "plating") -> str:
    """盛り（MAIN/SUB）および一片に同送する短い固定核。

    caller: piece | plating | main | sub（ログ用。文面は共通＋プロファイル）
    """
    _ = caller
    prof = plating_profile()
    # 汎用核（どの商品でも同じ型）
    core = (
        "【共通・盛り／一片核】\n"
        "参照の役割を混ぜるな（面の正・形の正・盛り配置・袋は別）。\n"
        "規則的なアーティファクト（等間隔の歯・くし歯・作った感じの縁）は不合格。\n"
        "見たい属性は面の中に残せ。縁の飾りで代用するな。\n"
    )
    # 個別適用（プロファイル）
    if prof == "food_flake":
        core += (
            "【個別適用・削り／薄片】\n"
            "身の平らな面に長い木目筋を残せ。つるつる無地は不合格。\n"
            "縁は滑らか〜弱い不規則の丸み／薄いささくれ。鋸歯・ピンキング禁止。\n"
            "しなやかなねじれ・丸まりを一部に。硬い板の平行並べ禁止。\n"
        )
    else:
        core += (
            "【個別適用・generic】\n"
            "その商品で必須の質感属性を面に残せ。規則縁で代用するな。\n"
        )
    return core.strip()


def is_hand_macro_surface_name(name: str) -> bool:
    """スマホマクロ／Webカメラ手撮り面（本線外）。Web由来の close 等は残す。"""
    u = str(name or "").lower()
    if not u:
        return False
    return any(
        tok in u
        for tok in (
            "phone",
            "ref_surface_cam",
            "_cam_",
            "webcam",
            "web_cam",
        )
    )


def is_test_surface_artifact_name(name: str) -> bool:
    """試験過程の産物（色転送・実験スタンプ等）。本番同送に渡さない。"""
    u = str(name or "").lower()
    if not u:
        return False
    return any(
        tok in u
        for tok in (
            "reinhard",
            "ab_only",
            "keepl",
            "_lab_",
            "color_transfer",
            "surface_seed_ab",
            "virt_probe",
        )
    )


def is_excluded_surface_attach_name(name: str) -> bool:
    """マクロ・試験産物など、SURFACE同送から外す名前。"""
    return is_hand_macro_surface_name(name) or is_test_surface_artifact_name(name)


def filter_macro_excluded_surfaces(paths: Sequence[Path]) -> List[Path]:
    """手撮りマクロ面・試験産物を同送候補から外す。"""
    out: List[Path] = []
    for p in paths:
        pp = Path(p)
        if not pp.is_file():
            continue
        if is_excluded_surface_attach_name(pp.name):
            continue
        out.append(pp)
    return out


def ensure_surface_face_crop(
    surface_src: Path,
    dest: Path,
    *,
    box_frac: tuple[float, float, float, float] = (0.22, 0.22, 0.78, 0.78),
) -> Optional[Path]:
    """C4: 縁をほぼ切った面中央クロップ。"""
    try:
        from PIL import Image
    except ImportError:
        return None
    src = Path(surface_src)
    if not src.is_file():
        return None
    im = Image.open(src).convert("RGB")
    w, h = im.size
    l, t, r, b = box_frac
    box = (int(w * l), int(h * t), int(w * r), int(h * b))
    dest = Path(dest)
    dest.parent.mkdir(parents=True, exist_ok=True)
    im.crop(box).save(dest, quality=95)
    return dest if dest.is_file() else None


def prefer_face_crop_surfaces(
    theme_dir: Path,
    surface_paths: Sequence[Path],
    *,
    max_n: int = 3,
    work_dir: Optional[Path] = None,
) -> List[Path]:
    """SURFACE先頭に面中央クロップを足す（既存パスは維持）。

    種は解決済み SURFACE の先頭（close 等）。phone 固定優先はしない。
    """
    paths = filter_macro_excluded_surfaces(
        [Path(p) for p in surface_paths if Path(p).is_file()]
    )
    if not paths or not generic_plating_core_enabled():
        return paths[: int(max_n or 3)]
    work = Path(work_dir) if work_dir else (Path(theme_dir) / "food_piece_set5_virt" / "deepen_refs")
    # close 名を優先、無ければ先頭
    src = paths[0]
    for p in paths:
        if "close" in p.name.lower():
            src = p
            break
    crop = ensure_surface_face_crop(src, work / "_generic_face_crop.jpg")
    out: List[Path] = []
    if crop and crop.is_file():
        out.append(crop)
    for p in paths:
        if p.resolve() not in {x.resolve() for x in out}:
            out.append(p)
        if len(out) >= int(max_n or 3):
            break
    return out[: int(max_n or 3)]


def should_skip_shape_donors_for_plating(
    *,
    surface_n: int,
    piece_n: int,
) -> bool:
    """§2: SURFACE＋一片があるとき SHAPE は縁・筋の正にしない（不要なら付けない）。"""
    if not generic_plating_core_enabled():
        return False
    return int(surface_n or 0) > 0 and int(piece_n or 0) > 0


def apply_core_to_prompt_text(text: str, *, caller: str = "plating") -> str:
    """危険文掃除＋核ブロック追記。"""
    if not generic_plating_core_enabled():
        return str(text or "")
    base = sanitize_danger_prompt(text)
    core = plating_core_prompt_block(caller=caller)
    if core and core not in base:
        base = (base + "\n\n" + core).strip() if base else core
    return base


def recommend_quality(logic: str) -> str:
    """C7: 判定できる最安段。high は返さない（既定オフ）。

    正本 PLATING §3: 商品固有語で上げない。
    medium は「細い必須属性」または「最終で細部が論点」のときだけ。
    """
    key = str(logic or "").strip().lower()
    medium_keys = (
        "細",
        "微",
        "micro",
        "texture",
        "surface_detail",
        "面詳細",
        "属性微",
        "pale",
        "淡",
        "adopt",
        "合否",
        "plating_final",
        "盛り最終",
    )
    if any(k in key for k in medium_keys):
        return "medium"
    return "low"


def quality_after_effect_pass(*, fine_attr_at_issue: bool) -> str:
    """C8 E5: 効きPASS後の本線 quality。効き前は常に low（呼び出し側で固定）。"""
    return "medium" if fine_attr_at_issue else "low"


def effect_vision_model() -> str:
    return (
        str(os.environ.get("B5_EFFECT_VISION_MODEL") or "").strip() or "gpt-4o-mini"
    )


def _mime_for_path(path: Path) -> str:
    u = path.suffix.lower()
    if u in (".jpg", ".jpeg"):
        return "image/jpeg"
    if u == ".webp":
        return "image/webp"
    if u == ".gif":
        return "image/gif"
    return "image/png"


def _image_content_part(path: Path) -> Dict[str, Any]:
    b64 = base64.b64encode(path.read_bytes()).decode("ascii")
    return {
        "type": "image_url",
        "image_url": {
            "url": f"data:{_mime_for_path(path)};base64,{b64}",
            "detail": "low",
        },
    }


_VISION_FAIL_CLASSES = {
    "multi_unit",
    "wrong_subject",
    "attr_zero",
    "slot_mix",
    "danger_residual",
    "artifact_fake",
}


@dataclass
class VisionJudgeResult:
    """C8 E3 の画像合否。

    keep / next_fix は次の1枚のプロンプトに必ず載せる具体文。
    """

    verdict: str
    msg: str
    fail_class: str
    ng_one_line: str
    keep_one_line: str = ""
    diff_one_line: str = ""
    next_fix: str = ""


def vision_judge_piece_effect(
    candidate: Path,
    *,
    seed_paths: Optional[Sequence[Path]] = None,
    extra_criteria: str = "",
) -> VisionJudgeResult:
    """C8 E3: できた1枚を見て合否（Vision）。

    抽象的な「別物」だけ禁止。色・縁・筋・厚みなど観測可能な差と、
    次生成向けの具体直し文を必ず出させる。
    """
    if not candidate.is_file() or candidate.stat().st_size < 30_000:
        return VisionJudgeResult(
            "FAIL",
            "空",
            FAIL_ATTR_ZERO,
            "画像が空または壊れている",
            "",
            "",
            "有効な1片画像を生成せよ。",
        )

    seeds = [Path(p) for p in (seed_paths or []) if p and Path(p).is_file()]
    criteria = (
        "役割: 生成された「1片見本」1枚を、種と比較して合否し、次の生成指示を書く。\n"
        "目標は開封後1片の見本（画面に1つ）。種が盛り・山でも、種の中の1片の"
        "色・縁・質感・厚み系統に合わせる。盛りレイアウト自体のコピーは求めない。\n"
        "PASS条件:\n"
        "1) 画面内の対象はちょうど1つ（山・ばら・複数は multi_unit）\n"
        "2) 白〜生成り背景。手・皿・袋・文字が目立つなら slot_mix\n"
        "3) 種の1片と同じ素材系統（色味・縁・面の質感・厚み感）。"
        "ここが大きく違うなら wrong_subject\n"
        "4) 必須の面の見えがゼロなら attr_zero\n"
        "5) 作った規則くり返しだけで属性を偽るなら artifact_fake"
        "（種の正が規則縁なら使わない）\n"
        "禁止: 他商品専用の禁止リスト。曖昧な「別物です」だけで終わること。\n"
        "FAIL時は必須:\n"
        "- diffOneLine: 種の1片と比べて何が違うか（色／縁／面／厚み／形から具体）\n"
        "- keepOneLine: 残してよい見え（無ければ「なし」と書かず、少しでも良い点を書く。"
        "どうしても無ければ空）\n"
        "- nextFix: 次の1枚生成への具体指示を2〜4文。"
        "「同じ対象だけ描け」だけの抽象文は不可。観測した差を直す書き方にせよ。\n"
    )
    if extra_criteria.strip():
        criteria += "追加条件:\n" + extra_criteria.strip() + "\n"
    criteria += (
        "JSONのみ返せ:\n"
        '{"verdict":"PASS"|"FAIL","failClass":"multi_unit"|"wrong_subject"|'
        '"attr_zero"|"slot_mix"|"danger_residual"|"artifact_fake"|"",'
        '"ngOneLine":"不合格を一文",'
        '"diffOneLine":"種の1片との具体差",'
        '"keepOneLine":"残す見え",'
        '"nextFix":"次生成の具体指示2〜4文",'
        '"reason":"短い理由"}\n'
    )

    content: List[Dict[str, Any]] = [{"type": "text", "text": criteria}]
    content.append({"type": "text", "text": "【判定対象】次の1枚"})
    content.append(_image_content_part(candidate))
    for i, sp in enumerate(seeds[:3], start=1):
        content.append(
            {
                "type": "text",
                "text": f"【種{i}】参照（盛りでも中の1片の材質を正とする）",
            }
        )
        content.append(_image_content_part(sp))

    try:
        from openai import OpenAI
        from openai_image import load_openai_api_key
    except Exception as e:
        return VisionJudgeResult(
            "FAIL",
            f"vision import失敗: {e}",
            FAIL_ATTR_ZERO,
            "Vision判定不能",
            "",
            "",
            "",
        )

    try:
        client = OpenAI(api_key=load_openai_api_key())
        resp = client.chat.completions.create(
            model=effect_vision_model(),
            messages=[{"role": "user", "content": content}],
            max_tokens=500,
            temperature=0,
        )
        text = ""
        try:
            text = resp.choices[0].message.content or ""
        except Exception:
            text = ""
    except Exception as e:
        return VisionJudgeResult(
            "FAIL",
            f"vision API失敗: {e}",
            FAIL_ATTR_ZERO,
            "Vision判定不能",
            "",
            "",
            "",
        )

    t = re.sub(r"```\w*\n?", "", text or "").strip()
    start, end = t.find("{"), t.rfind("}")
    if start < 0 or end <= start:
        return VisionJudgeResult(
            "FAIL",
            f"vision JSONなし: {(text or '')[:120]}",
            FAIL_ATTR_ZERO,
            "Vision応答不正",
            "",
            "",
            "",
        )
    try:
        obj = json.loads(t[start : end + 1])
    except json.JSONDecodeError:
        return VisionJudgeResult(
            "FAIL", "vision JSON壊", FAIL_ATTR_ZERO, "Vision応答不正", "", "", ""
        )

    verdict = str(obj.get("verdict") or "").strip().upper()
    if verdict not in ("PASS", "FAIL"):
        verdict = "FAIL"
    fc = str(obj.get("failClass") or "").strip()
    if verdict == "PASS":
        fc = ""
    elif fc not in _VISION_FAIL_CLASSES:
        fc = FAIL_MULTI_UNIT if not fc else fc
        if fc not in _VISION_FAIL_CLASSES:
            fc = FAIL_ATTR_ZERO
    ng = str(obj.get("ngOneLine") or "").strip()
    keep = str(obj.get("keepOneLine") or "").strip()
    diff = str(obj.get("diffOneLine") or "").strip()
    nxt = str(obj.get("nextFix") or "").strip()
    reason = str(obj.get("reason") or "").strip()
    if verdict == "FAIL" and not ng:
        ng = diff or reason or "画像合否で不合格"
    if verdict == "FAIL" and not nxt:
        # 最低限の具体寄せ（抽象「別物禁止」だけにしない）
        nxt = (
            (f"差: {diff}。" if diff else "")
            + "種の中の1片に合わせて色味・縁・面の質感・厚み感を寄せよ。"
            "盛り全体のコピーはするな。画面は1片だけ。"
        )
    msg = reason or diff or ("vision PASS" if verdict == "PASS" else f"vision FAIL {fc}")
    return VisionJudgeResult(verdict, msg, fc, ng, keep, diff, nxt)


def effect_verify_max_attempts(default: int = 10) -> int:
    """C8: 1スロットあたり low 効き確認の上限（正本既定10）。"""
    raw = str(os.environ.get("B5_EFFECT_VERIFY_MAX") or "").strip()
    if raw.isdigit():
        return max(1, int(raw))
    return max(1, int(default))


@dataclass
class EffectVerifyAttempt:
    attempt: int
    ng_one_line: str
    fail_class: str
    keep_one_line: str
    diff_one_line: str
    next_fix: str
    prompt_note: str
    quality: str
    verdict: str
    msg: str
    image_path: str = ""


@dataclass
class EffectVerifyResult:
    ok: bool
    effect_passed: bool
    ng_one_line: str
    fail_class: str
    keep_one_line: str = ""
    diff_one_line: str = ""
    next_fix: str = ""
    attempts: List[EffectVerifyAttempt] = field(default_factory=list)
    final_prompt: str = ""
    final_image: Optional[Path] = None
    stopped_reason: str = ""

    def as_dict(self) -> Dict[str, Any]:
        return {
            "ok": self.ok,
            "effectPassed": self.effect_passed,
            "ngOneLine": self.ng_one_line,
            "failClass": self.fail_class,
            "keepOneLine": self.keep_one_line,
            "diffOneLine": self.diff_one_line,
            "nextFix": self.next_fix,
            "stoppedReason": self.stopped_reason,
            "finalImage": str(self.final_image) if self.final_image else "",
            "attempts": [
                {
                    "attempt": a.attempt,
                    "ngOneLine": a.ng_one_line,
                    "failClass": a.fail_class,
                    "keepOneLine": a.keep_one_line,
                    "diffOneLine": a.diff_one_line,
                    "nextFix": a.next_fix,
                    "promptNote": a.prompt_note,
                    "quality": a.quality,
                    "verdict": a.verdict,
                    "msg": a.msg,
                    "imagePath": a.image_path,
                }
                for a in self.attempts
            ],
        }


# 失敗クラス（正本§3.1）。商品名を焼かない。
FAIL_MULTI_UNIT = "multi_unit"
FAIL_WRONG_SUBJECT = "wrong_subject"
FAIL_ATTR_ZERO = "attr_zero"
FAIL_SLOT_MIX = "slot_mix"
FAIL_DANGER_RESIDUAL = "danger_residual"
FAIL_ARTIFACT_FAKE = "artifact_fake"


def prompt_fix_from_vision(
    *,
    fail_class: str,
    ng_one_line: str,
    keep_one_line: str = "",
    diff_one_line: str = "",
    next_fix: str = "",
) -> str:
    """C8: Visionの具体差＋次直しを主にし、クラス定型は補助のみ。"""
    fc = str(fail_class or "").strip() or "generic"
    ng = str(ng_one_line or "").strip() or "今回のNG"
    keep = str(keep_one_line or "").strip()
    diff = str(diff_one_line or "").strip()
    nxt = str(next_fix or "").strip()
    lines: List[str] = []
    if keep:
        lines.append(f"【合格として残す】{keep}")
    if diff:
        lines.append(f"【種の1片との具体差】{diff}")
    lines.append(f"【不合格・禁止・{fc}】{ng}")
    if nxt:
        lines.append(f"【次の1枚への具体直し】{nxt}")
    else:
        # フォールバック（弱 ban のみは避ける）
        lines.append(
            "【次の1枚への具体直し】種の中の1片に色・縁・面質感・厚みを寄せよ。"
            "画面は1片だけ。盛り全体のコピーはするな。"
        )
    return "\n".join(lines)


def prompt_fix_for_fail_class(
    fail_class: str, ng_one_line: str, keep_one_line: str = ""
) -> str:
    """後方互換。新規は prompt_fix_from_vision を使う。"""
    return prompt_fix_from_vision(
        fail_class=fail_class,
        ng_one_line=ng_one_line,
        keep_one_line=keep_one_line,
    )


def run_effect_verify_loop(
    *,
    build_prompt: Callable[[str, str, str, str, str, int], str],
    generate_low: Callable[[str, int], Path],
    judge: Callable[[Path], VisionJudgeResult],
    max_attempts: Optional[int] = None,
    initial_ng: str = "",
    initial_fail_class: str = "",
    initial_keep: str = "",
    initial_diff: str = "",
    initial_next_fix: str = "",
) -> EffectVerifyResult:
    """C8 1枚直列ループ（正本 §3.1）。

    1枚生成→Vision→FAILなら具体差・残す文・次直しを積んでまた1枚。
    build_prompt(ng, fc, keep, diff, next_fix, attempt)
    """
    n_max = effect_verify_max_attempts(max_attempts or 10)
    ng = str(initial_ng or "").strip()
    fc = str(initial_fail_class or "").strip()
    keep = str(initial_keep or "").strip()
    diff = str(initial_diff or "").strip()
    nxt = str(initial_next_fix or "").strip()
    attempts: List[EffectVerifyAttempt] = []
    last_prompt = ""
    last_path: Optional[Path] = None

    for i in range(1, n_max + 1):
        prompt = build_prompt(ng, fc, keep, diff, nxt, i)
        last_prompt = prompt
        note = (
            f"ng={ng or '（初回）'} diff={diff or '-'} keep={keep or '-'} "
            f"fix={('有' if nxt else '-')} class={fc or '-'} a={i}"
        )
        path = generate_low(prompt, i)
        last_path = path
        jr = judge(path)
        if not isinstance(jr, VisionJudgeResult):
            raise TypeError("judge must return VisionJudgeResult")
        verdict = jr.verdict
        msg = jr.msg
        if jr.fail_class:
            fc = jr.fail_class
        if jr.ng_one_line:
            ng = jr.ng_one_line
        if jr.keep_one_line:
            keep = jr.keep_one_line
        if jr.diff_one_line:
            diff = jr.diff_one_line
        if jr.next_fix:
            nxt = jr.next_fix
        attempts.append(
            EffectVerifyAttempt(
                attempt=i,
                ng_one_line=ng,
                fail_class=fc,
                keep_one_line=keep,
                diff_one_line=diff,
                next_fix=nxt,
                prompt_note=note,
                quality="low",
                verdict=verdict,
                msg=msg,
                image_path=str(path),
            )
        )
        if str(verdict).upper() == "PASS":
            return EffectVerifyResult(
                ok=True,
                effect_passed=True,
                ng_one_line=ng,
                fail_class=fc,
                keep_one_line=keep,
                diff_one_line=diff,
                next_fix=nxt,
                attempts=attempts,
                final_prompt=last_prompt,
                final_image=path,
                stopped_reason="",
            )

    return EffectVerifyResult(
        ok=False,
        effect_passed=False,
        ng_one_line=ng,
        fail_class=fc,
        keep_one_line=keep,
        diff_one_line=diff,
        next_fix=nxt,
        attempts=attempts,
        final_prompt=last_prompt,
        final_image=last_path,
        stopped_reason=f"maxAttempts={n_max}",
    )
