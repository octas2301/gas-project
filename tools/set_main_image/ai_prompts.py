# -*- coding: utf-8 -*-
"""モール別プロンプト（見本拘束・楽天はレイヤー専用のため未使用）。"""
from __future__ import annotations


def prompt_amazon(
    *,
    set_count: int,
    is_food: bool,
    has_octas: bool,
    blueprint_file: str = "",
    pattern_hint: str = "",
    layout_intent_ja: str = "",
    has_fact_ref: bool = False,
    fact_asin: str = "",
    ink_fill_min: float | None = None,
    has_unit_base: bool = False,
    octas_tilt_deg: float | None = None,
) -> str:
    n = int(set_count)
    small_n = n <= 4

    food_line = ""
    if is_food and has_octas:
        tilt_txt = (
            f"IMAGE 3 is already rotated about {octas_tilt_deg:+.1f} degrees — keep that tilt; "
            if octas_tilt_deg is not None
            else "IMAGE 3 may already include a slight left/right tilt — keep it; "
        )
        food_line = (
            "OCTAS SEAL (IMAGE 3): composite as a SEPARATE floating seal graphic near the "
            "lower-right of the composition (beside / in front of products), NOT stuck onto "
            "any can label like a product sticker. "
            + tilt_txt
            + "Keep the seal artwork as-is (readable). Do not warp it onto curved metal.\n"
        )
    elif is_food and not has_octas:
        food_line = "Do not invent expiry stickers.\n"

    blueprint_line = (
        f"- Selected LAYOUT BLUEPRINT file: {blueprint_file} "
        f"(patternHint={pattern_hint or 'n/a'}).\n"
        if blueprint_file
        else ""
    )
    intent_line = (
        f"- Human-readable selection reason (JA): {layout_intent_ja}\n"
        if layout_intent_ja
        else ""
    )

    if has_unit_base:
        roles = [
            "- IMAGE 1 = HERO PRODUCT (transparent/cutout). Use for the main/hero unit only.",
            "- IMAGE 1B = UNIT / STOCK PRODUCT (transparent/cutout). Use for the other "
            f"{max(n - 1, 0)} inventory units (same SKU family, usually closed or non-action).",
            "- IMAGE 2 = LAYOUT BLUEPRINT (set MAIN sample). Energy / placement / fill only — "
            "NOT product redesign.",
        ]
    else:
        roles = [
            "- IMAGE 1 = PRODUCT BASE (transparent/cutout). This is the ONLY allowed product appearance.",
            "- IMAGE 2 = LAYOUT BLUEPRINT (set MAIN sample). Use for energy / placement / fill — "
            "NOT to redesign the product.",
        ]
    if is_food and has_octas:
        roles.append(
            "- IMAGE 3 = OCTAS seal asset (already slightly tilted). Floating composite — "
            "not glued onto packaging art."
        )
    if has_fact_ref:
        roles.append(
            f"- IMAGE_FACT = COMPETITOR / REAL-WORLD REFERENCE"
            f"{f' (ASIN {fact_asin})' if fact_asin else ''}. "
            "Use ONLY to verify physical facts (e.g. one pull-tab on the lid, real can geometry). "
            "Do NOT copy competitor branding, layout, or text. Do NOT invent extra tabs/parts."
        )
    roles_block = "\n".join(roles)

    if small_n:
        pattern_guide = (
            f"SMALL SET (N={n}≤4): ALL {n} units MUST be the SAME visual size as the hero "
            "(do NOT shrink supporting cans). Stronger overlap than large-N layouts is OK "
            "to reduce white space, but every unit must remain countable and the hero must "
            "stay readable. Prefer tight cluster / staggered overlap over tiny floating cans."
        )
    else:
        pattern_guide = {
            "hero_right_stack_left": (
                "Hero (largest) on the RIGHT; quantity stack / smaller copies on the LEFT. "
                "Use IMAGE 2 for size hierarchy inspiration, but COUNT and REGULARITY rules below override "
                "any ambiguous pile in the blueprint."
            ),
            "hero_left_stack_right": (
                "Hero (largest) on the LEFT; quantity stack on the RIGHT. "
                "COUNT/REGULARITY rules override ambiguous piles in IMAGE 2."
            ),
            "centered_cluster": (
                "Centered cluster for many units. Still show exactly the set count as clearly countable "
                "units — never a vague mountain that implies extras behind."
            ),
        }.get(pattern_hint or "", "Copy IMAGE 2 geometry faithfully, subject to COUNT rules.")

    unit_src = "IMAGE 1B" if has_unit_base else "IMAGE 1"
    fact_block = f"""
=== HARD CONSTRAINT E — DESIGN / LID TEXT LOCK (anti-hallucination) ===
E0. ABSOLUTE: Do NOT invent, rewrite, or \"improve\" any packaging text, lid print, pull-tab art,
    or label characters. Garbled / fake Japanese on lids is an automatic FAIL.
E1. Closed-can lids and side labels on stock units MUST look exactly like {unit_src}
    (and IMAGE_FACT only if provided for physical cross-check).
E2. Allowed pixel sources for product appearance: IMAGE 1
    {", IMAGE 1B" if has_unit_base else ""}{", IMAGE_FACT (facts only)" if has_fact_ref else ""}.
    FORBIDDEN sources for redesign: IMAGE 2 blueprint product art, memory, or invented text.
E3. If a lid/label region is unclear, keep the provided base pixels — never hallucinate new wording.
E4. IMAGE_FACT (if present): use ONLY to verify physical facts (e.g. one pull-tab). Do NOT copy
    competitor layout/props. Do NOT invent a second pull-tab.
E5. Artwork priority: hero art from IMAGE 1; stock unit art from {unit_src}.
"""

    if ink_fill_min is None:
        try:
            from fill_metrics import fill_target_for_n

            fill_min, _band = fill_target_for_n(n)
        except Exception:
            fill_min = 0.48 if n <= 3 else (0.52 if n <= 6 else 0.58)
    else:
        fill_min = float(ink_fill_min)
    fill_pct = int(round(fill_min * 100))

    aspect_lock = """
=== HARD CONSTRAINT G — ASPECT RATIO LOCK (strictest) ===
G1. FORBIDDEN: changing width/height ratio of any product (stretch, squash, tall oval lids,
    fat cans, perspective squash that alters the can's native proportions).
G2. Scale MUST be uniform (same factor on X and Y). Paste-like compositing only.
G3. A circular can rim must stay circular (or the same ellipse as in the source photo) —
    never become a taller/skinnier oval than the source.
G4. If fill/SEO pressure conflicts with aspect lock, KEEP aspect — do not distort to fill.
"""

    if has_unit_base:
        d_block = f"""=== HARD CONSTRAINT D — HERO / UNIT BASE LOCK (paste, do not redraw) ===
D1. Hero = pixel-faithful paste of IMAGE 1 (ヒーロー). Stock units = pixel-faithful paste of IMAGE 1B (単体).
D2. Treat this as PHOTO COMPOSITING, not illustration. Do not re-render metal, food, or print.
D3. FORBIDDEN: stretch/squash, warp, perspective fake that changes proportions, redraw labels,
    reinvent lid text, change package design, or AI-redesign so it looks \"enhanced\".
D4. Allowed only: uniform scale + translate + soft contact shadow.
    Optional tiny 2D rotation ONLY if aspect ratio of the product bitmap stays identical.
D5. Do not invent a spoon/action on IMAGE 1B units if IMAGE 1B has none.
D6. Lid / top print on IMAGE 1B units must match IMAGE 1B exactly — no new characters.
"""
    else:
        d_block = """=== HARD CONSTRAINT D — HERO / MAIN PRODUCT = BASE LOCK (paste, do not redraw) ===
D1. Every product unit MUST be a pixel-faithful paste of IMAGE 1 (compositing, not redraw).
D2. FORBIDDEN: stretch/squash (aspect change), warp, perspective fake that changes proportions,
    redraw labels/lid text, reinvent packaging, or AI-redesign.
D3. Allowed only: uniform scale + translate + soft contact shadow; tiny 2D rotation only if
    the product bitmap aspect ratio stays identical to IMAGE 1.
D4. Supporting units are copies of IMAGE 1 under the same lock — no redesigned art.
D5. Keep IMAGE 1's open/closed look; do not invent a different lid or open-can rendering.
"""

    if small_n:
        size_overlap = f"""
=== HARD CONSTRAINT F — SMALL SET N≤4 (same on-canvas size + overlap) ===
F1. CRITICAL: ALL {n} units must have the SAME on-canvas scale as the hero.
    Stock/unit cans must NOT look smaller due to fake perspective or \"depth\".
    Use the SAME uniform scale factor for hero and every IMAGE 1B (or IMAGE 1) copy.
F2. Depth may be shown by overlap / slight vertical offset ONLY — never by shrinking stock units.
F3. Overlap MAY be stronger than for N≥5 to reduce empty white — every unit still countable.
F4. Hero remains the focal unit by placement/overlap order, NOT by being drawn larger.
"""
        c11 = (
            f"For N={n}≤4: identical on-canvas size for every unit (no smaller stock cans); "
            "stronger overlap OK; do NOT bury the hero past readability."
        )
    else:
        size_overlap = ""
        c11 = (
            "Hero may be slightly larger; stock units may be slightly smaller. "
            "While filling the frame, do NOT bury the hero. Aspect ratio still locked."
        )

    unit_lock_note = (
        "hero=IMAGE1, stock units=IMAGE1B"
        if has_unit_base
        else "all units aspect-locked to IMAGE 1 base"
    )

    return f"""You are a STRICT layout-transfer compositor for Amazon Japan MAIN images.

Images arrive interleaved with labels. Roles:
{roles_block}
{blueprint_line}{intent_line}
{d_block}{aspect_lock}{fact_block}{size_overlap}
=== HARD CONSTRAINT A — EXACT COUNT (Amazon compliance) ===
1. Show EXACTLY {n} physical product units. Not {n - 1}, not {n + 1}, not \"about {n}\".
2. A shopper must be able to COUNT every unit without guessing.
3. FORBIDDEN: compositions that look like more items are stacked BEHIND / deeper in the pile
   (human imagination of hidden cans). Avoid deep opaque piles, tunnel stacks, or mountain heaps
   where the rear implies extra inventory.
4. Overlap is OK only if EVERY unit's silhouette (or lid/top) remains clearly countable.
5. If IMAGE 2 looks denser than {n}, SIMPLIFY — do not copy ambiguous depth.

=== HARD CONSTRAINT B — REGULARITY ===
6. Columns / rows of the same product must be REGULAR (even spacing, aligned axes, consistent tilt).
7. FORBIDDEN: a vertical column that zig-zags irregularly or looks accidentally messy.
8. Prefer clean stacks, arcs, or grids with intentional rhythm — not random scatter.

=== HARD CONSTRAINT C — MOVEMENT + FILL (SEO / thumbnail) ===
9. FILL TARGET (tuning knob #1 — ink occupancy): non-white product area must cover at least
   {fill_pct}% of the square canvas (inkFillRatio ≥ {fill_min:.2f}).
   Especially for small set counts ({n} units): SCALE UP and place products so large empty
   white regions (top / sides) disappear. Tiny products floating in a sea of white = FAIL.
10. Add controlled \"movement\" so the listing stands out among flat catalog shots:
    - Tall/slim products (bottles, jars): tasteful diagonal tilts / inward lean while keeping
      each unit's aspect ratio locked.
    - Short/wide products (cans): gentle fan, arc, or slight rotation — NOT a chaotic dump;
      never stretch the can.
11. {c11}
12. Pattern guide (secondary to hard constraints): {pattern_guide}

=== OTHER ===
13. Final background must be pure white (#FFFFFF).
14. Soft contact shadow only; no banners, prices, watermarks, or extra props.
15. Do not copy IMAGE 2's product branding — only layout energy / hierarchy ideas.
{food_line}
OUTPUT: one square Amazon MAIN photo with EXACTLY {n} clearly countable units,
{unit_lock_note}.
"""


def prompt_rakuten_forbidden_fullgen() -> str:
    """楽天フル生成は禁止。呼び出し側で使わないこと。"""
    return (
        "Rakuten must use pixel-preserving layer compose only; "
        "full image generation/editing is forbidden."
    )
