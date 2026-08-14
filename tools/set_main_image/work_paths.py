# -*- coding: utf-8 -*-
"""Drive「05.画像生成（セットMAIN）」配下の解決。"""
from __future__ import annotations

import logging
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Tuple

LOG = logging.getLogger("set_main_image.paths")

DEFAULT_WORK_ROOT = Path(r"G:/マイドライブ/05.画像生成（セットMAIN）")

# 楽天マトリクス／人間目視の正（サブ・MAINともここ）
DEFAULT_RAKUTEN_UPLOAD_DIR = Path(
    r"G:/マイドライブ/03.楽天・Yahoo!商品登録（CSV一括UL）/02.楽天アップロード画像保存場所"
)

# セットMAIN Bメニューと同系のマスタCSV既定（古ければ c1_fetch_inputs で再取得）
DEFAULT_MASTER_CSV = Path(
    r"C:/Users/takuy/Downloads/Lv4_Amazon_PACKAGED/input/master_export.csv"
)

SUB_AMAZON = "01.amazon白抜きベース"
SUB_RAKUTEN = "02.楽天ベース"
SUB_RAKUTEN_DONE = "処理済み"
SUB_AMAZON_REF = "03.amazon見本"
SUB_RAKUTEN_REF = "04.楽天見本"
SUB_DIGIT = "97.楽天数字レイヤ"
SUB_KINMARU = "98.楽天金丸素材"
SUB_OCTAS = "99.octas期限管理シール素材"
SUB_TEST_OUT = "00.テスト出力"
SUB_META = "_meta"

# C→D 本線の Amazon MAIN 置き場（画像のみ。メタは _meta/）
DEFAULT_AMAZON_07_OUT = Path(
    r"G:/マイドライブ/04.amazonカタログ作成（CSV一括UL）/07.白抜きの置き場（人間が入れる）"
)

# 量産アスペクト（Cメニュー選択）
ASPECT_SQUARE = "square"
ASPECT_PORTRAIT = "portrait"
ASPECT_LANDSCAPE = "landscape"
ASPECT_CHOICES = (ASPECT_SQUARE, ASPECT_PORTRAIT, ASPECT_LANDSCAPE)
ASPECT_LABEL_JA = {
    ASPECT_SQUARE: "正方形",
    ASPECT_PORTRAIT: "縦長タイプ（合格固定）",
    ASPECT_LANDSCAPE: "横長",
}

# 人間が01直下で付ける指定タグ（ファイル名に含める）
TAG_HERO = "ヒーロー"
TAG_UNIT = "単体"


@dataclass
class AmazonProductBases:
    """Amazon 01ベース。ヒーロー／単体の使い分け。"""

    hero: Path
    unit: Path
    mode: str  # paired_hero_unit | single_legacy
    pair_key: str = ""

    def to_dict(self) -> dict:
        return {
            "hero": str(self.hero),
            "unit": str(self.unit),
            "mode": self.mode,
            "pairKey": self.pair_key,
            "heroName": self.hero.name,
            "unitName": self.unit.name,
            "sameFile": self.hero.resolve() == self.unit.resolve()
            if self.hero.is_file() and self.unit.is_file()
            else self.hero == self.unit,
        }


def meta_dir_for(out_dir: Path) -> Path:
    """画像直下を汚さないメタ置き場（_meta）。"""
    d = Path(out_dir) / SUB_META
    d.mkdir(parents=True, exist_ok=True)
    return d


def default_work_root() -> Path:
    return DEFAULT_WORK_ROOT


def _list_images(folder: Path) -> List[Path]:
    if not folder.is_dir():
        return []
    out: List[Path] = []
    for p in folder.iterdir():
        if p.suffix.lower() in (".png", ".jpg", ".jpeg", ".webp"):
            out.append(p)
    # 本線は透過PNG優先（同名があれば png > webp > jpg）
    rank = {".png": 0, ".webp": 1, ".jpg": 2, ".jpeg": 2}

    def _key(p: Path):
        return (rank.get(p.suffix.lower(), 9), -p.stat().st_mtime, p.name.lower())

    return sorted(out, key=_key)


def _pick(folder: Path, parent_sku: str, explicit: Optional[Path]) -> Path:
    if explicit and explicit.is_file():
        return explicit
    imgs = _list_images(folder)
    if not imgs:
        raise FileNotFoundError(f"画像がありません: {folder}")
    want = (parent_sku or "").strip().lower()
    if want:
        for p in imgs:
            if want in p.name.lower():
                return p
    LOG.warning(
        "親SKU名がファイル名に無いため最新を使用 folder=%s file=%s",
        folder.name,
        imgs[0].name,
    )
    return imgs[0]


def resolve_amazon_base(work_root: Path, parent_sku: str, explicit: Optional[Path] = None) -> Path:
    """互換: ヒーロー優先（単体のみなら単体）。明示パスがあればそれ。"""
    bases = resolve_amazon_product_bases(work_root, parent_sku, explicit)
    return bases.hero


def _strip_role_tags(stem: str) -> str:
    s = stem
    for tag in (TAG_HERO, TAG_UNIT):
        s = s.replace(tag, "")
    return s.strip(" _-\t")


def resolve_amazon_product_bases(
    work_root: Path,
    parent_sku: str = "",
    explicit: Optional[Path] = None,
) -> AmazonProductBases:
    """
    01.amazon白抜きベース直下から Amazon 貼付素材を解決。

    方針（2026-08）: **ヒーロー画像は使わない**。メインも在庫個も『単体』のみ。
    - 『単体』タグがあればそれを hero=unit に使う
    - 単体が無くヒーローだけ → 警告のうえヒーローを流用（フォールバック）
    - --base 明示時 → その1枚を両方に使う
    """
    if explicit and explicit.is_file():
        return AmazonProductBases(
            hero=explicit, unit=explicit, mode="unit_only", pair_key=""
        )

    folder = work_root / SUB_AMAZON
    imgs = _list_images(folder)
    if not imgs:
        raise FileNotFoundError(f"画像がありません: {folder}")

    heroes = [p for p in imgs if TAG_HERO in p.stem]
    units = [p for p in imgs if TAG_UNIT in p.stem]

    def _prefer(cands: List[Path]) -> Optional[Path]:
        if not cands:
            return None
        want = (parent_sku or "").strip().lower()
        if want:
            for p in cands:
                if want in p.name.lower():
                    return p
        return sorted(cands, key=lambda x: x.stat().st_mtime, reverse=True)[0]

    if units:
        u = _prefer(units)
        assert u is not None
        LOG.info(
            "amazon bases UNIT_ONLY (hero素材中止) unit=%s",
            u.name,
        )
        return AmazonProductBases(
            hero=u,
            unit=u,
            mode="unit_only",
            pair_key=_strip_role_tags(u.stem),
        )

    if heroes:
        h = _prefer(heroes)
        assert h is not None
        LOG.warning(
            "単体タグ無し — フォールバックでヒーローを使用: %s",
            h.name,
        )
        return AmazonProductBases(
            hero=h, unit=h, mode="single_legacy", pair_key=_strip_role_tags(h.stem)
        )

    # タグなし従来
    one = _pick(folder, parent_sku, None)
    return AmazonProductBases(hero=one, unit=one, mode="unit_only", pair_key="")


def resolve_rakuten_base(work_root: Path, parent_sku: str, explicit: Optional[Path] = None) -> Path:
    return _pick(work_root / SUB_RAKUTEN, parent_sku, explicit)


def move_rakuten_base_to_processed(base_path: Path, work_root: Optional[Path] = None) -> Optional[Path]:
    """
    量産成功後、02.楽天ベース直下の使用済みベースを 処理済み/ へ移動。
    既に処理済み内・02外・欠落ファイルはスキップ。
    """
    src = Path(base_path)
    if not src.is_file():
        LOG.warning("処理済み移動スキップ（ファイル無し）: %s", src)
        return None

    rakuten_dir = (work_root or default_work_root()) / SUB_RAKUTEN
    try:
        src_resolved = src.resolve()
        rakuten_resolved = rakuten_dir.resolve()
    except OSError:
        src_resolved = src
        rakuten_resolved = rakuten_dir

    # 02直下のみ（処理済み内や別パスは動かさない）
    if src_resolved.parent != rakuten_resolved:
        LOG.info(
            "処理済み移動スキップ（02直下以外）: %s",
            src,
        )
        return None

    dest_dir = rakuten_dir / SUB_RAKUTEN_DONE
    dest = unique_path_in_dir(dest_dir, src.name)

    src.rename(dest)
    LOG.info("rakuten base → 処理済み: %s → %s", src.name, dest.name)
    return dest


def unique_path_in_dir(dest_dir: Path, filename: str) -> Path:
    """
    同名があれば stem_2, stem_3…、それでもダメならタイムスタンプ付き。
    ヒーロー／単体のリネーム運用で処理済みが衝突しないようにする。
    """
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / filename
    if not dest.exists():
        return dest
    stem = Path(filename).stem
    suf = Path(filename).suffix
    for i in range(2, 100):
        cand = dest_dir / f"{stem}_{i}{suf}"
        if not cand.exists():
            LOG.info("処理済み名衝突回避: %s → %s", filename, cand.name)
            return cand
    stamped = dest_dir / f"{stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{suf}"
    LOG.info("処理済み名衝突回避(timestamp): %s → %s", filename, stamped.name)
    return stamped


def move_amazon_base_to_processed(base_path: Path, work_root: Optional[Path] = None) -> Optional[Path]:
    """
    量産成功後、01.amazon白抜きベース直下の使用済みベースを 処理済み/ へ移動。
    同名がある場合は名前を変えて保存（上書きしない）。
    """
    src = Path(base_path)
    if not src.is_file():
        LOG.warning("Amazon処理済み移動スキップ（ファイル無し）: %s", src)
        return None

    amazon_dir = (work_root or default_work_root()) / SUB_AMAZON
    try:
        src_resolved = src.resolve()
        amazon_resolved = amazon_dir.resolve()
    except OSError:
        src_resolved = src
        amazon_resolved = amazon_dir

    if src_resolved.parent != amazon_resolved:
        LOG.info("Amazon処理済み移動スキップ（01直下以外）: %s", src)
        return None

    dest_dir = amazon_dir / SUB_RAKUTEN_DONE  # 同じ「処理済み」名
    dest = unique_path_in_dir(dest_dir, src.name)

    src.rename(dest)
    LOG.info("amazon base → 処理済み: %s → %s", src.name, dest.name)
    return dest


def move_amazon_product_bases_to_processed(
    bases: AmazonProductBases, work_root: Optional[Path] = None
) -> List[Path]:
    """ヒーロー／単体の両方を処理済みへ（同一ファイルなら1回だけ）。衝突時はリネーム。"""
    moved: List[Path] = []
    seen = set()
    for p in (bases.hero, bases.unit):
        try:
            key = str(Path(p).resolve())
        except OSError:
            key = str(p)
        if key in seen:
            continue
        seen.add(key)
        out = move_amazon_base_to_processed(p, work_root)
        if out is not None:
            moved.append(out)
    return moved


def resolve_digit_layer(work_root: Path, parent_sku: str, explicit: Optional[Path] = None) -> Optional[Path]:
    """任意。空フォルダなら None（書体＋default_digit_box で描画）。"""
    if explicit and explicit.is_file():
        return explicit
    folder = work_root / SUB_DIGIT
    imgs = _list_images(folder)
    if not imgs:
        return None
    want = (parent_sku or "").strip().lower()
    if want:
        for p in imgs:
            if want in p.name.lower():
                return p
    return imgs[0]


def resolve_octas(work_root: Path, explicit: Optional[Path] = None) -> Optional[Path]:
    if explicit and explicit.is_file():
        return explicit
    imgs = _list_images(work_root / SUB_OCTAS)
    return imgs[0] if imgs else None


def resolve_kinmaru(work_root: Path, explicit: Optional[Path] = None) -> Optional[Path]:
    """金丸単体。02に金丸込みなら不要。"""
    if explicit and explicit.is_file():
        return explicit
    imgs = _list_images(work_root / SUB_KINMARU)
    return imgs[0] if imgs else None


def resolve_amazon_reference(
    work_root: Path, parent_sku: str = "", explicit: Optional[Path] = None
) -> Path:
    """配置・大きさ・重なりの見本（03）。"""
    return _pick(work_root / SUB_AMAZON_REF, parent_sku, explicit)


def resolve_rakuten_reference(
    work_root: Path, parent_sku: str = "", explicit: Optional[Path] = None
) -> Path:
    """書体・金丸内バランスの見本（04）。"""
    return _pick(work_root / SUB_RAKUTEN_REF, parent_sku, explicit)
