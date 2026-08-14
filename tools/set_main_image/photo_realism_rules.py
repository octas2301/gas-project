# -*- coding: utf-8 -*-
"""
サブ画像向け・写真実写ルール（進化用の正本）。

- compose が毎回ここを読む。日々の改善は本ファイルの RULES_JA / VERSION を更新する。
- 「最新の写真に見える」ことを優先。AIっぽい均一照明・全面ピンシャープを禁止。
"""
from __future__ import annotations

from typing import Any, Dict

# ルール改定時は VERSION を上げ、CHANGE に1行追記する
VERSION = "2026-08-10.2"

# 案A: EC向け固定カメラプロファイル（機種別ボケ・光学を引き出す）
CAMERA_PROFILE = {
    "body": "Canon EOS R5",
    "lens": "50mm f/1.8",
    "iso": "200-400",
}

RULES_JA = """
【写真実写ロック／PHOTO_REALISM（常に最新の写真に見えること）】
目的: 生成結果を「スタジオ実写のECカット」に見せる。CGI・均一AI合成感を排除する。

■ カメラ実写ロック／CAMERA_LOOK（偽物輪郭対策・案A）
- 撮影前提: Canon EOS R5 + 50mm f/1.8、ISO 200–400、浅い被写界深度（フルフレーム実写）。
- 「photorealistic / 8K / hyper-detailed / ultra-sharp everywhere」は使わず、上記カメラ言語で光学を表現する。
- 主役（商品パッケージの主面、または料理の主フォーカス面）のみシャープ。
- 背景・小物・布・テーブル縁・遠い料理面は creamy / smooth bokeh（円形ハイライトの玉ボケ可）。
- 被写体の輪郭は切り抜きパスのように硬くしない。わずかな光学ソフトネス・フォーカスフォールオフを入れる。
- 全面ピンシャープのコラージュ／ベクター縁取り感は禁止。
- subtle luminance grain（センサー粒感）を残す。ゼロノイズのプラスチック平滑は禁止。
- ラベル可読性: 正本パッケージの主面文字は読める範囲でシャープ。それ以外の縁は柔らかくてよい。

■ 光・反射（PACKAGE_TRUTHの光学を踏襲）
- IMAGE_PACKAGE_TRUTH に写っている光源方向・ハイライト位置・縁の反射・影の落ち方を維持する。
- フラットな全面照明・のっぺりしたプラスチック感・金属/ガラスの反射を消すことは禁止。
- 缶・瓶・ガラスは環境反射とスペキュラを自然に残す（塗りつぶし禁止）。

■ 被写界深度・ピンボケ（複数被写体で特に必須）
- 距離の違う被写体・前景・背景は、実写同様に距離に応じたボケを入れる。
- 瓶＋丼など複数素材: 前後差のある自然なピンボケ・ボケの玉・輪郭のにじみを必須。両方を同じ鋭さで全面シャープにするのは禁止。
- 背景・布・テーブル奥は柔らかく落とす。

■ 陰影・空気感
- 接地影・接触影を弱く自然に。浮遊感のある切り貼り禁止。
- 湯気・粒子・ふりかけは線画CG禁止。写真的なボケ・粒子・半透明の塊で描く。

■ 質感
- 食品は適度な照りに留め、過剰なオイルテカリ・均一ハイライトを避ける。
- ノイズをゼロにしすぎない（わずかな写真粒感は可）。過度なスムージング禁止。

■ HARD BAN（AI感の典型）
- 全面ピンシャープのコラージュ感・切り抜きエッジ
- 均一スタジオライトだけで影が無い
- ガラス/缶の反射が消えたマット塗り
- 蒸気や粉がベクター線のように整いすぎている
- hyper-detailed / 8K / ultra-sharp everywhere / perfect edge cutout
"""

RULES_EN = """
CAMERA_LOOK (EC product secondary image):
shot on Canon EOS R5, 50mm f/1.8, ISO 200-400, shallow depth of field.
Hero package face (or main food plane) sharp enough to read label; props, fabric, table edges, background: creamy smooth bokeh.
No cutout-hard outlines; slight optical softness / focus falloff on non-hero edges; subtle sensor grain.
Avoid: hyper-detailed, 8K, ultra-sharp everywhere, CGI collage, perfect vector edges, flat plastic lighting.
Keep PACKAGE_TRUTH specular highlights and reflections.
"""

CHANGELOG = [
    "2026-08-10.2: CAMERA_LOOK 案A（Canon R5 50mm f/1.8・creamy bokeh・輪郭ソフトネス・粒感・8K系禁止）。",
    "2026-08-10.1: 初版。DOF/距離ボケ/反射/陰影。複数被写体のピンボケ必須。",
]


def photo_realism_block_ja() -> str:
    return f"{RULES_JA.strip()}\n（rulesVersion={VERSION}）"


def photo_realism_block_en() -> str:
    return f"{RULES_EN.strip()}\n(rulesVersion={VERSION})"


def photo_realism_meta() -> Dict[str, Any]:
    return {
        "version": VERSION,
        "cameraProfile": dict(CAMERA_PROFILE),
        "changelog": list(CHANGELOG),
    }
