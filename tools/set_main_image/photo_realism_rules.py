# -*- coding: utf-8 -*-
"""
サブ画像向け・写真実写ルール（進化用の正本）。

- compose が毎回ここを読む。日々の改善は本ファイルの RULES_JA / VERSION を更新する。
- 「最新の写真に見える」ことを優先。AIっぽい均一照明・全面ピンシャープを禁止。
"""
from __future__ import annotations

from typing import Any, Dict

# ルール改定時は VERSION を上げ、CHANGE に1行追記する
VERSION = "2026-11-03.1"

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
- 包材を置いた面と、カメラから同じ距離にあるもの（包材も、同じ距離の料理・小物も）を同じシャープさにする。包材全体をぼかすな。
- ぼかすのはカメラから遠いものだけ（奥の壁・奥の小物・遠い縁）。creamy / smooth bokeh（円形ハイライトの玉ボケ可）。
- 被写体の輪郭は切り抜きパスのように硬くしない。包材は外側の縁だけ、同じ距離の隣の物の縁と同じ柔らかさにする。印字はぼかすな。
- 全面ピンシャープのコラージュ／ベクター縁取り感は禁止。
- subtle luminance grain（センサー粒感）を残す。ゼロノイズのプラスチック平滑は禁止。
- ラベル可読性: 正本パッケージの主面文字は読める範囲でシャープ。それ以外の縁は柔らかくてよい。

■ 光・反射（場面の光）
- 光は画面で1つ。袋・食品・皿は同じ方向。白抜きの袋写真の照明を写すな。
- 包材の影の方向と濃さは、同じ面の食品・皿と同じ。包材だけ明るくしない。包材だけ影を薄くしない。
- 影を合わせるために、包材の輪郭を直線に伸ばすな。ファイルの縁のうねり・しわを残せ。影は包材の下端の形に沿わせる。
- 包材は盛り付け（器・食品）の後ろに置く。重なったところは、前の器や食品が包材を隠す。器を透明にして後ろの包材を透かすな。
- 包材の印刷面は、ファイルにある窓以外は透明にするな。袋・缶・瓶・箱のどれでも同じ。
- 包材は置いた面の光・接地の影・縁の反射に乗せる。浮いた切り抜きにするな。
- フラットな全面照明・のっぺりしたプラスチック感・金属/ガラスの反射を消すことは禁止。
- 缶・瓶・ガラスは環境反射とスペキュラを自然に残す（塗りつぶし禁止）。

■ 被写界深度・ピンボケ（複数被写体で特に必須）
- 距離の違う被写体・前景・背景は、実写同様に距離に応じたボケを入れる。
- 瓶＋丼など複数素材: 前後差のある自然なピンボケ・ボケの玉・輪郭のにじみを必須。両方を同じ鋭さで全面シャープにするのは禁止。
- 背景・布・テーブル奥は柔らかく落とす。

■ 陰影・空気感
- 接地影・接触影は、同じ面の食品と同じ濃さ。浮遊感のある切り貼り禁止。
- 湯気・粒子・ふりかけは線画CG禁止。写真的なボケ・粒子・半透明の塊で描く。

■ 質感
- 食品は適度な照りに留め、過剰なオイルテカリ・均一ハイライトを避ける。
- ノイズをゼロにしすぎない（わずかな写真粒感は可）。過度なスムージング禁止。

■ 物・背景・小物の実写（人物と同水準／必須）
- 人物だけ実写で、皿・グラス・テーブル・布・壁・家具がCG／均一塗りになるのは不合格。
- 小物: 微細な傷・指紋・埃・結露・使い込まれた質感をごく自然に。新品カタログのテカテカ禁止。
- 食器・ガラス: 環境反射とスペキュラを残す。マット塗りつぶし禁止。
- 食品小片・削り節等: 厚み・繊維・影のムラ。プラスチック粘土感禁止。
- 背景: soft side light の陰影。壁・棚・カーテンにわずかな色ムラ・ボケ。スタジオ紙背景ののっぺり禁止。
- 布・服・ナプキン: 自然なシワと接触影。アイロン完璧のCGI布禁止。

■ HARD BAN（AI感の典型）
- 全面ピンシャープのコラージュ感・切り抜きエッジ
- 均一スタジオライトだけで影が無い
- ガラス/缶の反射が消えたマット塗り
- 蒸気や粉がベクター線のように整いすぎている
- hyper-detailed / 8K / ultra-sharp everywhere / perfect edge cutout
- 人物だけリアルで小物・背景がフラットCGI
- 人物の手・腕・肩が袋の輪郭から生える／袋を貫通する／袋の天面から出る（手前の袋の上に人物部位を重ねない）
"""

RULES_EN = """
CAMERA_LOOK (EC product secondary image):
shot on Canon EOS R5, 50mm f/1.8, ISO 200-400, shallow depth of field.
Packages and anything at the same camera distance share one focus. Blur only what is farther from the camera.
No cutout-hard outlines; slight optical softness / focus falloff on distant edges; subtle sensor grain.
Avoid: hyper-detailed, 8K, ultra-sharp everywhere, CGI collage, perfect vector edges, flat plastic lighting.
Light comes from the scene. Do not copy the white-cutout photo's lighting. Keep contact shadow and edge reflection.
"""

CHANGELOG = [
    "2026-11-03.1: HARD BANに「手・腕が袋から生える／袋を貫通」を追加。",
    "2026-09-29.1: 包材は盛り付けの後ろ。窓以外は不透明。縁だけ隣と同じ柔らかさ。影は下端に沿う。",
    "2026-09-27.2: 光は1つ。袋の影の方向と濃さは同じ面の食品と同じ。",
    "2026-09-27.1: 同じ距離は同じピント。遠いものだけぼかす。光は場面。白抜き照明は写さない。",
    "2026-08-30.1: 物・背景・小物の実写（人物と同水準）。CGI小物／のっぺり背景を HARD BAN。",
    "2026-08-10.2: CAMERA_LOOK 案A（Canon R5 50mm f/1.8・creamy bokeh・輪郭ソフトネス・粒感・8K系禁止）。",
    "2026-08-10.1: 初版。DOF/距離ボケ/反射/陰影。複数被写体のピンボケ必須。",
]


def photo_realism_block_ja(photo_arm: str = "") -> str:
    """既定は RULES_JA のまま。photo_arm は試験の1条件だけ差し替える。"""
    text = RULES_JA.strip()
    arm = str(photo_arm or "").strip()
    if arm == "drop_dof":
        text = text.replace(
            "Canon EOS R5 + 50mm f/1.8、ISO 200–400、浅い被写界深度（フルフレーム実写）。",
            "Canon EOS R5 + 50mm f/1.8、ISO 200–400。包材と料理にピントを合わせる。ぼかすのは遠い背景だけ。",
        )
        text = text.replace(
            "ぼかすのはカメラから遠いものだけ（奥の壁・奥の小物・遠い縁）。creamy / smooth bokeh（円形ハイライトの玉ボケ可）。",
            "ぼかすのはカメラから遠い背景だけ。手前の包材と料理はぼかすな。",
        )
        text = text.replace(
            "■ 被写界深度・ピンボケ（複数被写体で特に必須）",
            "■ ピント",
        )
        text = text.replace(
            "距離の違う被写体・前景・背景は、実写同様に距離に応じたボケを入れる。\n"
            "- 瓶＋丼など複数素材: 前後差のある自然なピンボケ・ボケの玉・輪郭のにじみを必須。両方を同じ鋭さで全面シャープにするのは禁止。",
            "包材と料理はピントを合わせる。ぼかすのは遠い背景だけ。",
        )
    elif arm == "mm85":
        text = text.replace("50mm f/1.8", "85mm f/1.8")
    elif arm:
        raise SystemExit(f"photo_arm は drop_dof か mm85（指定={arm!r}）")
    return f"{text}\n（rulesVersion={VERSION}）"


def photo_realism_block_en() -> str:
    return f"{RULES_EN.strip()}\n(rulesVersion={VERSION})"


def photo_realism_meta() -> Dict[str, Any]:
    return {
        "version": VERSION,
        "cameraProfile": dict(CAMERA_PROFILE),
        "changelog": list(CHANGELOG),
    }
