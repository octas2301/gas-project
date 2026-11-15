# -*- coding: utf-8 -*-
"""PCカメラで面アップを撮って SURFACE に登録する（テスト用）。

操作:
  画面下の「撮影」ボタンをクリック … 枠内を保存
  Space / s … 同上（キーボード）
  [ / ]     … ピントを近く / 遠く（マニュアル）
  a         … オートフォーカス ON/OFF
  q / Esc   … 終了
  c         … カメラ番号を切り替え（0→1→2…）

保存先（既定）:
  theme_new_run_{jan}/food_piece_set5_virt/deepen_refs/ref_surface_cam_{時刻}.jpg

例:
  python b5_surface_webcam_capture.py --jan 4906283047410
  python b5_surface_webcam_capture.py --jan 4906283047410 --camera 0
"""
from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

SCRIPT = Path(__file__).resolve().parent

# ガイド枠（この中に被写体を入れて撮影）
GUIDE_MARGIN = 0.18
# 画面下の撮影ボタン（相対座標・表示フレーム基準）
BTN_H = 56
BTN_PAD = 10


def _theme_dir(jan: str) -> Path:
    from food_unit_canon import theme_new_dir_for_jan

    return Path(theme_new_dir_for_jan(str(jan)))


def _out_dir(jan: str, *, custom: Path | None) -> Path:
    if custom is not None:
        d = Path(custom)
        d.mkdir(parents=True, exist_ok=True)
        return d
    d = _theme_dir(jan) / "food_piece_set5_virt" / "deepen_refs"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _guide_xy(w: int, h: int) -> tuple[int, int, int, int]:
    m = GUIDE_MARGIN
    return int(w * m), int(h * m), int(w * (1 - m)), int(h * (1 - m))


def _btn_xy(w: int, h: int) -> tuple[int, int, int, int]:
    """撮影ボタン矩形 (x0,y0,x1,y1)。"""
    bw = min(220, max(140, w // 3))
    x0 = (w - bw) // 2
    y0 = h - BTN_H - BTN_PAD
    return x0, y0, x0 + bw, y0 + BTN_H


def _apply_near_focus(cap, *, autofocus: bool, focus_val: float) -> None:
    """近い面アップ向けにフォーカスを寄せる（機種により無視される）。"""
    try:
        cap.set(cv2.CAP_PROP_BUFFERSIZE, 1)
    except Exception:
        pass
    # 連写バッファを減らす
    try:
        if autofocus:
            cap.set(cv2.CAP_PROP_AUTOFOCUS, 1)
        else:
            cap.set(cv2.CAP_PROP_AUTOFOCUS, 0)
            # 多くのUSBカメラは値が小さいほど近い
            cap.set(cv2.CAP_PROP_FOCUS, float(focus_val))
    except Exception:
        pass
    # 近接寄りのヒント（対応機のみ）
    for prop, val in (
        (getattr(cv2, "CAP_PROP_ZOOM", 13), 0),
        (getattr(cv2, "CAP_PROP_FOCUS", 28), focus_val if not autofocus else -1),
    ):
        if val < 0:
            continue
        try:
            cap.set(prop, float(val))
        except Exception:
            pass


def main() -> int:
    ap = argparse.ArgumentParser(description="Webcam → SURFACE登録（テスト）")
    ap.add_argument("--jan", required=True, help="JAN（theme_new_run 解決用）")
    ap.add_argument("--camera", type=int, default=0, help="カメラ番号 既定0")
    ap.add_argument(
        "--out-dir",
        type=Path,
        default=None,
        help="保存先（省略時は deepen_refs）",
    )
    ap.add_argument(
        "--max-shots",
        type=int,
        default=0,
        help="この枚数撮ったら終了（0=手動でq）",
    )
    ap.add_argument(
        "--focus",
        type=float,
        default=30.0,
        help="マニュアル時の初期フォーカス値（小さいほど近い・機種依存。既定30）",
    )
    args = ap.parse_args()
    jan = "".join(ch for ch in str(args.jan) if ch.isdigit())
    if not jan:
        print("JANが空です", file=sys.stderr)
        return 2

    global cv2
    try:
        import cv2
    except ImportError:
        print("OpenCV がありません: pip install opencv-python", file=sys.stderr)
        return 2

    out_dir = _out_dir(jan, custom=args.out_dir)
    cam_i = int(args.camera or 0)
    shots = 0
    autofocus = True
    focus_val = float(args.focus)
    shoot_flag = {"go": False}
    win = "SURFACE webcam"

    print("JAN", jan)
    print("保存先", out_dir)
    print("操作: 撮影ボタン or Space  /  [ ]=ピント  a=AF  q=終了  c=カメラ")

    cap = None

    def open_cam(idx: int) -> bool:
        nonlocal cap
        if cap is not None:
            try:
                cap.release()
            except Exception:
                pass
            cap = None
        # DirectShow 優先（Windowsでフォーカス操作しやすい）
        cap = cv2.VideoCapture(idx, cv2.CAP_DSHOW)
        if not cap.isOpened():
            cap = cv2.VideoCapture(idx)
        if not (cap is not None and cap.isOpened()):
            return False
        try:
            cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
            cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)
        except Exception:
            pass
        _apply_near_focus(cap, autofocus=autofocus, focus_val=focus_val)
        # 起動直後に近接AFを促す（対応機のみ）
        if autofocus:
            try:
                cap.set(cv2.CAP_PROP_AUTOFOCUS, 1)
            except Exception:
                pass
        return True

    def save_shot(frame) -> None:
        nonlocal shots
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        name = f"ref_surface_cam_{stamp}.jpg"
        dest = out_dir / name
        h, w = frame.shape[:2]
        x0, y0, x1, y1 = _guide_xy(w, h)
        crop = frame[y0:y1, x0:x1]
        if crop.size == 0:
            crop = frame
        cv2.imwrite(str(dest), crop, [int(cv2.IMWRITE_JPEG_QUALITY), 95])
        shots += 1
        print("保存+", dest)
        full = out_dir / f"ref_surface_cam_{stamp}_full.jpg"
        cv2.imwrite(str(full), frame, [int(cv2.IMWRITE_JPEG_QUALITY), 92])

    def on_mouse(event, x, y, flags, param) -> None:
        if event != cv2.EVENT_LBUTTONDOWN:
            return
        # param = {"w","h"} 最新フレームサイズ
        if not isinstance(param, dict):
            return
        vw = int(param.get("w") or 0)
        vh = int(param.get("h") or 0)
        if vw <= 0 or vh <= 0:
            return
        bx0, by0, bx1, by1 = _btn_xy(vw, vh)
        if bx0 <= x <= bx1 and by0 <= y <= by1:
            shoot_flag["go"] = True

    if not open_cam(cam_i):
        print(f"カメラ {cam_i} を開けませんでした", file=sys.stderr)
        return 1

    cv2.namedWindow(win, cv2.WINDOW_NORMAL)
    # マウス座標用に表示サイズを共有
    view_size = {"w": 0, "h": 0}
    cv2.setMouseCallback(win, on_mouse, view_size)

    last_ok = False
    last_frame = None
    try:
        while True:
            # ウィンドウが閉じられたら終了
            try:
                if cv2.getWindowProperty(win, cv2.WND_PROP_VISIBLE) < 1:
                    break
            except Exception:
                break

            ok, frame = cap.read()
            last_ok = bool(ok and frame is not None)
            if last_ok:
                last_frame = frame

            if not last_ok or last_frame is None:
                print("フレーム取得失敗。c でカメラ切替を試してください")
                key = cv2.waitKey(200) & 0xFF
            else:
                vis = last_frame.copy()
                h, w = vis.shape[:2]
                view_size["w"] = w
                view_size["h"] = h
                x0, y0, x1, y1 = _guide_xy(w, h)
                cv2.rectangle(vis, (x0, y0), (x1, y1), (0, 200, 255), 2)
                cv2.putText(
                    vis,
                    "put subject inside orange box",
                    (x0, max(24, y0 - 8)),
                    cv2.FONT_HERSHEY_SIMPLEX,
                    0.65,
                    (0, 180, 255),
                    2,
                    cv2.LINE_AA,
                )
                mode = "AF" if autofocus else f"MF:{int(focus_val)}"
                cv2.putText(
                    vis,
                    f"cam={cam_i} shots={shots} focus={mode}  [/]=near/far  a=AF",
                    (12, 28),
                    cv2.FONT_HERSHEY_SIMPLEX,
                    0.6,
                    (30, 30, 30),
                    2,
                    cv2.LINE_AA,
                )
                # 撮影ボタン
                bx0, by0, bx1, by1 = _btn_xy(w, h)
                cv2.rectangle(vis, (bx0, by0), (bx1, by1), (40, 160, 40), -1)
                cv2.rectangle(vis, (bx0, by0), (bx1, by1), (20, 90, 20), 2)
                label = "SHOT"
                ts = cv2.getTextSize(label, cv2.FONT_HERSHEY_SIMPLEX, 1.0, 2)[0]
                cv2.putText(
                    vis,
                    label,
                    (bx0 + (bx1 - bx0 - ts[0]) // 2, by0 + (BTN_H + ts[1]) // 2 - 2),
                    cv2.FONT_HERSHEY_SIMPLEX,
                    1.0,
                    (255, 255, 255),
                    2,
                    cv2.LINE_AA,
                )
                cv2.imshow(win, vis)
                key = cv2.waitKey(16) & 0xFF

            do_shoot = shoot_flag["go"] or key in (ord(" "), ord("s"))
            if shoot_flag["go"]:
                shoot_flag["go"] = False

            if key in (ord("q"), 27):
                break
            if key in (ord("c"),):
                cam_i = (cam_i + 1) % 4
                print("カメラ切替 →", cam_i)
                if not open_cam(cam_i):
                    print(f"カメラ {cam_i} 失敗（前の番号を再試行）")
                    cam_i = int(args.camera or 0)
                    open_cam(cam_i)
                continue
            if key in (ord("a"),):
                autofocus = not autofocus
                print("オートフォーカス", "ON" if autofocus else "OFF", "focus=", focus_val)
                _apply_near_focus(cap, autofocus=autofocus, focus_val=focus_val)
                continue
            if key in (ord("["), ord("-")):
                # 近くへ
                autofocus = False
                focus_val = max(0.0, focus_val - 5.0)
                print("ピント近く →", focus_val)
                _apply_near_focus(cap, autofocus=False, focus_val=focus_val)
                continue
            if key in (ord("]"), ord("="), ord("+")):
                autofocus = False
                focus_val = min(255.0, focus_val + 5.0)
                print("ピント遠く →", focus_val)
                _apply_near_focus(cap, autofocus=False, focus_val=focus_val)
                continue
            if do_shoot:
                if not last_ok or last_frame is None:
                    print("保存スキップ（フレーム無し）")
                    continue
                save_shot(last_frame)
                if int(args.max_shots or 0) > 0 and shots >= int(args.max_shots):
                    print("max-shots 到達")
                    break
    finally:
        if cap is not None:
            try:
                cap.release()
            except Exception:
                pass
        try:
            cv2.destroyAllWindows()
        except Exception:
            pass

    print("完了 shots=", shots)
    print("SURFACE解決は deepen_refs/ref_surface* を優先します。")
    print("次: SUB再生成でこの面が載るか確認。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
