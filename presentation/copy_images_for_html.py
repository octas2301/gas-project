# -*- coding: utf-8 -*-
"""
html_slides 用に、提供いただいた4枚の南の海画像を images/ にコピーする。
beach1.png ～ beach4.png として保存し、index.html から参照できるようにする。
"""
import os
import glob
import shutil

BASE = os.path.dirname(os.path.abspath(__file__))
IMAGES_SRC = os.path.join(BASE, "images")
IMAGES_DST = os.path.join(BASE, "html_slides", "images")
UUIDS = ("c72406ed", "056ad239", "e3f6f5b4", "e5f74cef")

def main():
    os.makedirs(IMAGES_DST, exist_ok=True)
    for i, uid in enumerate(UUIDS, start=1):
        pattern = os.path.join(IMAGES_SRC, "*" + uid + "*.png")
        found = glob.glob(pattern)
        if found:
            dst = os.path.join(IMAGES_DST, "beach%d.png" % i)
            shutil.copy2(found[0], dst)
            print("OK: %s -> %s" % (os.path.basename(found[0]), dst))
        else:
            print("SKIP: %s に該当画像なし (beach%d.png)" % (IMAGES_SRC, i))
    print("完了: html_slides/images/ を確認してください。")

if __name__ == "__main__":
    main()
