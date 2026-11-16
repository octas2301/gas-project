# -*- coding: utf-8 -*-
"""判定単位 judgeSeed の候補出し＋人間ラジオ選択（正本 PLATING §3.3）。

機械が種から候補クロップを出し、人間がラジオでちょうど1つ選ぶ。
選んだものが以降の合否・敢えて順位の種になる。

切戻し: 呼び出し側で使わなければ従来どおり種全体。
"""
from __future__ import annotations

import base64
import json
import os
import re
import socket
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple
from urllib.parse import urlparse

from PIL import Image

DEFAULT_PORT = 8850
MAX_CANDIDATES = 8
MIN_CANDIDATES = 1
MODEL = "gpt-4o-mini"


def _esc(s: object) -> str:
    import html as htmlmod

    return htmlmod.escape(str(s or ""), quote=True)


def _b64_data_url(path: Path) -> str:
    raw = path.read_bytes()
    mime = "image/png" if path.suffix.lower() == ".png" else "image/jpeg"
    return f"data:{mime};base64,{base64.b64encode(raw).decode('ascii')}"


def _parse_json(text: str) -> dict:
    t = re.sub(r"```\w*\n?", "", text or "").strip()
    s, e = t.find("{"), t.rfind("}")
    if s < 0 or e <= s:
        raise RuntimeError(f"JSONなし: {(text or '')[:240]}")
    return json.loads(t[s : e + 1])


def _clamp01(v: float) -> float:
    return max(0.0, min(1.0, float(v)))


def _port_free(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            s.bind(("127.0.0.1", port))
            return True
        except OSError:
            return False


def find_free_port(base: int = DEFAULT_PORT, span: int = 20) -> int:
    for p in range(base, base + span):
        if _port_free(p):
            return p
    raise RuntimeError(f"空きポートなし base={base}")


def propose_unit_boxes(seed: Path) -> List[Dict[str, Any]]:
    """Visionで種内の1単位候補の正規化框を返す。商品固有語は焼かない。"""
    from openai import OpenAI
    from openai_image import load_openai_api_key

    prompt = (
        "この写真の中で「1単位の見本」として使えそうな対象を列挙せよ。"
        "盛り全体・山・ばら撒きは1候補にまとめるな。重なりが少なく縁が見える単位を優先。"
        "最大8件。座標は画像に対する正規化 0〜1（x0,y0=左上, x1,y1=右下）。"
        "単位が1つだけなら candidates は1件で全体でもよい。"
        "JSONのみ:\n"
        '{"unitCountHint":1,"candidates":[{"id":"u1","x0":0.1,"y0":0.1,"x1":0.5,"y1":0.6,"why":"一文"}]}'
    )
    client = OpenAI(api_key=load_openai_api_key())
    content = [
        {"type": "text", "text": prompt},
        {
            "type": "image_url",
            "image_url": {"url": _b64_data_url(seed), "detail": "low"},
        },
    ]
    resp = client.chat.completions.create(
        model=MODEL,
        messages=[{"role": "user", "content": content}],
        max_tokens=900,
        temperature=0,
    )
    obj = _parse_json(resp.choices[0].message.content or "")
    rows = obj.get("candidates") or []
    out: List[Dict[str, Any]] = []
    for i, r in enumerate(rows):
        if not isinstance(r, dict):
            continue
        try:
            box = {
                "id": str(r.get("id") or f"u{i+1}"),
                "x0": _clamp01(r["x0"]),
                "y0": _clamp01(r["y0"]),
                "x1": _clamp01(r["x1"]),
                "y1": _clamp01(r["y1"]),
                "why": str(r.get("why") or "").strip(),
            }
        except Exception:
            continue
        if box["x1"] <= box["x0"] or box["y1"] <= box["y0"]:
            continue
        # 面積が極端に小さいものは落とす
        area = (box["x1"] - box["x0"]) * (box["y1"] - box["y0"])
        if area < 0.01:
            continue
        out.append(box)
        if len(out) >= MAX_CANDIDATES:
            break
    if not out:
        out = [
            {
                "id": "whole",
                "x0": 0.0,
                "y0": 0.0,
                "x1": 1.0,
                "y1": 1.0,
                "why": "Vision候補なしのため全体を候補1",
            }
        ]
    return out[:MAX_CANDIDATES]


def _pad_box(
    x0: float, y0: float, x1: float, y1: float, *, pad: float = 0.04
) -> Tuple[float, float, float, float]:
    return (
        _clamp01(x0 - pad),
        _clamp01(y0 - pad),
        _clamp01(x1 + pad),
        _clamp01(y1 + pad),
    )


def crop_candidates(
    seed: Path, boxes: Sequence[Dict[str, Any]], out_dir: Path
) -> List[Dict[str, Any]]:
    """框から候補画像を切り、メタ付きで返す。全体候補も必ず末尾近くに残す。"""
    out_dir.mkdir(parents=True, exist_ok=True)
    im = Image.open(seed).convert("RGB")
    w, h = im.size
    seed_copy = out_dir / "seed_full.jpg"
    if seed.resolve() != seed_copy.resolve():
        im.save(seed_copy, quality=95)

    cands: List[Dict[str, Any]] = []
    for i, b in enumerate(boxes):
        x0, y0, x1, y1 = _pad_box(b["x0"], b["y0"], b["x1"], b["y1"])
        box_px = (
            int(x0 * w),
            int(y0 * h),
            max(int(x0 * w) + 1, int(x1 * w)),
            max(int(y0 * h) + 1, int(y1 * h)),
        )
        cid = re.sub(r"[^\w\-]+", "_", str(b.get("id") or f"u{i+1}"))[:40]
        name = f"cand_{i+1:02d}_{cid}.jpg"
        dest = out_dir / name
        im.crop(box_px).save(dest, quality=95)
        cands.append(
            {
                "id": cid,
                "index": i + 1,
                "file": name,
                "path": str(dest),
                "why": str(b.get("why") or ""),
                "boxNorm": {
                    "x0": b["x0"],
                    "y0": b["y0"],
                    "x1": b["x1"],
                    "y1": b["y1"],
                },
                "boxPx": list(box_px),
            }
        )

    # 全体を候補に含める（まだ無ければ）
    if not any(
        abs(c["boxNorm"]["x0"]) < 0.02
        and abs(c["boxNorm"]["y0"]) < 0.02
        and c["boxNorm"]["x1"] > 0.98
        and c["boxNorm"]["y1"] > 0.98
        for c in cands
    ):
        whole = out_dir / f"cand_{len(cands)+1:02d}_whole.jpg"
        im.save(whole, quality=95)
        cands.append(
            {
                "id": "whole",
                "index": len(cands) + 1,
                "file": whole.name,
                "path": str(whole),
                "why": "種の全体（切り出さない）",
                "boxNorm": {"x0": 0.0, "y0": 0.0, "x1": 1.0, "y1": 1.0},
                "boxPx": [0, 0, w, h],
            }
        )
    return cands[:MAX_CANDIDATES]


def build_candidates(seed: Path, work_dir: Path) -> Dict[str, Any]:
    work_dir.mkdir(parents=True, exist_ok=True)
    boxes = propose_unit_boxes(seed)
    cands = crop_candidates(seed, boxes, work_dir / "candidates")
    meta = {
        "seed": str(seed),
        "seedFull": str(work_dir / "candidates" / "seed_full.jpg"),
        "candidates": cands,
        "status": "awaiting_human",
        "selectedId": None,
        "judgeSeed": None,
        "note": "正本§3.3 機械候補→人間ラジオ1つ",
    }
    (work_dir / "candidates.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return meta


def _page_html(meta: Dict[str, Any]) -> str:
    seed_rel = "candidates/seed_full.jpg"
    cards = []
    for c in meta.get("candidates") or []:
        cid = _esc(c.get("id"))
        f = _esc(c.get("file"))
        why = _esc(c.get("why"))
        idx = _esc(c.get("index"))
        cards.append(
            f"""
<label class="card">
  <div class="top">
    <input type="radio" name="pick" value="{cid}" data-file="{f}"/>
    <span>候補 {idx} · {cid}</span>
  </div>
  <img src="candidates/{f}" alt="{cid}" loading="lazy"/>
  <p class="why">{why}</p>
</label>"""
        )
    return f"""<!doctype html>
<html lang="ja"><meta charset="utf-8"/>
<title>判定種の選択（judgeSeed）</title>
<style>
body{{margin:0;font-family:sans-serif;background:#121212;color:#eee}}
header{{padding:18px 22px;background:#1c1c1c;border-bottom:1px solid #333}}
h1{{margin:0 0 6px;font-size:20px}}
.note{{margin:0;font-size:13px;line-height:1.5;color:#bbb}}
main{{padding:18px 22px 64px;max-width:1100px;margin:0 auto}}
.seed{{display:grid;grid-template-columns:minmax(220px,340px) 1fr;gap:16px;
  background:#1a1a1a;border:1px solid #333;border-radius:12px;padding:12px;margin-bottom:18px}}
.seed img{{width:100%;border-radius:8px;cursor:zoom-in;border:1px solid #333}}
.grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:14px}}
.card{{display:block;background:#181818;border:1px solid #333;border-radius:12px;padding:10px;cursor:pointer}}
.card:has(input:checked){{border-color:#6af;box-shadow:0 0 0 1px #6af}}
.card .top{{display:flex;gap:8px;align-items:center;font-size:13px;margin-bottom:8px}}
.card img{{width:100%;height:auto;display:block;border-radius:8px;cursor:zoom-in}}
.why{{font-size:12px;color:#bcd;line-height:1.4;margin:8px 0 0}}
.bar{{position:fixed;left:0;right:0;bottom:0;background:#1c1c1c;border-top:1px solid #333;
  padding:12px 18px;display:flex;gap:12px;align-items:center;justify-content:center}}
button{{background:#2a5;color:#041;border:0;border-radius:8px;padding:10px 18px;font-size:15px;cursor:pointer}}
button:disabled{{opacity:.4;cursor:not-allowed}}
#msg{{font-size:13px;color:#9c9}}
#lightbox{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.92);z-index:99;
  align-items:center;justify-content:center;cursor:zoom-out}}
#lightbox.on{{display:flex}}
#lightbox img{{max-width:96vw;max-height:96vh;object-fit:contain}}
</style>
<header>
  <h1>判定種（judgeSeed）を1つ選ぶ</h1>
  <p class="note">正本 §3.3｜機械が出した候補からラジオでちょうど1つ。型・文案は書かない。画像クリックで拡大。</p>
</header>
<main>
  <section class="seed">
    <div>
      <img src="{seed_rel}" alt="seed full"/>
      <p style="font-size:12px;color:#888;margin:6px 0 0">種（全体）</p>
    </div>
    <div>
      <h2 style="margin:0 0 8px;font-size:15px;color:#9cf">手順</h2>
      <ol style="margin:0;padding-left:1.2em;font-size:14px;line-height:1.55;color:#ccc">
        <li>候補を見て、合否の正にしたい1単位を選ぶ</li>
        <li>「この候補で進む」を押す</li>
        <li>選んだ画像が以降の一片合否の種になる</li>
      </ol>
    </div>
  </section>
  <div class="grid">
    {''.join(cards)}
  </div>
</main>
<div class="bar">
  <button id="go" type="button" disabled>この候補で進む</button>
  <span id="msg">候補を1つ選んでください</span>
</div>
<div id="lightbox"><img id="lbImg" alt="拡大"/></div>
<script>
const go=document.getElementById('go');
const msg=document.getElementById('msg');
const lb=document.getElementById('lightbox');
const lbImg=document.getElementById('lbImg');
function selected(){{
  return document.querySelector('input[name=pick]:checked');
}}
document.querySelectorAll('input[name=pick]').forEach(r=>{{
  r.addEventListener('change',()=>{{
    go.disabled=!selected();
    msg.textContent=selected()?('選択中: '+selected().value):'候補を1つ選んでください';
  }});
}});
document.querySelectorAll('img').forEach(img=>{{
  img.addEventListener('click',e=>{{
    e.preventDefault(); e.stopPropagation();
    lbImg.src=img.src; lb.classList.add('on');
  }});
}});
lb.addEventListener('click',()=>lb.classList.remove('on'));
document.addEventListener('keydown',e=>{{if(e.key==='Escape')lb.classList.remove('on');}});
go.addEventListener('click',async()=>{{
  const r=selected();
  if(!r) return;
  go.disabled=true;
  msg.textContent='保存中…';
  const res=await fetch('/api/select',{{
    method:'POST',
    headers:{{'Content-Type':'application/json'}},
    body:JSON.stringify({{id:r.value, file:r.dataset.file}})
  }});
  const j=await res.json().catch(()=>({{}}));
  if(!res.ok){{
    msg.textContent='失敗: '+(j.error||res.status);
    go.disabled=false;
    return;
  }}
  msg.textContent='保存しました。このタブは閉じてよいです。 judgeSeed='+(j.judgeSeed||'');
}});
</script>
</html>
"""


def _write_index(work_dir: Path, meta: Dict[str, Any]) -> Path:
    html = _page_html(meta)
    path = work_dir / "index.html"
    path.write_text(html, encoding="utf-8")
    return path


def apply_selection(work_dir: Path, selected_id: str, selected_file: str) -> Dict[str, Any]:
    meta_path = work_dir / "candidates.json"
    meta = json.loads(meta_path.read_text(encoding="utf-8"))
    cands = meta.get("candidates") or []
    hit = None
    for c in cands:
        if str(c.get("id")) == selected_id or str(c.get("file")) == selected_file:
            hit = c
            break
    if hit is None:
        raise ValueError(f"候補なし id={selected_id} file={selected_file}")
    src = Path(hit["path"])
    if not src.is_file():
        src = work_dir / "candidates" / hit["file"]
    dest = work_dir / "judge_seed.jpg"
    Image.open(src).convert("RGB").save(dest, quality=95)
    meta["status"] = "selected"
    meta["selectedId"] = hit["id"]
    meta["selectedFile"] = hit["file"]
    meta["judgeSeed"] = str(dest)
    meta["rejectedIds"] = [c["id"] for c in cands if c.get("id") != hit["id"]]
    meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    (work_dir / "selection_done.json").write_text(
        json.dumps(
            {
                "ok": True,
                "selectedId": hit["id"],
                "judgeSeed": str(dest),
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    return meta


def serve_until_selected(
    work_dir: Path,
    *,
    port: int,
    open_browser: bool = True,
    timeout_sec: Optional[float] = None,
) -> Path:
    """HTTPでラジオ選択を待ち、judge_seed.jpg の Path を返す。"""
    meta = json.loads((work_dir / "candidates.json").read_text(encoding="utf-8"))
    _write_index(work_dir, meta)
    done_file = work_dir / "selection_done.json"
    if done_file.is_file():
        done_file.unlink()

    state = {"error": None}

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, fmt: str, *args: Any) -> None:
            return

        def _send(self, code: int, body: bytes, ctype: str) -> None:
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def do_GET(self) -> None:  # noqa: N802
            parsed = urlparse(self.path)
            rel = parsed.path.lstrip("/") or "index.html"
            if ".." in rel:
                self._send(400, b"bad path", "text/plain")
                return
            path = work_dir / rel
            if not path.is_file():
                self._send(404, b"not found", "text/plain")
                return
            ctype = "text/html; charset=utf-8"
            if path.suffix.lower() in (".jpg", ".jpeg"):
                ctype = "image/jpeg"
            elif path.suffix.lower() == ".png":
                ctype = "image/png"
            elif path.suffix.lower() == ".json":
                ctype = "application/json; charset=utf-8"
            self._send(200, path.read_bytes(), ctype)

        def do_POST(self) -> None:  # noqa: N802
            parsed = urlparse(self.path)
            if parsed.path != "/api/select":
                self._send(404, b'{"error":"not found"}', "application/json")
                return
            n = int(self.headers.get("Content-Length") or 0)
            raw = self.rfile.read(n) if n else b"{}"
            try:
                payload = json.loads(raw.decode("utf-8"))
                m = apply_selection(
                    work_dir,
                    str(payload.get("id") or ""),
                    str(payload.get("file") or ""),
                )
                body = json.dumps(
                    {"ok": True, "judgeSeed": m.get("judgeSeed"), "selectedId": m.get("selectedId")},
                    ensure_ascii=False,
                ).encode("utf-8")
                self._send(200, body, "application/json; charset=utf-8")
            except Exception as e:
                state["error"] = str(e)
                body = json.dumps({"error": str(e)}, ensure_ascii=False).encode("utf-8")
                self._send(400, body, "application/json; charset=utf-8")

    httpd = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    thread = threading.Thread(target=httpd.serve_forever, daemon=True)
    thread.start()
    url = f"http://127.0.0.1:{port}/"
    print(f"JUDGE_SEED_PICK_URL {url}")
    print(f"JUDGE_SEED_PICK_DIR {work_dir}")
    if open_browser:
        try:
            webbrowser.open(url)
        except Exception:
            pass

    t0 = time.time()
    try:
        while True:
            if done_file.is_file():
                data = json.loads(done_file.read_text(encoding="utf-8"))
                jp = Path(str(data.get("judgeSeed") or ""))
                if jp.is_file():
                    print(f"JUDGE_SEED_SELECTED {jp}")
                    return jp
            if state["error"]:
                raise RuntimeError(state["error"])
            if timeout_sec is not None and (time.time() - t0) > timeout_sec:
                raise TimeoutError(f"judgeSeed選択タイムアウト {timeout_sec}s")
            time.sleep(0.4)
    finally:
        httpd.shutdown()


def resolve_judge_seed(
    seed: Path,
    work_dir: Path,
    *,
    port: Optional[int] = None,
    open_browser: bool = True,
    auto_if_single: Optional[bool] = None,
    timeout_sec: Optional[float] = None,
) -> Path:
    """種から候補を出し、必要なら人間選択を待って judge_seed.jpg を返す。"""
    seed = Path(seed)
    work_dir = Path(work_dir)
    if not seed.is_file():
        raise FileNotFoundError(seed)

    if auto_if_single is None:
        raw = str(os.environ.get("B5_JUDGE_SEED_SKIP_IF_SINGLE") or "0").strip().lower()
        auto_if_single = raw in ("1", "true", "on", "yes")

    meta = build_candidates(seed, work_dir)
    cands = meta.get("candidates") or []
    # 全体以外が実質0で whole だけのときも「単一」扱い
    non_whole = [c for c in cands if c.get("id") != "whole"]
    single = len(cands) == 1 or (len(non_whole) <= 1 and len(cands) <= 2 and len(non_whole) == 1)

    if auto_if_single and single:
        hit = non_whole[0] if non_whole else cands[0]
        apply_selection(work_dir, str(hit["id"]), str(hit["file"]))
        jp = work_dir / "judge_seed.jpg"
        print(f"JUDGE_SEED_AUTO_SINGLE {jp}")
        _write_index(work_dir, json.loads((work_dir / "candidates.json").read_text(encoding="utf-8")))
        return jp

    p = int(port or find_free_port(DEFAULT_PORT))
    return serve_until_selected(
        work_dir,
        port=p,
        open_browser=open_browser,
        timeout_sec=timeout_sec,
    )


__all__ = [
    "DEFAULT_PORT",
    "build_candidates",
    "resolve_judge_seed",
    "serve_until_selected",
    "find_free_port",
]
