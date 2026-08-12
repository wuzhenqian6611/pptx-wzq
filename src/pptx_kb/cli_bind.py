"""cli_bind.py — pptx-bind 薄壳：把教材文案与对应页图片关系绑定为 JSON。

输入：textbook.md（每页文案，## 第 N 页）、images 目录（文件名 slide_NN_
     编码页码）、captions.md（图片 AI 解读，### IMGxxxx — slide_NN_）。
输出：<名>_binding.json，按页组织：

    {
      "stem": "xxx",
      "pages": [
        {"page": 4,
         "text": "第 4 页文案全文…",
         "images": [{"file": "slide_04_pic_05.png",
                     "caption": "1) 图片类型…" }],
         "has_image": true}
      ]
    }

用法：
    pptx-bind <产物目录> [-o binding.json]
              [--textbook a.md] [--images-dir dir] [--captions c.md]
              [--json] [--version]

退出码：0 成功 / 1 处理异常 / 2 参数或环境错误。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

from pptx_kb.cli_common import EXIT_ERR, EXIT_OK, EXIT_USAGE, print_json, quiet_stdout

VERSION = "pptx-bind 1.0.0 (方案B薄壳)"


def _split_pages(content: str) -> dict:
    pages = {}
    cur = None
    for line in content.splitlines():
        m = re.match(r"^##\s*第\s*(\d+)\s*页\s*$", line.strip())
        if m:
            cur = int(m.group(1))
            pages.setdefault(cur, [])
        elif cur is not None:
            pages[cur].append(line)
    return pages


def parse_textbook(path: Path) -> dict:
    """textbook.md → {page: 文案全文}。"""
    out = {}
    for page, lines in _split_pages(path.read_text(encoding="utf-8")).items():
        text = "\n".join(x for x in lines if x.strip()).strip()
        if text:
            out[page] = text
    return out


def _page_of(file_name: str):
    """从文件名 slide_NN_... 提取页码；失败返回 None。"""
    m = re.match(r"slide_(\d+)_", file_name)
    return int(m.group(1)) if m else None


def index_images(images_dir: Path) -> dict:
    """images 目录 → {page: [文件名]}（按 slide_NN 分页）。"""
    out = {}
    if not images_dir.is_dir():
        return out
    for p in sorted(images_dir.iterdir()):
        pg = _page_of(p.name)
        if pg is not None:
            out.setdefault(pg, []).append(p.name)
    return out


def parse_captions(path: Path) -> dict:
    """captions.md → {file_name: 解读文本}。"""
    out = {}
    cur_file = None
    buf = []
    for ln in path.read_text(encoding="utf-8").splitlines():
        m = re.match(r"^###\s+IMG\d+\s*—\s*`?([^`\s]+\.\w+)\s*`?", ln)
        if m:
            if cur_file is not None and buf:
                out[cur_file] = " ".join(x.strip() for x in buf if x.strip())
            cur_file = m.group(1)
            buf = []
        elif cur_file is not None:
            buf.append(ln)
    if cur_file is not None and buf:
        out[cur_file] = " ".join(x.strip() for x in buf if x.strip())
    return out


def build_binding(stem: str, textbook: dict, img_by_page: dict,
                  cap_by_file: dict) -> dict:
    """按页绑定文案与图片。"""
    pages = sorted(set(textbook) | set(img_by_page))
    out_pages = []
    for p in pages:
        imgs = []
        for fname in img_by_page.get(p, []):
            imgs.append({"file": fname,
                         "caption": cap_by_file.get(fname, "")})
        out_pages.append({"page": p,
                          "text": textbook.get(p, ""),
                          "images": imgs,
                          "has_image": bool(imgs)})
    n_img = sum(len(p["images"]) for p in out_pages)
    n_bound = sum(1 for p in out_pages if p["has_image"])
    return {
        "stem": stem,
        "pages": out_pages,
        "summary": {"pages": len(out_pages),
                    "images_total": n_img,
                    "pages_with_image": n_bound},
    }


def _locate(out_dir: Path, args):
    """定位 textbook / images / captions（参数优先，目录自动找兜底）。"""
    tb = args.textbook
    if tb is None:
        cands = sorted(out_dir.glob("*_textbook.md"))
        tb = cands[0] if cands else None
    img_dir = args.images_dir or (out_dir / "images")
    cap = args.captions
    if cap is None:
        cands = sorted(out_dir.glob("*_captions.md")) + \
            sorted(out_dir.glob("images_captions.md"))
        cap = cands[0] if cands else None
    return tb, img_dir, cap


def _main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        prog="pptx-bind",
        description="把教材文案与对应页图片关系绑定为 JSON")
    ap.add_argument("dir", nargs="?", default=".",
                    help="产物目录（自动找 textbook/images/captions）")
    ap.add_argument("-o", "--output", default=None,
                    help="输出 json 路径（默认 <目录>/<名>_binding.json）")
    ap.add_argument("--textbook", type=Path, default=None, help="教材文案 md")
    ap.add_argument("--images-dir", type=Path, default=None, help="images 目录")
    ap.add_argument("--captions", type=Path, default=None, help="图片解读 md")
    ap.add_argument("--json", action="store_true",
                    help="把统计输出到 stdout")
    ap.add_argument("--version", action="version", version=VERSION)
    args = ap.parse_args(argv)

    out_dir = Path(args.dir)
    tb, img_dir, cap = _locate(out_dir, args)
    missing = [n for n, p in (("textbook", tb), ("captions", cap))
               if p is None or not p.exists()]
    if missing:
        print(f"[错误] 缺少输入：{', '.join(missing)}"
              f"（请传目录或 --textbook/--captions）", file=sys.stderr)
        return EXIT_USAGE
    if not img_dir.is_dir():
        print(f"[错误] images 目录不存在：{img_dir}", file=sys.stderr)
        return EXIT_USAGE

    try:
        stem = tb.stem[: -len("_textbook")] if tb.stem.endswith("_textbook") \
            else tb.stem
        binding = build_binding(stem, parse_textbook(tb),
                                index_images(img_dir), parse_captions(cap))
        out_path = Path(args.output) if args.output else \
            out_dir / f"{stem}_binding.json"
        out_path.write_text(
            json.dumps(binding, ensure_ascii=False, indent=1),
            encoding="utf-8")
        s = binding["summary"]
        print(f"[OK] 图文绑定已写出：{out_path}")
        print(f"     {s['pages']} 页 / 图片 {s['images_total']} 张 / "
              f"含图页 {s['pages_with_image']}")
        if args.json:
            print_json({"pages": s["pages"], "images": s["images_total"],
                        "bound_pages": s["pages_with_image"],
                        "output": str(out_path)})
        return EXIT_OK
    except Exception as e:
        print(f"[错误] 图文绑定失败：{e}", file=sys.stderr)
        return EXIT_ERR


def main() -> int:  # console
    return _main()


if __name__ == "__main__":
    sys.exit(_main())
