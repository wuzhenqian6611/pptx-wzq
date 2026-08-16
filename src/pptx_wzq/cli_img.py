#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
cli_img.py — pptx-img：PPTX 图片解析 CLI（方案 B 薄壳 + v5 过滤）

用法
----
  pptx-img extract <input.pptx> [-o OUT] [选项]
  pptx-img extract deck.pptx -o out --no-fill --no-crop --json

图片全部提取完成后，按教学判据做取舍后处理（img_filter）：
  规则层：尺寸过小 / 特别细长 / 纯色近纯色 → 移到 out/discarded/；
  YOLO 层：规则命中但检测到明确物体 → 强保留（防误删照片）。
被弃明细写 out/filter_report.json。--no-filter 关闭过滤。

启动时先做环境检查（Pillow / YOLO 引擎与本地权重 / LibreOffice）。
退出码：0 成功 / 1 处理异常 / 2 参数或文件错误。

作者：吴振谦 · wuzhenqian@nbu.edu.cn · QQ：38328063"""
from __future__ import annotations

import argparse
import base64
import os
import re
import shutil
import sys
from collections import Counter
from pathlib import Path

from pptx_wzq import extract_pptx_images as E
from pptx_wzq import img_filter
from pptx_wzq import pptrender
from pptx_wzq.cli_common import (EXIT_ERR, EXIT_OK, EXIT_USAGE,
                        make_progress, print_json, quiet_stdout,
                        resolve_input, resolve_output,
                        banner, banner_end)

VERSION = "pptx-img 2.1.0 (方案B薄壳+v5过滤+PPT渲染)"


def _img_progress(page_no: int, n_slides: int, info: dict) -> None:
    """图片提取进度回调：逐页打印每张图的 文件名/类型/大小/位置/形状。"""
    items = info.get("page_items", []) or []
    kinds = Counter(it["kind"] for it in items)
    kinds_s = "、".join(f"{k}×{v}" for k, v in sorted(kinds.items())) \
        if kinds else "无"
    print(f"[第 {page_no}/{n_slides} 页] 提取 {len(items)} 张（{kinds_s}）：",
          file=sys.stderr)
    for it in items:
        loc = f"({it['x']},{it['y']})" if (it["x"] or it["y"]) else "-"
        size = f"{it['w']}x{it['h']}px" if (it["w"] and it["h"]) else "-"
        print(f"    {it['file']:32s} {it['kind']:10s} {size:>12s}"
              f" 位置[{loc:>10s}] 形状:{it['shape'][:20]}",
              file=sys.stderr)


# --------------------------------------------------------------------------
# 矢量图 PowerPoint 渲染 + 图片公式识别（方案B）
# --------------------------------------------------------------------------
_OCR_SYSTEM = ("你是数学公式识别专家。识别图片中的数学公式，"
               "输出标准 LaTeX 代码（仅输出 LaTeX，不要任何其他文字）。")


def _ocr_formula_latex(png_path: Path, engine: str,
                       api_key_env: str) -> str | None:
    """把渲染好的公式 PNG 识别为 LaTeX。

    engine: qwen（视觉模型，失败回退 pix2tex）/ pix2tex（本地）/ skip。
    返回 LaTeX 字符串或 None。
    """
    if engine == "skip":
        return None
    if engine in ("qwen", "auto"):
        key = os.environ.get(api_key_env or "DASHSCOPE_API_KEY", "")
        if key:
            try:
                from openai import OpenAI
                client = OpenAI(
                    api_key=key,
                    base_url="https://dashscope.aliyuncs.com/"
                             "compatible-mode/v1")
                b64 = base64.b64encode(png_path.read_bytes()).decode("ascii")
                resp = client.chat.completions.create(
                    model="qwen3.7-plus",
                    messages=[{"role": "system", "content": _OCR_SYSTEM},
                              {"role": "user", "content": [
                                  {"type": "image_url",
                                   "image_url": {"url":
                                    f"data:image/png;base64,{b64}"}}]}],
                    stream=False, max_tokens=400)
                latex = (resp.choices[0].message.content or "").strip()
                if latex and latex not in ("", "```"):
                    return latex
            except Exception:
                pass
        if engine == "qwen":
            print("[公式识别] qwen 视觉不可用，回退本地 pix2tex",
                  file=sys.stderr)
    # 兜底：本地 pix2tex（需传 PIL Image，不接受路径字符串）
    try:
        from PIL import Image
        from pix2tex.cli import LatexOCR
        model = LatexOCR()
        with Image.open(str(png_path)) as im:
            return str(model(im)).strip() or None
    except Exception:
        return None


def _append_formula_to_md(fm_path: Path, page: int, latex: str) -> None:
    """把图片公式 LaTeX 追加进 formulas.md 的对应页小节。"""
    if not fm_path.parent.exists():
        fm_path.parent.mkdir(parents=True, exist_ok=True)
    lines = fm_path.read_text(encoding="utf-8").splitlines() \
        if fm_path.exists() else []
    if not lines:
        lines = ["# 图片公式补充（PowerPoint 渲染 + 视觉识别）", ""]

    # 页内序号：统计该页已有 (p{page}-f/img 条目
    seq = 1
    pat = re.compile(rf"\(p{page}-(?:f|img)\d+\)")
    for ln in lines:
        if pat.search(ln):
            seq += 1
    entry = (f"- **(p{page}-img{seq})** $$\n{latex.strip()}\n$$\n\n"
             "  > 来源：图片公式（PowerPoint 渲染+视觉识别）\n\n")

    # 定位 ## 第 N 页 节：有则插入节末，无则按页序新建
    page_marks = []
    for i, ln in enumerate(lines):
        m = re.match(r"^##\s*第\s*(\d+)\s*页\s*$", ln.strip())
        if m:
            page_marks.append((int(m.group(1)), i))
    if page_marks:
        pages_sorted = sorted(page_marks)
        insert_at = None
        for pg, idx in pages_sorted:
            if pg == page:
                # 插到本节末尾（下一个页节之前）
                nxt = next((j for p2, j in pages_sorted if p2 > pg), len(lines))
                insert_at = nxt
                break
            if pg > page:
                insert_at = idx
                break
        if insert_at is None:
            insert_at = len(lines)
        # 补一个空行分隔
        if insert_at < len(lines) and lines[insert_at - 1].strip():
            entry = "\n" + entry
        lines.insert(insert_at, entry.strip("\n"))
        lines.insert(insert_at, "")   # 前置空行
        out_text = "\n".join(lines)
    else:
        out_text = "\n".join(lines) + \
            f"\n## 第 {page} 页\n\n" + entry
    fm_path.write_text(out_text, encoding="utf-8")


def _vector_convert_one(src: Path, out_path: Path, tool) -> bool:
    """把矢量文件转为目标格式（soffice 或 inkscape）。返回是否成功。"""
    try:
        import subprocess as _sp
        if tool[0] == "soffice":
            r = _sp.run([tool[1], "--headless", "--convert-to",
                         out_path.suffix.lstrip("."),
                         "--outdir", str(out_path.parent), str(src)],
                        capture_output=True, timeout=180)
            if r.returncode != 0:
                return False
            produced = out_path.parent / (src.stem + out_path.suffix)
            if produced.is_file():
                if produced != out_path:
                    produced.replace(out_path)
                return True
            return False
        if tool[0] == "inkscape":
            r = _sp.run([tool[1], str(src), "--export-type=" +
                         out_path.suffix.lstrip("."),
                         "--export-filename=" + str(out_path)],
                        capture_output=True, timeout=180)
            return r.returncode == 0 and out_path.is_file()
    except Exception:
        pass
    return False


def _normalize_vectors(records, by_page: Path, args) -> dict:
    """需求2：保留的矢量图（emf/wmf/svg）规范化为 svg（--vector-out svg，默认）
    或 wmf（--vector-out wmf）。转换失败回退保留原格式文件。
    公式版 WMF（位图封装）不在此处理（走渲染+OCR 链路）；
    visio（vsdx/vsd）按需求1 直接存原文件，也不在此处理。
    返回 {"converted": n, "kept": n}。
    """
    want = getattr(args, "vector_out", "svg")
    tool = None
    try:
        tool = E._detect_rasterizer(getattr(args, "raster_prefer", "auto"))
    except Exception:
        tool = None
    if tool is None:
        print("[矢量] 未检测到 LibreOffice/Inkscape，矢量图保留原格式文件",
              file=sys.stderr)
        return {"converted": 0, "kept": 0}
    n_conv = 0
    for rec in list(records):
        if rec.original_format not in E.VECTOR or not rec.output_file:
            continue
        if rec.kind in ("visio", "formula_ole"):
            continue
        if rec.output_file.lower().endswith("." + want):
            continue
        src = by_page / rec.output_file
        if not src.is_file():
            continue
        dst = src.with_suffix("." + want)
        if _vector_convert_one(src, dst, tool) and dst.is_file():
            rec.output_file = dst.name
            rec.note = (rec.note + "；" if rec.note else "") + \
                f"由 {rec.original_format} 规范化为 {want}"
            try:
                src.unlink()
            except OSError:
                pass
            n_conv += 1
            print(f"[矢量] {src.name} → {dst.name}", file=sys.stderr)
    return {"converted": n_conv, "kept": len(records) - n_conv}


def _process_vectors(records, by_page: Path, out: Path, stem: str,
                     filtered: list, args) -> dict:
    """方案B后处理（过滤后、gallery 前调用）：

    1) 位图封装 WMF（公式版，已在过滤中弃用）→ PowerPoint 渲染 PNG →
       视觉识别 LaTeX → 追加进 formulas.md 对应页；
    2) 保留的真矢量图（emf/wmf/svg）→ 规范化为 svg/wmf（--vector-out，
       LibreOffice/Inkscape），失败保留原文件（不再默认栅格化 PNG）。
    返回 {"formulas": n, "rendered": n, "normalized": n}。
    """
    engine = getattr(args, "render_engine", "auto")
    if engine == "off":
        return {"formulas": 0, "rendered": 0, "normalized": 0}
    if engine == "auto" and not pptrender.check_available():
        print("[渲染] PowerPoint 不可用，跳过矢量渲染（可用 --render-engine off）",
              file=sys.stderr)
        return {"formulas": 0, "rendered": 0, "normalized": 0}

    # 1) 公式版 WMF（filtered 中 reason 含"位图封装"；源文件已被移到 discarded/）
    formula_jobs = []
    discarded_dir = out / "discarded"
    for f in filtered:
        if "位图封装" in f.get("reason", "") and f.get("file", "").endswith(".wmf"):
            src = discarded_dir / f["file"]
            if not src.exists():
                src = by_page / f["file"]
            if src.exists():
                png = by_page / (Path(f["file"]).stem + "_fm.png")
                formula_jobs.append((src, png, 4))

    n_rendered = pptrender.render_wmfs(formula_jobs, quiet=False)
    # 2) 公式识别并追加
    n_formulas = 0
    fm_path = Path(args.append_formulas) if args.append_formulas else \
        out / f"{stem}_formulas.md"
    for src, png, _ in formula_jobs:
        if not png.exists():
            continue
        page = int(Path(src).stem.split("_")[1])
        latex = _ocr_formula_latex(png, args.ocr_engine, args.api_key_env)
        if latex:
            _append_formula_to_md(fm_path, page, latex)
            n_formulas += 1
            print(f"[公式] 第 {page} 页 图片公式识别: {latex[:40]}…",
                  file=sys.stderr)
    # 3) 保留矢量图规范化（需求2：svg 优先，wmf 兜底）
    norm = _normalize_vectors(records, by_page, args)
    return {"formulas": n_formulas, "rendered": n_rendered,
            "normalized": norm["converted"]}


def _clean_output(out: Path) -> None:
    """清理输出目录中本命令生成的旧产物（防重跑残留）。

    只删本命令 100% 生成的内容：by_page/、images/、discarded/、
    all_media/ 四个子目录 + manifest.csv/json、filter_report.json、
    images.md 文件；不动 out 目录里其他任何文件。
    否则旧版本产物（如 *_bg_*.png 背景图）会残留并混入新结果。
    """
    if not out.exists():
        return
    for sub in ("by_page", "images", "discarded", "all_media"):
        d = out / sub
        if d.is_dir():
            shutil.rmtree(d, ignore_errors=True)
    for name in ("manifest.csv", "manifest.json", "filter_report.json",
                 "images.md"):
        f = out / name
        if f.is_file():
            try:
                f.unlink()
            except OSError:
                pass


def _stats(records: list, out_dir: str, filtered: list, yolo_used: bool):
    """基于 records 计算统计（与核心库 _write_manifest 的口径一致）。"""
    kinds = Counter(r.kind for r in records)
    n_slides = max((r.page for r in records), default=0)
    cropped = sum(1 for r in records if r.cropped == "yes")
    vector_skipped = sum(1 for r in records
                         if r.original_format in E.VECTOR and not r.converted_to_png)
    media_cnt = Counter(r.source_media for r in records
                        if r.source_media and r.source_media not in ("(missing)", ""))
    cross = sum(1 for m, c in media_cnt.items() if c > 1)
    reasons = Counter(f["reason"].split("(")[0] for f in filtered)
    return {
        "command": "pptx-img",
        "output": out_dir,
        "n_slides": n_slides,
        "n_records": len(records),
        "picture": kinds.get("picture", 0),
        "fill": kinds.get("fill", 0),
        "background": kinds.get("background", 0),
        "formula_ole": kinds.get("formula_ole", 0),
        "formula_omath": kinds.get("formula_omath", 0),
        "chart": kinds.get("chart", 0),
        "cropped": cropped,
        "vector_skipped": vector_skipped,
        "cross_page_reuse": cross,
        "filtered": len(filtered),
        "filtered_reasons": dict(reasons),
        "yolo_used": yolo_used,
        "manifest": f"{out_dir}/manifest.csv",
        "filter_report": f"{out_dir}/filter_report.json",
    }


def _check_env(args) -> None:
    """启动环境检查（stderr，不阻断）。Pillow/YOLO 由 img_filter.print_env 统一报告。"""
    img_filter.print_env("[环境]")
    if args.rasterize_vector:
        tool = E._detect_rasterizer(args.raster_prefer)
        if tool is None:
            print("[环境] 矢量栅格化: 未检测到 LibreOffice/Inkscape"
                  "（矢量图将保留原文件）", file=sys.stderr)
        else:
            print(f"[环境] 矢量栅格化: {tool[0]} 可用", file=sys.stderr)


def main(argv=None) -> int:
    banner("pptx-img")
    ap = argparse.ArgumentParser(
        prog="pptx-img",
        description="PPTX 图片解析：独立图片/形状填充/背景/公式OLE/图表 → 独立文件（标注页码）；"
                    "提取后按教学判据过滤（YOLO+规则）")
    ap.add_argument("pptx", help="输入的 .pptx 文件路径")
    ap.add_argument("-o", "--output", default=None,
                    help="输出目录（默认 <输入名>_images）")
    ap.add_argument("--no-convert", action="store_true",
                    help="不调用 Pillow 转换，直接按原格式落盘")
    ap.add_argument("--all-media", action="store_true",
                    help="额外导出 ppt/media 下所有媒体文件")
    ap.add_argument("--no-fill", action="store_true",
                    help="不提取形状填充图（<p:sp> 的 blipFill）")
    ap.add_argument("--no-bg-layout", action="store_true",
                    help="不追溯 slideLayout/slideMaster 背景，仅取幻灯片本体背景")
    ap.add_argument("--rasterize-vector", action="store_true",
                    help="用 LibreOffice/Inkscape 把 EMF/WMF/SVG 栅格化为 PNG")
    ap.add_argument("--raster-dpi", type=int, default=150,
                    help="矢量栅格化 DPI（Inkscape 生效；LibreOffice 忽略）")
    ap.add_argument("--raster-prefer", default="auto",
                    choices=["auto", "soffice", "inkscape"],
                    help="优先使用的栅格化引擎")
    ap.add_argument("--no-crop", action="store_true",
                    help="退回旧行为：不沿 srcRect 裁剪，导出完整媒体（用于溯源对比）")
    ap.add_argument("--min-crop", type=int, default=64,
                    help="srcRect 裁出区域小于该像素尺寸时退化保存完整图（默认 64）")
    ap.add_argument("--no-filter", action="store_true",
                    help="关闭取舍过滤（不移 discarded/，不调用 YOLO）")
    ap.add_argument("--min-size", type=int, default=48,
                    help="过滤：宽或高小于该像素视为尺寸过小（默认 48）")
    ap.add_argument("--max-ratio", type=float, default=10.0,
                    help="过滤：宽高比超过该值视为特别细长（默认 10）")
    ap.add_argument("--min-colors", type=int, default=4,
                    help="过滤：缩小采样后颜色种类不超过该值视为纯色图（默认 4）")
    ap.add_argument("--min-area", type=int, default=40000,
                    help="过滤：面积小于该像素值(px²)且前景占比过低的图视为"
                         "孤立字符/碎片图（默认 40000=200x200）")
    ap.add_argument("--max-sparse-ink", type=float, default=0.20,
                    help="过滤：面积 < --min-area 时前景像素占比低于该值"
                         "视为孤立字符/碎片图（默认 0.20=20%%）")
    ap.add_argument("--vec-min-area", type=int, default=10000,
                    help="过滤：矢量图(WMF/EMF/SVG)从文件头解析的内容尺寸"
                         "面积小于该像素值视为矢量碎片图（默认 10000=100x100）")
    ap.add_argument("--keep-background", action="store_true",
                    help="保留 PPT 背景图（母版/布局背景，默认过滤）")
    ap.add_argument("--keep-fill", action="store_true",
                    help="保留形状/文本框填充图（默认过滤）")
    ap.add_argument("--keep-vec-bitmap", action="store_true",
                    help="保留位图封装型 WMF（公式/符号渲染成位图嵌入，"
                         "默认视为公式矢量版过滤）")
    ap.add_argument("--render-engine", default="auto",
                    choices=["auto", "ppt", "off"],
                    help="矢量图渲染引擎：auto=探测 PowerPoint，ppt=强制 "
                         "PowerPoint，off=关闭渲染（默认 auto）")
    ap.add_argument("--vector-out", default="svg",
                    choices=["svg", "wmf"],
                    help="保留矢量图的规范化目标格式：svg=转 SVG（默认，"
                         "LibreOffice/Inkscape），wmf=转 WMF；转换失败回退"
                         "保留原格式文件")
    ap.add_argument("--ocr-engine", default="pix2tex",
                    choices=["qwen", "pix2tex", "skip"],
                    help="图片公式识别引擎：pix2tex=本地 latexocr（默认，零费用），"
                         "qwen=视觉大模型(失败回退 pix2tex)，skip=不识别")
    ap.add_argument("--api-key-env", default="DASHSCOPE_API_KEY",
                    help="公式识别视觉模型的 Key 环境变量（默认 "
                         "DASHSCOPE_API_KEY）")
    ap.add_argument("--append-formulas", default=None,
                    help="把图片公式 LaTeX 追加到该 formulas.md（默认 "
                         "输出目录 <名>_formulas.md）")
    ap.add_argument("--yolo-model", default="auto",
                    help="本地 YOLO 权重路径（默认 auto：探测当前目录/用户目录）")
    ap.add_argument("--yolo-conf", type=float, default=0.25,
                    help="YOLO 检测置信度阈值（默认 0.25）")
    ap.add_argument("--no-gallery", action="store_true",
                    help="关闭图片集整理（不生成 images/ 目录与 images.md 清单）")
    ap.add_argument("--no-clean", action="store_true",
                    help="不清理输出目录旧产物（默认启动时清理 by_page/images/"
                         "discarded 等本命令生成的内容，防止旧背景图等残留）")
    ap.add_argument("--json", action="store_true",
                    help="结构化统计输出到 stdout（核心库状态输出被抑制）")
    ap.add_argument("--version", action="version", version=VERSION)
    args = ap.parse_args(argv)

    pptx_norm = resolve_input(args.pptx)
    if pptx_norm is None:
        return EXIT_USAGE
    out = resolve_output(args.output, pptx_norm, "_images")

    _check_env(args)
    if not args.no_clean:
        _clean_output(Path(out))
    try:
        cb = _img_progress
        filter_cb = make_progress("图片过滤")
        gallery_cb = make_progress("图片清单")
        if args.json:
            with quiet_stdout():
                records = E.extract(
                    pptx_norm, out,
                    convert=not args.no_convert,
                    all_media=args.all_media,
                    with_fill=not args.no_fill,
                    with_bg_layout=not args.no_bg_layout,
                    rasterize=args.rasterize_vector,
                    raster_dpi=args.raster_dpi,
                    raster_prefer=args.raster_prefer,
                    crop=not args.no_crop,
                    min_crop=args.min_crop,
                    on_progress=cb)
                filtered = []
                yolo_used = False
                if not args.no_filter:
                    by_page = Path(out) / "by_page"
                    kept, filtered, yolo_used = img_filter.filter_images(
                        records, by_page, Path(out),
                        min_size=args.min_size, max_ratio=args.max_ratio,
                        min_colors=args.min_colors,
                        min_area=args.min_area,
                        max_sparse_ink=args.max_sparse_ink,
                        vec_min_area=args.vec_min_area,
                        keep_background=args.keep_background,
                        keep_fill=args.keep_fill,
                        keep_vec_bitmap=args.keep_vec_bitmap,
                        yolo_model=args.yolo_model, yolo_conf=args.yolo_conf,
                        on_progress=filter_cb)
                    records = kept
                # 方案B：矢量图 PowerPoint 渲染（公式版→LaTeX 追加 formulas；
                # 保留矢量→规范化 svg/wmf）
                vec_extra = {"formulas": 0, "rendered": 0, "normalized": 0}
                if not args.no_filter:
                    vec_extra = _process_vectors(
                        records, by_page, Path(out), Path(pptx_norm).stem,
                        filtered, args)
                n_images = 0
                images_md = ""
                if not args.no_gallery:
                    n_images, images_md = img_filter.build_image_gallery(
                        records, Path(out) / "by_page", Path(out),
                        Path(pptx_norm).stem,
                        yolo_model=args.yolo_model, yolo_conf=args.yolo_conf,
                        on_progress=gallery_cb)
            stat = _stats(records, out, filtered, yolo_used)
            stat["images"] = n_images
            stat["vec_formulas"] = vec_extra["formulas"]
            stat["vec_rendered"] = vec_extra["rendered"]
            stat["vec_normalized"] = vec_extra["normalized"]
            stat["images_md"] = images_md
            print_json(stat)
        else:
            records = E.extract(
                pptx_norm, out,
                convert=not args.no_convert,
                all_media=args.all_media,
                with_fill=not args.no_fill,
                with_bg_layout=not args.no_bg_layout,
                rasterize=args.rasterize_vector,
                raster_dpi=args.raster_dpi,
                raster_prefer=args.raster_prefer,
                crop=not args.no_crop,
                min_crop=args.min_crop,
                on_progress=cb)
            if not args.no_filter:
                by_page = Path(out) / "by_page"
                kept, filtered, yolo_used = img_filter.filter_images(
                    records, by_page, Path(out),
                    min_size=args.min_size, max_ratio=args.max_ratio,
                    min_colors=args.min_colors,
                    min_area=args.min_area,
                    max_sparse_ink=args.max_sparse_ink,
                    vec_min_area=args.vec_min_area,
                    keep_background=args.keep_background,
                    keep_fill=args.keep_fill,
                    keep_vec_bitmap=args.keep_vec_bitmap,
                    yolo_model=args.yolo_model, yolo_conf=args.yolo_conf,
                    on_progress=filter_cb)
                print(f"[OK] 取舍过滤：保留 {len(kept)} 张，"
                      f"舍弃 {len(filtered)} 张"
                      f"{'（YOLO 已启用）' if yolo_used else '（YOLO 未启用，仅规则层）'}")
                print(f"     明细：{Path(out) / 'filter_report.json'}")
                records = kept
                # 方案B：矢量图 PowerPoint 渲染（公式版→LaTeX 追加 formulas；
                # 保留矢量→规范化 svg/wmf）
                vec_extra = _process_vectors(
                    records, by_page, Path(out), Path(pptx_norm).stem,
                    filtered, args)
                if vec_extra["formulas"]:
                    print(f"[OK] 图片公式识别：{vec_extra['formulas']} 条"
                          f"已追加到 formulas.md")
                if vec_extra["rendered"]:
                    print(f"[OK] 矢量渲染：{vec_extra['rendered']} 张 → PNG")
                if vec_extra["normalized"]:
                    print(f"[OK] 矢量规范化：{vec_extra['normalized']} 张 "
                          f"→ {args.vector_out}")
            if not args.no_gallery:
                n_images, images_md = img_filter.build_image_gallery(
                    records, Path(out) / "by_page", Path(out),
                    Path(pptx_norm).stem,
                    yolo_model=args.yolo_model, yolo_conf=args.yolo_conf,
                    on_progress=gallery_cb)
                print(f"[OK] 图片集：{n_images} 张 → {Path(out) / 'images'}")
                print(f"     清单：{images_md}")
    except Exception as e:  # 核心库内部已优雅降级，此处兜底
        print(f"[错误] 图片提取失败：{e}", file=sys.stderr)
        return EXIT_ERR
    banner_end("pptx-img")
    return EXIT_OK




if __name__ == "__main__":
    raise SystemExit(main())
