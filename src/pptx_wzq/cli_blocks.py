#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
cli_blocks.py — pptx-blocks：把 PPT 页面解析为可视逻辑块（Visual Block）
========================================================================

输入：原始 pptx + 原子对象（extract_pptx_images 输出的 atomic_objects.json）
     + 页面文本（<名>_texts.md）+ 公式（<名>_formulas.md）
流程：空间聚类 → 块渲染 → VLM 类型识别 + Semantic Captioning → 拓扑生成
     → 跨模态关系 → 输出 <名>_visual_blocks.json + <名>_captions.md
     （captions 条目对象从单图变为单块，合并原 image captioning）

输出 JSON schema：pptx_multimodal_slide_v2.0（方案 §2），全课件一文件。

用法：
    pptx-blocks <产物目录> [--pptx a.pptx] [--output vb.json]
               [--texts a_texts.md] [--formulas a_formulas.md]
               [--model deepseek-v4-flash] [--base-url …] [--api-key-env …]
               [--no-vlm] [--resume] [--json] [--version]

退出码：0 成功 / 1 处理异常 / 2 参数或环境错误。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from pathlib import Path

from pptx_wzq.cli_common import (EXIT_ERR, EXIT_OK, EXIT_USAGE, print_json,
                                 quiet_stdout)
from pptx_wzq import visual_blocks as VB

VERSION = "pptx-blocks 1.0.0 (可视逻辑块全栈解析 + Semantic Captioning 合并)"
DEFAULT_MODEL = "qwen3.7-plus"
DEFAULT_BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"
DEFAULT_KEY_ENV = "DASHSCOPE_API_KEY"


def _split_pages(content: str) -> dict:
    """按 '## 第 N 页/节' 分块，返回 {page: 文本块}。"""
    pages = {}
    cur = None
    for line in content.splitlines():
        m = re.match(r"^##\s*第\s*(\d+)\s*(?:页|节)\s*$", line.strip())
        if m:
            cur = int(m.group(1))
            pages.setdefault(cur, [])
        elif cur is not None:
            pages[cur].append(line)
    return {k: "\n".join(v) for k, v in pages.items()}


def _load_texts(texts_path: Path) -> dict:
    """读 texts.md → {page: 页文本}。"""
    if not texts_path or not texts_path.is_file():
        return {}
    return _split_pages(texts_path.read_text(encoding="utf-8"))


def _load_formulas(formulas_path: Path) -> dict:
    """读 formulas.md → {page: 公式文本}。"""
    if not formulas_path or not formulas_path.is_file():
        return {}
    return _split_pages(formulas_path.read_text(encoding="utf-8"))


def _write_captions(captions_path: Path, slides: list, model: str) -> int:
    """把每页每块写成 captions.md（条目对象=单块），返回块数。"""
    n = 0
    lines = ["# images 图片 AI 解读（可视逻辑块级）", "",
             f"> 由 `pptx-blocks` 生成 · 模型 `{model}` · "
             f"每块一条；块渲染图见 images/ 目录。", ""]
    for s in slides:
        page = s["page"]
        blocks = s["blocks"]
        lines.append(f"## 第 {page} 页")
        lines.append("")
        for bi, blk in enumerate(blocks, start=1):
            n += 1
            img_name = f"slide_{page:02d}_{blk['block_id']}.png"
            sd = blk.get("semantic_description") or {}
            goal = sd.get("expression_goal", "")
            role = sd.get("expression_role", "")
            cap = sd.get("vlm_caption", "")
            feat = "、".join(sd.get("expression_features") or [])
            use = sd.get("teaching_use", "")
            lines.append(f"### IMG{n:04d} — `{img_name}`  ✅")
            lines.append("")
            lines.append(f"**块类型**：{blk.get('block_type')}。")
            if goal:
                lines.append(f"**表达目标**：{goal}")
            if role:
                lines.append(f"**表达作用**：{role}")
            if feat:
                lines.append(f"**表达特征**：{feat}。")
            lines.append(f"**内容理解**：{cap}")
            if use:
                lines.append(f"**教学用途**：{use}")
            lines.append("")
        lines.append("")
    captions_path.write_text("\n".join(lines), encoding="utf-8")
    return n


def _assemble_slides(slides: list, page_texts: dict,
                     page_formulas: dict, pptx_path: str,
                     stem: str, image_dir: Path, source_dir: Path) -> list:
    """组装 pptx_multimodal_slide_v2.0 的 slides[]（含 slide_info / textual_content
    / assets 路径 / reading_order）。"""
    out = []
    for s in slides:
        page = s["page"]
        raw = page_texts.get(page, "")
        fm = page_formulas.get(page, "")
        blocks = s["blocks"]
        # assets：块渲染图 + 内部资源路径
        for blk in blocks:
            img_name = f"slide_{page:02d}_{blk['block_id']}.png"
            blk.setdefault("assets", {})
            b = blk.get("bbox") or {}
            if b.get("w") and b.get("h"):
                blk["assets"]["rendered_image"] = f"./images/{img_name}"
            else:
                blk["assets"]["rendered_image"] = None
            vec = blk.get("vector_svg")
            if vec:
                blk["assets"]["vector_svg"] = f"./sources/{vec}"
            else:
                blk["assets"]["vector_svg"] = None
            blk["assets"]["internal_resources"] = []
        # reading_order：标题/正文 shape + 块
        order = [f"title_shape_{page}", f"text_body_{page}"] + \
            [b["block_id"] for b in blocks]
        out.append({
            "slide_info": {
                "slide_id": f"slide_{page:03d}",
                "slide_index": page,
                "layout_name": "",
                "title": _page_title(raw),
                "notes_text": "",
            },
            "textual_content": {
                "raw_text": raw + (f"\n[公式] {fm}" if fm else ""),
                "semantic_summary": "",
                "reading_order": order,
            },
            "visual_blocks": blocks,
            "cross_modal_relations": s.get("relations", []),
        })
    return out


def _page_title(raw_text: str) -> str:
    """取页面首行非空文本作为标题（启发式）。"""
    for ln in raw_text.splitlines():
        ln = ln.strip()
        if ln:
            return ln[:60]
    return ""


def _locate(args) -> tuple:
    """定位输入：atomic_objects / texts / formulas / pptx。"""
    out_dir = Path(args.dir)
    atomic = args.atomic_objects
    if atomic is None:
        atomic = VB.load_atomic_objects(out_dir) or None
    texts = args.texts
    if texts is None:
        cands = sorted(out_dir.glob("*_texts.md")) + \
            sorted(out_dir.rglob("*_texts.md"))
        texts = cands[0] if cands else None
    formulas = args.formulas
    if formulas is None:
        cands = sorted(out_dir.glob("*_formulas.md")) + \
            sorted(out_dir.rglob("*_formulas.md"))
        formulas = cands[0] if cands else None
    return atomic, texts, formulas


def _main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        prog="pptx-blocks",
        description="把 PPT 页面解析为可视逻辑块（空间聚类 + VLM 类型识别/"
                    "Semantic Captioning + 拓扑 + 跨模态关系）→ "
                    "<名>_visual_blocks.json + <名>_captions.md")
    ap.add_argument("dir", nargs="?", default=".",
                    help="产物目录（自动找 atomic_objects/texts/formulas）")
    ap.add_argument("--atomic-objects", type=Path, default=None,
                    help="原子对象 json（默认自动查找）")
    ap.add_argument("--texts", type=Path, default=None, help="页面文本 md")
    ap.add_argument("--formulas", type=Path, default=None, help="公式 md")
    ap.add_argument("--pptx", type=Path, default=None,
                    help="原始 pptx（用于渲染块 PNG；缺省尝试从目录找）")
    ap.add_argument("-o", "--output", default=None,
                    help="输出 json 路径（默认 <目录>/<名>_visual_blocks.json）")
    ap.add_argument("--captions", default=None,
                    help="输出 captions.md 路径（默认 <目录>/<名>_captions.md）")
    ap.add_argument("--model", default=DEFAULT_MODEL,
                    help=f"视觉模型（默认 {DEFAULT_MODEL}）")
    ap.add_argument("--base-url", default=DEFAULT_BASE_URL,
                    help=f"OpenAI 兼容端点（默认 {DEFAULT_BASE_URL}）")
    ap.add_argument("--api-key-env", default=DEFAULT_KEY_ENV,
                    help=f"API Key 环境变量名（默认 {DEFAULT_KEY_ENV}）")
    ap.add_argument("--no-vlm", action="store_true",
                    help="不调用 VLM，纯规则聚类/模板描述（0 Token）")
    ap.add_argument("--resume", action="store_true",
                    help="断点续跑：已有 visual_blocks.json 时保留完成页")
    ap.add_argument("--json", action="store_true",
                    help="把统计输出到 stdout")
    ap.add_argument("--version", action="version", version=VERSION)
    args = ap.parse_args(argv)

    out_dir = Path(args.dir)
    atomic, texts_path, formulas_path = _locate(args)
    if not atomic:
        print("[错误] 未找到 atomic_objects.json"
              "（请先运行 pptx-img 或指定 --atomic-objects）",
              file=sys.stderr)
        return EXIT_USAGE
    page_texts = _load_texts(Path(texts_path)) if texts_path else {}
    page_formulas = _load_formulas(Path(formulas_path)) if formulas_path else {}

    # 模型 client
    client = None
    if not args.no_vlm:
        api_key = os.environ.get(args.api_key_env, "")
        if not api_key:
            print(f"[警告] 未设置 {args.api_key_env}，"
                  f"使用规则聚类/模板描述（--no-vlm 等效）",
                  file=sys.stderr)
        else:
            try:
                from openai import OpenAI
                client = OpenAI(api_key=api_key, base_url=args.base_url)
            except Exception as e:
                print(f"[警告] OpenAI 客户端初始化失败：{e}，使用规则模式",
                      file=sys.stderr)
                client = None

    t0 = time.time()
    slides = VB.extract_blocks(
        str(args.pptx) if args.pptx else out_dir.name,
        str(out_dir),
        atomic_objects=atomic,
        page_texts=page_texts,
        client=client,
        model=args.model,
        on_progress=None,
    )
    # 块渲染 PNG（放 images/ 统一目录）
    stem = ""
    if texts_path:
        stem = Path(texts_path).stem[:-len("_texts")] \
            if str(texts_path).endswith("_texts.md") else Path(texts_path).stem
    image_dir = out_dir / "images"
    image_dir.mkdir(parents=True, exist_ok=True)
    # 清空旧块渲染图（块编号随规则变化，避免旧图残留造成 引用数≠文件数）
    for old in image_dir.glob("slide_*_blk_*.png"):
        try:
            old.unlink()
        except OSError:
            pass
    n_rendered = 0
    if args.pptx and args.pptx.is_file():
        # 单次整页渲染（缓存到课件旁 .render_cache），再按块 bbox 裁剪，
        # 避免每个块重复跑 LibreOffice（237 块 × 整页渲染会极慢）
        from pptx_wzq import extract_pptx_images as E
        cache = Path(args.pptx).parent / ".render_cache"
        pages = E.render_pptx_pages(str(args.pptx), cache, dpi=150)
        # 真实页面尺寸（EMU）：16:9 等非常规比例必须传入，否则裁剪
        # 按 4:3 默认值换算会把块的左侧内容切掉（战略管理实测）
        sld_cx, sld_cy = E.read_sld_size(str(args.pptx))
        if pages:
            for s in slides:
                for blk in s["blocks"]:
                    img_path = image_dir / \
                        f"slide_{s['page']:02d}_{blk['block_id']}.png"
                    if img_path.is_file():
                        n_rendered += 1
                        continue
                    ok = _crop_block_png(pages, s["page"], blk, img_path, 150,
                                         sld_cx=sld_cx, sld_cy=sld_cy)
                    if ok:
                        n_rendered += 1

    # 兜底清理：images/ 只保留 JSON 中可视逻辑块对应的渲染图
    # （块编号随规则变化后，旧图/无结构引用图必须移除，保证目录与
    # JSON 双向一致——缺失 0、多余 0）
    want = {f"slide_{s['page']:02d}_{blk['block_id']}.png"
            for s in slides for blk in s["blocks"]}
    removed = 0
    for old in image_dir.glob("slide_*_blk_*.png"):
        if old.name not in want:
            try:
                old.unlink()
                removed += 1
            except OSError as e:
                print(f"[警告] 清理未被 JSON 引用的渲染图失败："
                      f"{old.name}（{e}）", file=sys.stderr)
    if removed:
        print(f"[渲染] 清理未被 JSON 引用的渲染图 {removed} 张",
              file=sys.stderr)

    # 组装最终 JSON
    out_slides = _assemble_slides(slides, page_texts, page_formulas,
                                  args.pptx, stem, image_dir, out_dir / "sources")
    summary = {
        "slides": len(out_slides),
        "blocks_total": sum(len(s["visual_blocks"]) for s in out_slides),
        "slides_with_blocks": sum(
            1 for s in out_slides if s["visual_blocks"]),
        "block_types": _block_type_counter(out_slides),
    }
    binding = {
        "$schema": "pptx_multimodal_slide_v2.0",
        "stem": stem or out_dir.name,
        "tool_version": "pptx-wzq 1.5.0",
        "generated_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
        "slides": out_slides,
        "summary": summary,
    }
    out_path = Path(args.output) if args.output else \
        out_dir / f"{stem or 'slide'}_visual_blocks.json"
    out_path.write_text(json.dumps(binding, ensure_ascii=False, indent=1),
                        encoding="utf-8")

    # captions.md（块级）
    cap_path = Path(args.captions) if args.captions else \
        out_dir / f"{stem or 'slide'}_captions.md"
    n_cap = _write_captions(cap_path, slides, args.model)

    dt = time.time() - t0
    print(f"[OK] 可视逻辑块解析完成：{summary['slides']} 页 / "
          f"块 {summary['blocks_total']} 个 / 渲染图 {n_rendered} 张",
          file=sys.stderr)
    print(f"     JSON → {out_path}", file=sys.stderr)
    print(f"     captions → {cap_path}（{n_cap} 条）", file=sys.stderr)
    if args.json:
        print_json({**summary, "rendered": n_rendered,
                    "output": str(out_path), "captions": str(cap_path),
                    "cost_s": round(dt, 1)})
    return EXIT_OK


def _block_type_counter(slides: list) -> dict:
    from collections import Counter
    c = Counter()
    for s in slides:
        for blk in s["visual_blocks"]:
            c[blk.get("block_type", "unknown")] += 1
    return dict(c)


def _crop_block_png(pages: list, page_no: int, blk: dict,
                    out_png: Path, dpi: int = 150,
                    sld_cx=None, sld_cy=None) -> bool:
    """从整页渲染结果中按块 bbox 裁剪出 PNG（px→EMU）。
    块 bbox 无效（宽或高为 0）时返回 False（无坐标无法裁剪）。
    sld_cx/sld_cy：真实页面尺寸（EMU），缺失时 crop_page_png 按 4:3
    默认值换算——16:9 等非常规页面会把块的左侧内容切掉，必须传入。"""
    try:
        from pptx_wzq import extract_pptx_images as E
        if (page_no - 1) >= len(pages):
            return False
        b = blk.get("bbox") or {"x": 0, "y": 0, "w": 0, "h": 0}
        if b["w"] <= 0 or b["h"] <= 0:
            return False
        xfrm = (int(b["x"] / 96 * 914400), int(b["y"] / 96 * 914400),
                int(b["w"] / 96 * 914400), int(b["h"] / 96 * 914400))
        ok, _, _ = E.crop_page_png(pages[page_no - 1], xfrm, dpi, out_png,
                                   sld_cx=sld_cx, sld_cy=sld_cy)
        return bool(ok)
    except Exception:
        return False


def main() -> int:  # console
    from pptx_wzq.cli_common import banner, banner_end
    banner("pptx-blocks")
    rc = _main()
    banner_end("pptx-blocks")
    return rc


if __name__ == "__main__":
    sys.exit(_main())
