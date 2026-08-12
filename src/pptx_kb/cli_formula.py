#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
cli_formula.py — pptx-formula：PPTX 公式解析 CLI（方案 B 薄壳）

用法
----
  pptx-formula extract <input.pptx> [-o OUT] [选项]
  pptx-formula extract deck.pptx -o out --no-ocr --json

三路径级联（路径1 原生 OMML→omml2latex；路径2 Equation.3 OLE→MTEF 解码器；
路径3 LO 渲染+xfrm 裁剪+数学 OCR）由核心库 extract_latex() 执行，本脚本仅做
参数映射、依赖探测提示与结果输出。核心函数体零改动。
退出码：0 成功 / 1 处理异常 / 2 参数或文件错误。
"""
from __future__ import annotations

import argparse
import sys

from pptx_kb import extract_pptx_images as E
from pptx_kb.cli_common import (EXIT_ERR, EXIT_OK, EXIT_USAGE,
                        make_progress, print_json, quiet_stdout,
                        resolve_input, resolve_output)

VERSION = "pptx-formula 2.0.0 (方案B薄壳+v5符号判据过滤)"


def _formula_progress(page_no: int, n_slides: int, info: dict) -> None:
    """公式提取进度回调：逐页打印公式来源与 LaTeX 片段。"""
    items = info.get("page_items", []) or []
    total = info.get("total", 0)
    print(f"[第 {page_no}/{n_slides} 页] 提取 {len(items)} 条公式"
          f"（累计 {total}）：", file=sys.stderr)
    for it in items:
        src = it.get("source", "?")
        latex = it.get("latex", "") or ""
        name = it.get("name", "") or ""
        tag = f"[{name}] " if name else ""
        print(f"    ({src:>11s}) {tag}{latex}", file=sys.stderr)


def _probe_deps(args) -> None:
    """启动时探测三路径依赖，缺失打印 [提示] 到 stderr（不阻断执行）。"""
    if not args.no_eq3:
        if E._make_eq3_converter() is None:
            print("[提示] EQ3-MTEF 解码器不可用（缺 olefile 或 mtef_decoder），"
                  "路径2 关闭，Equation.3 公式将走路径3/占位。", file=sys.stderr)
    if not args.no_ocr:
        if E._make_math_ocr(args.ocr_engine) is None:
            print(f"[提示] 数学 OCR 不可用（engine={args.ocr_engine}；"
                  "需 MATHPIX_APP_ID/KEY 或已安装 pix2tex（pip install pix2tex）），"
                  "路径3 仅渲染+裁剪，公式将按路径1/2/占位处理。", file=sys.stderr)
    if E._make_latex_converter() is None:
        print("[提示] 未安装 omml2latex（pip install omml2latex），"
              "原生 OMML 公式将仅以占位进入 md（不崩溃）。", file=sys.stderr)


def _check_env(args) -> None:
    """启动环境检查（stderr，[环境] 前缀；只报告，不阻断）。"""
    import importlib.util as u
    checks = [
        ("olefile（路径2 OLE 解包必需）", u.find_spec("olefile") is not None),
        ("mtef_decoder（路径2 MTEF 解码，随命令内置）",
         u.find_spec("mtef_decoder") is not None),
        ("omml2latex（路径1 OMML→LaTeX）",
         E._make_latex_converter() is not None),
        ("pix2tex（路径3 数学 OCR）",
         E._make_math_ocr(args.ocr_engine) is not None),
    ]
    for name, ok in checks:
        print(f"[环境] {name}: {'OK' if ok else '缺失（该路径自动降级）'}", file=sys.stderr)
    soffice = next((t[1] for t in (E._detect_rasterizer("soffice") or [])
                    if t[0] == "soffice"), None)
    pdftoppm = next((t[1] for t in (E._detect_rasterizer("auto") or [])
                     if t[0] == "pdftoppm"), None)
    print(f"[环境] LibreOffice 渲染: {'OK (' + soffice + ')' if soffice else '缺失（路径3 渲染降级）'}",
          file=sys.stderr)
    print(f"[环境] pdftoppm（渲染转 PNG）: {'OK (' + pdftoppm + ')' if pdftoppm else '缺失（路径3 渲染降级）'}",
          file=sys.stderr)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        prog="pptx-formula",
        description="PPTX 公式解析：三路径级联把全部公式以 LaTeX 汇总为 <名>_formulas.md")
    ap.add_argument("pptx", help="输入的 .pptx 文件路径")
    ap.add_argument("-o", "--output", default=None,
                    help="输出目录（默认 <输入名>_formulas；与 pptx-img 共用时传同一目录）")
    ap.add_argument("--no-eq3", action="store_true",
                    help="关闭路径2（EQ3-MTEF 解析），强制 OLE 公式走路径3/占位")
    ap.add_argument("--no-ocr", action="store_true",
                    help="关闭路径3 的 OCR（仍渲染+裁剪便于人工核验，但不调 OCR）")
    ap.add_argument("--ocr-engine", default="auto",
                    choices=["auto", "mathpix", "pix2tex"],
                    help="指定优先 OCR 引擎（默认 auto）")
    ap.add_argument("--no-filter", action="store_true",
                    help="关闭符号判据过滤（保留全部提取结果，不写 filtered_entries.json）")
    ap.add_argument("--json", action="store_true",
                    help="把公式计数 dict 输出到 stdout（核心库状态输出被抑制）")
    ap.add_argument("--version", action="version", version=VERSION)
    args = ap.parse_args(argv)

    pptx_norm = resolve_input(args.pptx)
    if pptx_norm is None:
        return EXIT_USAGE
    out = resolve_output(args.output, pptx_norm, "_formulas")

    _probe_deps(args)
    _check_env(args)
    try:
        cb = _formula_progress
        if args.json:
            with quiet_stdout():
                result = E.extract_latex(pptx_norm, out,
                                         no_eq3=args.no_eq3, no_ocr=args.no_ocr,
                                         ocr_engine=args.ocr_engine,
                                         on_progress=cb,
                                         filter_formulas=not args.no_filter)
            print_json(result)
        else:
            E.extract_latex(pptx_norm, out,
                            no_eq3=args.no_eq3, no_ocr=args.no_ocr,
                            ocr_engine=args.ocr_engine,
                            on_progress=cb,
                            filter_formulas=not args.no_filter)
    except Exception as e:  # 核心库内部已优雅降级，此处兜底
        print(f"[错误] 公式提取失败：{e}", file=sys.stderr)
        return EXIT_ERR
    return EXIT_OK




if __name__ == "__main__":
    raise SystemExit(main())
