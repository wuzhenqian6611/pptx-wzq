"""cli_text.py — pptx-text 薄壳：逐页提取 PPT 文本到 <名>_texts.md。

用法：
    pptx-text <file.pptx> [-o out] [--no-filter] [--min-len N]
             [--no-tables] [--json] [--version]

输出：
    <名>_texts.md          每页文本清单（每对象一行，ID 页内递增）
    <名>_text_entries.json 全部条目审计（含被排除项及原因）

过滤（默认开，--no-filter 关闭）：页眉/页脚/页码/日期占位符、
母版/布局固定文本、跨页全局固定文本（≥90% 页面相同）、过短碎片。

启动环境检查（stderr，不阻断）：Python / lxml 或标准库 ET。
退出码：0 成功 / 1 处理异常 / 2 参数或文件错误。

作者：吴振谦（宁波大学科学技术学院教务部 · wuzhenqian@nbu.edu.cn）"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from pptx_kb import extract_texts as T
from pptx_kb.cli_common import (EXIT_ERR, EXIT_OK, EXIT_USAGE,
                        make_progress, print_json, quiet_stdout,
                        resolve_input, resolve_output,
                        banner, banner_end)

VERSION = "pptx-text 1.0.0 (方案B薄壳)"


def _check_env(args) -> None:
    """启动环境检查（stderr，不阻断）。"""
    try:
        import xml.etree.ElementTree  # noqa
        print("[环境] ElementTree(XML): OK", file=sys.stderr)
    except Exception:
        print("[环境] ElementTree(XML): 缺失（无法解析 slide XML）",
              file=sys.stderr)
    if not args.no_tables:
        # 表格依赖 ElementTree 对 a:tbl 的遍历，无额外依赖
        pass


def _main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        prog="pptx-text",
        description="逐页提取 PPT 文本（每文本对象一行，页内 ID，"
                    "排除页眉页脚/页码/母版固定文本）")
    ap.add_argument("pptx", help="输入的 .pptx 文件路径")
    ap.add_argument("-o", "--output", default=None,
                    help="输出目录（默认 <输入名>_text）")
    ap.add_argument("--no-filter", action="store_true",
                    help="关闭排除过滤（页眉页脚/母版固定文本等全部提取）")
    ap.add_argument("--min-len", type=int, default=2,
                    help="过滤：文本字符数小于该值视为过短碎片（默认 2）")
    ap.add_argument("--no-tables", action="store_true",
                    help="不提取表格文本（默认提取，行以 | 连接）")
    ap.add_argument("--json", action="store_true",
                    help="结构化统计输出到 stdout（核心库状态输出被抑制）")
    ap.add_argument("--version", action="version", version=VERSION)
    args = ap.parse_args(argv)

    pptx_norm = resolve_input(args.pptx)
    if pptx_norm is None:
        return EXIT_USAGE
    out = resolve_output(args.output, pptx_norm, "_text")

    _check_env(args)
    try:
        cb = make_progress("文本提取")
        if args.json:
            with quiet_stdout():
                result = T.extract_texts(
                    pptx_norm, out,
                    filter_texts=not args.no_filter,
                    min_len=args.min_len,
                    with_tables=not args.no_tables,
                    on_progress=cb)
            print_json(result)
        else:
            T.extract_texts(
                pptx_norm, out,
                filter_texts=not args.no_filter,
                min_len=args.min_len,
                with_tables=not args.no_tables,
                on_progress=cb)
            print(f"[OK] 文本清单已写出：{Path(out) / (Path(pptx_norm).stem + '_texts.md')}")
        return EXIT_OK
    except Exception as e:
        print(f"[错误] 文本提取失败：{e}", file=sys.stderr)
        return EXIT_ERR


def main() -> int:  # console
    banner("pptx-text")
    rc = _main()
    banner_end("pptx-text")
    return rc


if __name__ == "__main__":
    sys.exit(_main())
