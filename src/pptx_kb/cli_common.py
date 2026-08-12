#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
cli_common.py — pptx-img / pptx-formula 公共 CLI 工具（方案 B 轻量形态）

两个薄壳脚本共享：路径归一化（转调核心库）、退出码约定、--json 输出、
以及「--json 模式下吞掉核心库 print 状态输出」的上下文管理器。
核心库 extract_pptx_images.py 保持原样，本文件不改动任何核心逻辑。
"""
from __future__ import annotations

import contextlib
import io
import json
import sys
from pathlib import Path

from pptx_kb import extract_pptx_images as E

# 退出码约定：0 成功 / 1 处理异常 / 2 参数或文件错误（沿用旧 CLI）
EXIT_OK = 0
EXIT_ERR = 1
EXIT_USAGE = 2


def resolve_input(path_arg: str):
    """归一化输入路径；文件不存在时打印错误并返回 None。"""
    p = E.normalize_path(path_arg)
    if not Path(p).is_file():
        print(f"[错误] 找不到文件：{path_arg}（归一化后：{p}）", file=sys.stderr)
        return None
    return p


def resolve_output(out_arg: str | None, pptx_norm: str, default_suffix: str) -> str:
    """输出目录：显式 -o 则归一化；否则 <输入名><default_suffix>。"""
    if out_arg:
        return E.normalize_path(out_arg)
    return Path(pptx_norm).stem + default_suffix


@contextlib.contextmanager
def quiet_stdout():
    """吞掉核心库 print 的状态输出（--json 模式用）；stderr 保留（警告仍可见）。"""
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        yield


def print_json(obj) -> None:
    """以 UTF-8 缩进 JSON 输出到 stdout。"""
    print(json.dumps(obj, ensure_ascii=False, indent=2))


def make_progress(what: str):
    """返回核心库进度回调 on_progress(page_no, n_slides, info)。

    进度信息统一打印到 **stderr** 并实时 flush：普通模式与 --json 模式下
    用户都能在 cmd 窗口看到处理过程；同时不污染 --json 的 stdout。
    """
    def cb(page_no, n_slides, info):
        if info.get("kind") == "img":
            print(f"[进度] {what}：第 {page_no}/{n_slides} 页，"
                  f"累计提取对象 {info.get('objects', 0)} 个",
                  file=sys.stderr, flush=True)
        else:
            print(f"[进度] {what}：第 {page_no}/{n_slides} 页，"
                  f"本页公式 {info.get('page_entries', 0)} 条，"
                  f"累计 {info.get('total', 0)} 条",
                  file=sys.stderr, flush=True)
    return cb
