# -*- coding: utf-8 -*-
"""LaTeX 命令分隔工具（vendor 自 pptkb.formula_math.latex，仅取所需函数）。

来源（只读参考，未改动原文件）：
  C:\\jameswu\\教学\\模拟电子技术\\build-ppt-multimodal-kb-skill\\src\\pptkb\\formula_math\\latex.py
"""
from __future__ import annotations

import re


def separate_commands(latex: str, commands: set[str]) -> str:
    r"""命令与后续字母分隔（\Deltau → \Delta u），但绝不拆分完整白名单命令。

    算法：按 \\ 后完整字母 token 处理——token 本身在白名单 → 原样保留；
    否则按最长白名单前缀切分（如 Deltau → Delta + u）。确定性、幂等。
    """
    cmd_names = {c.lstrip("\\") for c in commands if c.startswith("\\")}

    def repl(m):
        tok = m.group(1)
        if tok in cmd_names:
            return "\\" + tok
        for i in range(len(tok), 0, -1):
            if tok[:i] in cmd_names:
                return "\\" + tok[:i] + " " + tok[i:]
        return m.group(0)

    return re.sub(r"\\([a-zA-Z]+)", repl, latex)
