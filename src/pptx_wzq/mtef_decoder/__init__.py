# -*- coding: utf-8 -*-
"""MTEF → LaTeX 通用解码器包（vendor 自 pptkb.formula_math，只读参考未改原文件）。

提供：
  - parse_mtef(data: bytes) -> dict            # 解析 MTEF 字节 → LaTeX
  - eq3_ole_bytes_to_latex(ole_bytes) -> (latex, ok, status)
        # 从 Equation.3 OLE2 复合文档抽取 "Equation Native" 流并解析
        # 返回 (latex, ok, status)；ok=True 表示拿到了可用 LaTeX（status != failed）
依赖：olefile（抽取 OLE 流）；解析本身仅标准库 + re。
"""
from __future__ import annotations

import io

from .mtef import MTEFError, MTEFParser, parse_mtef

__all__ = ["parse_mtef", "eq3_ole_bytes_to_latex", "MTEFParser", "MTEFError"]


def _extract_equation_native(ole_bytes: bytes):
    """从 Equation.3 OLE2 复合文档抽取 'Equation Native' 流（MTEF 字节）。

    返回 bytes；找不到或无 olefile 时返回 None。
    """
    try:
        import olefile
    except ImportError:
        return None
    try:
        ole = olefile.OleFileIO(io.BytesIO(ole_bytes))
    except Exception:
        return None
    try:
        for name in ole.listdir():
            if "/".join(name).lower() == "equation native":
                return ole.openstream(name).read()
        return None
    except Exception:
        return None
    finally:
        try:
            ole.close()
        except Exception:
            pass


def eq3_ole_bytes_to_latex(ole_bytes: bytes):
    """Equation.3 OLE 字节 → (latex, ok, status)。

    - 抽取 Equation Native 流（MTEF）；
    - 解析为 LaTeX；
    - ok = (status != "failed") and latex 非空；needs_review 也视为可用（best-effort）。
    任何异常返回 ("", False, "error")（调用方据此降级到 path3 / 占位）。
    """
    try:
        native = _extract_equation_native(ole_bytes)
        if not native:
            return "", False, "no-native-stream"
        res = parse_mtef(native)
        latex = (res.get("latex") or "").strip()
        status = res.get("status", "failed")
        ok = status != "failed" and bool(latex)
        return latex, ok, status
    except Exception:
        return "", False, "error"
