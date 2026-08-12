# -*- coding: utf-8 -*-
"""MTEF 二进制解析器 → LaTeX（通用解码器，vendor 自 pptkb.formula_math.mtef）

来源（只读参考，未改动原文件）：
  C:\\jameswu\\教学\\模拟电子技术\\build-ppt-multimodal-kb-skill\\src\\pptkb\\formula_math\\mtef.py
本模块为自包含副本，供 extract_pptx_images.py 的 path2（EQ3-MTEF → LaTeX）使用。

支持 MTEF v3（Equation Editor 3.x；Wiris 官方规范）：
- Tag byte：低 4 位 = 记录类型，高 4 位 = 选项标志；
- CHAR: tag + [nudge] + [typeface+128] + [16-bit character] + [embell list]；
  - 文本字体 → 16-bit Unicode；符号字体（fnSYMBOL=6）→ Symbol 字体索引；
- TMPL: tag + [nudge] + [selector] + [variation] + [options(积分/围栏)] + [subobject list(END 终止)]；
- LINE/PILE/MATRIX/EMBELL/RULER/FONT/SIZE/TYPESIZE 按 v3 规范；
- 子对象顺序按 v3 类定义（Scr: 下标→上标；Frac: 分子→分母；BigOp: 主→上→下→符号；Root: 主→被开方数）。

不支持的模板/记录标记 needs_review（禁止伪造）。
"""
from __future__ import annotations

import re

from .latex_util import separate_commands

# 常用 LaTeX 命令（命令分隔用）
_CMD_NAMES = [
    r"\alpha", r"\beta", r"\gamma", r"\Gamma", r"\delta", r"\Delta",
    r"\varepsilon", r"\zeta", r"\eta", r"\theta", r"\Theta", r"\lambda",
    r"\Lambda", r"\mu", r"\nu", r"\xi", r"\pi", r"\Pi", r"\rho", r"\sigma",
    r"\Sigma", r"\tau", r"\phi", r"\varphi", r"\chi", r"\psi", r"\Psi",
    r"\omega", r"\Omega", r"\times", r"\div", r"\cdot", r"\pm", r"\approx",
    r"\neq", r"\leq", r"\geq", r"\infty", r"\int", r"\iint", r"\iiint",
    r"\sum", r"\prod", r"\partial", r"\nabla", r"\sqrt", r"\frac",
    r"\left", r"\right", r"\overline", r"\underline", r"\vec", r"\lim",
    r"\cap", r"\cup", r"\subset", r"\supseteq", r"\in", r"\notin",
    r"\cong", r"\equiv", r"\propto", r"\sim", r"\Rightarrow", r"\rightarrow",
    r"\leftarrow", r"\leftrightarrow", r"\uparrow", r"\downarrow",
    r"\langle", r"\rangle", r"\lceil", r"\rceil", r"\lfloor", r"\rfloor",
    r"\therefore", r"\ast", r"\perp", r"\parallel", r"\circ", r"\oplus",
    r"\otimes", r"\emptyset", r"\wedge", r"\vee", r"\neg", r"\not",
    r"\ldots", r"\cdots", r"\prime", r"\cong",
]

# ---- v3 模板选择器 ----
TM_ANGLE = 0; TM_PAREN = 1; TM_BRACE = 2; TM_BRACK = 3; TM_BAR = 4; TM_DBAR = 5
TM_FLOOR = 6; TM_CEILING = 7
TM_ROOT = 13; TM_FRACT = 14; TM_SCRIPT = 15
TM_UBAR = 16; TM_OBAR = 17
TM_LARROW = 18; TM_RARROW = 19; TM_BARROW = 20
TM_SINT = 21; TM_DINT = 22; TM_TINT = 23
TM_SSINT = 24; TM_DSINT = 25; TM_TSINT = 26
TM_UHBRACE = 27; TM_LHBRACE = 28
TM_SUM = 29; TM_ISUM = 30; TM_PROD = 31; TM_IPROD = 32
TM_COPROD = 33; TM_ICOPROD = 34; TM_UNION = 35; TM_IUNION = 36
TM_INTER = 37; TM_IINTER = 38; TM_LIM = 39
TM_LDIV = 40; TM_SLFRACT = 41; TM_INTOP = 42; TM_SUMOP = 43
TM_LSCRIPT = 44; TM_DIRAC = 45; TM_UARROW = 46; TM_OARROW = 47; TM_OARC = 48

_SELECTOR_NAMES = {
    0: "angle", 1: "paren", 2: "brace", 3: "brack", 4: "bar", 5: "dbar",
    6: "floor", 7: "ceiling", 13: "root", 14: "fract", 15: "script",
    16: "ubar", 17: "obar", 18: "larrow", 19: "rarrow", 20: "barrow",
    21: "sint", 22: "dint", 23: "tint", 24: "ssint", 25: "dsint", 26: "tsint",
    27: "uhbrace", 28: "lhbrace", 29: "sum", 30: "isum", 31: "prod",
    32: "iprod", 33: "coprod", 34: "icoprod", 35: "union", 36: "iunion",
    37: "inter", 38: "iinter", 39: "lim", 40: "ldiv", 41: "slfract",
    42: "intop", 43: "sumop", 44: "lscript", 45: "dirac", 46: "uarrow",
    47: "oarrow", 48: "oarc",
}

# ---- typeface 值（v3，去偏后）----
FN_TEXT = 1; FN_FUNCTION = 2; FN_VARIABLE = 3; FN_LCGREEK = 4
FN_UCGREEK = 5; FN_SYMBOL = 6; FN_VECTOR = 7; FN_NUMBER = 8
FN_MTEXTRA = 11; FN_TEXT_FE = 12

# Windows Symbol 字体映射（fnSYMBOL=6 的字体索引 → Unicode/LaTeX）
_SYMBOL_FONT = {
    0x20: " ", 0x22: r"\times", 0x24: r"\partial", 0x25: r"\Delta",
    0x26: r"\nabla", 0x27: "'", 0x28: "(", 0x29: ")", 0x2A: r"\ast",
    0x2B: "+", 0x2C: ",", 0x2D: "-", 0x2E: ".", 0x2F: "/",
    0x30: "0", 0x31: "1", 0x32: "2", 0x33: "3", 0x34: "4", 0x35: "5",
    0x36: "6", 0x37: "7", 0x38: "8", 0x39: "9", 0x3A: ":", 0x3B: ";",
    0x3C: "<", 0x3D: "=", 0x3E: ">", 0x3F: "?", 0x40: r"\cong",
    0x41: r"\Alpha", 0x42: r"\Beta", 0x43: r"\Chi", 0x44: r"\Delta",
    0x45: r"\Epsilon", 0x46: r"\Phi", 0x47: r"\Gamma", 0x48: r"\Eta",
    0x49: r"\Iota", 0x4A: r"\vartheta", 0x4B: r"\Kappa", 0x4C: r"\Lambda",
    0x4D: r"\Mu", 0x4E: r"\Nu", 0x4F: r"\Omicron", 0x50: r"\Pi",
    0x51: r"\Theta", 0x52: r"\Rho", 0x53: r"\Sigma", 0x54: r"\Tau",
    0x55: r"\Upsilon", 0x56: r"\varsigma", 0x57: r"\Omega", 0x58: r"\Xi",
    0x59: r"\Psi", 0x5A: r"\Zeta", 0x5B: "[", 0x5C: r"\therefore",
    0x5D: "]", 0x5E: r"\perp", 0x5F: "_", 0x60: r"\lnot",
    0x61: r"\alpha", 0x62: r"\beta", 0x63: r"\chi", 0x64: r"\delta",
    0x65: r"\varepsilon", 0x66: r"\phi", 0x67: r"\gamma", 0x68: r"\eta",
    0x69: r"\iota", 0x6A: r"\varphi", 0x6B: r"\kappa", 0x6C: r"\lambda",
    0x6D: r"\mu", 0x6E: r"\nu", 0x6F: r"\omicron", 0x70: r"\pi",
    0x71: r"\theta", 0x72: r"\rho", 0x73: r"\sigma", 0x74: r"\tau",
    0x75: r"\upsilon", 0x76: r"\varpi", 0x77: r"\omega", 0x78: r"\xi",
    0x79: r"\psi", 0x7A: r"\zeta", 0x7B: "{", 0x7C: "|", 0x7D: "}",
    0x7E: r"\sim",
    0xA3: r"\leq", 0xA5: r"\infty", 0xAB: r"\leftrightarrow",
    0xAC: r"\leftarrow", 0xAD: r"\uparrow", 0xAE: r"\rightarrow",
    0xAF: r"\downarrow", 0xB0: r"^{\circ}", 0xB1: r"\pm",
    0xB3: r"\geq", 0xB4: r"\times", 0xB5: r"\propto", 0xB6: r"\partial",
    0xB7: r"\cdot", 0xB8: r"\div", 0xB9: r"\neq", 0xBA: r"\equiv",
    0xBB: r"\approx", 0xBC: r"\ldots", 0xBE: "-",
    0xC4: r"\otimes", 0xC5: r"\oplus", 0xC6: r"\emptyset",
    0xC7: r"\cap", 0xC8: r"\cup", 0xC9: r"\supset", 0xCA: r"\supseteq",
    0xCB: r"\not\subset", 0xCC: r"\subset", 0xCD: r"\subseteq",
    0xCE: r"\in", 0xCF: r"\notin", 0xD0: r"\angle", 0xD1: r"\nabla",
    0xD5: r"\prod", 0xD6: r"\sqrt", 0xD7: r"\cdot", 0xD8: r"\neg",
    0xD9: r"\wedge", 0xDA: r"\vee", 0xDB: r"\Leftrightarrow",
    0xDC: r"\Leftarrow", 0xDD: r"\Uparrow", 0xDE: r"\Rightarrow",
    0xDF: r"\Downarrow", 0xE0: r"\lozenge",
    0xE1: r"\langle", 0xE2: r"\rangle", 0xE3: r"\lceil", 0xE4: r"\rceil",
    0xE5: r"\lfloor", 0xE6: r"\rfloor",
    0xF1: r"\leftrightarrow", 0xF2: r"\blacktriangle", 0xF3: r"\blacktriangledown",
}

# 数学风格 → LaTeX 字体命令
_GREEK_LC = {"α": r"\alpha", "β": r"\beta", "γ": r"\gamma", "δ": r"\delta",
             "ε": r"\varepsilon", "ζ": r"\zeta", "η": r"\eta", "θ": r"\theta",
             "ι": r"\iota", "κ": r"\kappa", "λ": r"\lambda", "μ": r"\mu",
             "ν": r"\nu", "ξ": r"\xi", "π": r"\pi", "ρ": r"\rho",
             "σ": r"\sigma", "τ": r"\tau", "υ": r"\upsilon", "φ": r"\varphi",
             "χ": r"\chi", "ψ": r"\psi", "ω": r"\omega"}
_GREEK_UC = {"Α": r"\Alpha", "Β": r"\Beta", "Γ": r"\Gamma", "Δ": r"\Delta",
             "Ε": r"\Epsilon", "Ζ": r"\Zeta", "Η": r"\Eta", "Θ": r"\Theta",
             "Ι": r"\Iota", "Κ": r"\Kappa", "Λ": r"\Lambda", "Μ": r"\Mu",
             "Ν": r"\Nu", "Ξ": r"\Xi", "Ο": r"\Omicron", "Π": r"\Pi",
             "Ρ": r"\Rho", "Σ": r"\Sigma", "Τ": r"\Tau", "Υ": r"\Upsilon",
             "Φ": r"\Phi", "Χ": r"\Chi", "Ψ": r"\Psi", "Ω": r"\Omega"}


# Unicode 数学符号兜底映射（Symbol 字体索引命中后仍缺失的扩展字符）
_UNICODE_MATH_FALLBACK = {
    0x2212: "-", 0x2248: r"\approx", 0x2265: r"\geq", 0x2264: r"\leq",
    0x2260: r"\neq", 0x00D7: r"\times", 0x221A: r"\sqrt", 0x221E: r"\infty",
    0x222B: r"\int", 0x2211: r"\sum", 0x220F: r"\prod", 0x2202: r"\partial",
    0x2207: r"\nabla", 0x00B1: r"\pm", 0x00F7: r"\div", 0x00B7: r"\cdot",
    0x2192: r"\rightarrow", 0x2190: r"\leftarrow", 0x2194: r"\leftrightarrow",
    0x21D2: r"\Rightarrow", 0x21D4: r"\Leftrightarrow", 0x2191: r"\uparrow",
    0x2193: r"\downarrow", 0x03B1: r"\alpha", 0x03B2: r"\beta",
    0x03B3: r"\gamma", 0x03B4: r"\delta", 0x03B5: r"\varepsilon",
    0x03B6: r"\zeta", 0x03B7: r"\eta", 0x03B8: r"\theta", 0x03B9: r"\iota",
    0x03BA: r"\kappa", 0x03BB: r"\lambda", 0x03BC: r"\mu", 0x03BD: r"\nu",
    0x03BE: r"\xi", 0x03C0: r"\pi", 0x03C1: r"\rho", 0x03C3: r"\sigma",
    0x03C4: r"\tau", 0x03C5: r"\upsilon", 0x03C6: r"\varphi", 0x03C7: r"\chi",
    0x03C8: r"\psi", 0x03C9: r"\omega", 0x0394: r"\Delta", 0x0393: r"\Gamma",
    0x03A3: r"\Sigma", 0x03A9: r"\Omega", 0x03A6: r"\Phi", 0x03A8: r"\Psi",
    0x0398: r"\Theta", 0x039B: r"\Lambda", 0x039E: r"\Xi", 0x03A0: r"\Pi",
    0x2205: r"\emptyset", 0x2229: r"\cap", 0x222A: r"\cup",
    0x2282: r"\subset", 0x2283: r"\supset", 0x2286: r"\subseteq",
    0x2287: r"\supseteq", 0x2208: r"\in", 0x2209: r"\notin",
    0x2220: r"\angle", 0x22A5: r"\perp", 0x2225: r"\parallel",
    0x2261: r"\equiv", 0x221D: r"\propto", 0x223C: r"\sim",
    0x2245: r"\cong", 0x00B0: r"^{\circ}", 0x2103: r"^{\circ}\mathrm{C}",
    0x2192: r"\rightarrow", 0x226A: r"\ll", 0x226B: r"\gg",
    0x22C5: r"\cdot", 0x2213: r"\mp", 0x22EF: r"\cdots", 0x2026: r"\ldots",
    0x2215: "/", 0x2032: "'", 0x2033: "''", 0x00A5: r"\yen",
    0x20AC: r"\euro", 0x00D7: r"\times", 0x2265: r"\geq",
}
# 私有区/未知符号检测
_PUA_RE = range(0xE000, 0xF900)

# Prime 后缀正则：LaTeX 数学模式中的 ' 与 ''（及 \prime 命令形式）
_PRIME_SEQ_RE = re.compile(r"^(?:(?:'|'')|\\prime\s*)+\s*$")
_PRIME_LEAD_RE = re.compile(r"^(?:'|'')|\\prime\s*")


def _split_prime_superscript(sup: str) -> tuple[str, str]:
    """把上标槽内容拆成 (prime 后缀部分, 其余指数部分)。

    规则：
    - 内容纯为 prime 序列（'、''、\\prime）→ (全部, "")，作为后缀直接附加；
    - 内容为 prime 序列 + 其他（如 '^2、''^n、'^{2}）→ (prime 部分, 其余)，
      prime 作为后缀、其余作为独立上标；
    - 其他 → ("", 原内容)（保持原样 ^ {...}）。
    Prime 必须是**后缀修饰符**，禁止生成 "上标的上标"（^{'} / ^{'^2}）。
    """
    sup = (sup or "").strip()
    if not sup:
        return "", ""
    if _PRIME_SEQ_RE.match(sup):
        return sup, ""
    m = _PRIME_LEAD_RE.match(sup)
    if m:
        primes = m.group(0).strip()
        rest = sup[m.end():].strip()
        if rest:
            return primes, rest
    return "", sup


def _char_latex(unicode_char: str, typeface: int) -> str:
    """字符 → LaTeX（按 typeface 处理）。"""
    if typeface in (FN_LCGREEK, FN_UCGREEK):
        return _GREEK_LC.get(unicode_char) or _GREEK_UC.get(unicode_char) or unicode_char
    return unicode_char


class MTEFError(Exception):
    pass


class MTEFParser:
    """MTEF v3 解析器（Equation Editor 3.x）。"""

    def __init__(self, data: bytes):
        self.data = data
        self.pos = 0
        self.errors: list[str] = []
        self.unsupported: list[str] = []
        self.template_count = 0
        self.unknown_records = 0
        self.object_count = 0
        self.record_count = 0
        self.embell_count = 0
        self.empty_slot_count = 0
        self._ended_cleanly = False  # 顶层对象列表是否以 END 记录正常终止
        self._font_map: dict[int, str] = {}  # 显式字体编号 → 名称
        self._depth = 0  # 递归深度保护（防 C 栈溢出硬崩溃）

    def _enter(self):
        self._depth += 1
        if self._depth > 150:
            raise MTEFError("nesting too deep")

    def _leave(self):
        self._depth -= 1

    # ---------- 字节读取 ----------

    def _byte(self) -> int:
        b = self.data[self.pos]
        self.pos += 1
        return b

    def _word(self) -> int:
        lo = self._byte()
        hi = self._byte()
        return lo | (hi << 8)

    def _skip(self, n: int) -> None:
        self.pos = min(self.pos + n, len(self.data))

    def _nudge(self) -> None:
        if self.pos + 1 >= len(self.data):
            return
        dx = self.data[self.pos]
        dy = self.data[self.pos + 1]
        if dx == 128 and dy == 128 and self.pos + 6 <= len(self.data):
            self.pos += 6
        else:
            self.pos += 2

    def _null_string(self) -> str:
        start = self.pos
        while self.pos < len(self.data) and self.data[self.pos] != 0:
            self.pos += 1
        s = self.data[start:self.pos].decode("latin-1", errors="replace")
        if self.pos < len(self.data):
            self.pos += 1
        return s

    # ---------- 主解析 ----------

    def parse(self) -> dict:
        """解析 → {latex, errors, status, template_count, has_cjk}。"""
        if len(self.data) < 5:
            return {"latex": "", "errors": ["mtef data too short"], "status": "failed",
                    "template_count": 0, "has_cjk": False,
                    "mtef_completeness": {
                        "total_bytes": len(self.data), "consumed_bytes": 0,
                        "unconsumed_tail_bytes": len(self.data), "object_count": 0,
                        "unknown_records": 0, "parse_completeness": 0.0}}
        # 跳过 EQNOLEFILEHDR（28 字节，Equation Native 流头）
        if len(self.data) >= 2 and self.data[0] == 0x1C and self.data[1] == 0x00:
            self.pos = 28
        else:
            self.pos = 0
        ver = self.data[self.pos] if self.pos < len(self.data) else 0
        self.pos += 5  # MTEF header: version/platform/product/ver/subver
        latex = ""
        try:
            latex = self._object_list()
        except Exception as e:
            self.errors.append(type(e).__name__)
        has_cjk = any("\u4e00" <= ch <= "\u9fff" for ch in latex)
        if not latex.strip():
            self.unsupported.append("empty-formula")
        if any(ord(ch) in _PUA_RE for ch in latex):
            self.unsupported.append("pua-symbols")
        # 命令分隔（\Deltau → \Delta u；完整命令如 \rightarrow 绝不拆分）
        latex = separate_commands(latex, set(_CMD_NAMES))
        # 全角标点/符号规范化（公式内分隔符与关系符）
        latex = (latex.replace("，", ", ").replace("；", "; ").replace("。", ". ")
                 .replace("＝", "=").replace("－", "-").replace("＞", ">")
                 .replace("＜", "<").replace("≥", r"\geq").replace("≤", r"\leq"))
        # 关系符叠加（Symbol 字体 >< 组合）
        latex = re.sub(r">\s*>\s*", r" \\gg ", latex)
        latex = re.sub(r"<\s*<\s*", r" \\ll ", latex)
        latex = re.sub(r">\s*=\s*", r" \\geq ", latex)
        latex = re.sub(r"<\s*=\s*", r" \\leq ", latex)
        latex = re.sub(r"\s+", " ", latex).strip()
        total_bytes = len(self.data)
        consumed = min(self.pos, total_bytes)
        tail = total_bytes - consumed
        completeness = round(consumed / total_bytes, 4) if total_bytes else 0.0
        if tail > 0:
            self.unsupported.append(f"unconsumed-tail:{tail}bytes")
        if self.unknown_records:
            self.unsupported.append(f"unknown-record:{self.unknown_records}")
        # 完整树消费检查（v8）：对象列表未以 END 正常终止、模板空槽（分式缺
        # 分母等）、未知记录 → 结构异常，标记疑似截断（不得当作完整公式）
        if not self._ended_cleanly:
            self.unsupported.append("unterminated-object-list")
        if self.empty_slot_count:
            self.unsupported.append(f"empty-template-slot:{self.empty_slot_count}")
        # 状态计算必须在全部 unsupported 追加之后（v8 修复：树消费异常必须降级）
        status = "failed" if self.errors else ("needs_review" if self.unsupported else "ok")
        return {"latex": latex, "errors": self.errors + self.unsupported,
                "status": status, "template_count": self.template_count,
                "has_cjk": has_cjk, "version": ver,
                "mtef_completeness": {
                    "total_bytes": total_bytes,
                    "consumed_bytes": consumed,
                    "unconsumed_tail_bytes": tail,
                    "object_count": self.object_count,
                    "record_count": self.record_count,
                    "embell_count": self.embell_count,
                    "empty_template_slots": self.empty_slot_count,
                    "unknown_records": self.unknown_records,
                    "ended_cleanly": self._ended_cleanly,
                    "parse_completeness": completeness,
                }}

    def _object_list(self) -> str:
        """解析对象列表（records 直到 END）。"""
        self._enter()
        try:
            return self._object_list_inner()
        finally:
            self._leave()

    def _object_list_inner(self) -> str:
        out: list[str] = []
        while self.pos < len(self.data):
            tag = self._byte()
            self.record_count += 1
            rtype = tag & 0x0F
            opts = tag >> 4
            if rtype == 0:  # END
                self._ended_cleanly = True
                return "".join(out)
            if rtype == 1:  # LINE
                if opts & 0x01:  # xfNULL: 无对象列表
                    continue
                if opts & 0x04:  # xfLSPACE
                    self._word()
                if opts & 0x02:  # xfRULER → RULER 记录
                    self._ruler()
                out.append(self._object_list())
            elif rtype == 2:  # CHAR
                self.object_count += 1
                out.append(self._char(opts))
            elif rtype == 3:  # TMPL
                self.object_count += 1
                self.template_count += 1
                out.append(self._tmpl(opts))
            elif rtype == 4:  # PILE
                self.object_count += 1
                if opts & 0x08:
                    self._nudge()
                self._skip(2)  # halign + valign
                if opts & 0x02:
                    self._ruler()
                out.append(self._object_list())
            elif rtype == 5:  # MATRIX
                self.object_count += 1
                if opts & 0x08:
                    self._nudge()
                self._skip(4)  # valign + h_just + v_just + rows
                if self.pos < len(self.data):
                    cols = self._byte()
                else:
                    cols = 1
                self._skip((cols + 1 + 1 + 1) // 2)  # col_parts
                if self.pos < len(self.data):
                    rows_byte = self.data[self.pos - 6] if self.pos >= 6 else 1
                out.append(self._object_list())
            elif rtype == 6:  # EMBELL
                self.embell_count += 1
                if opts & 0x08:
                    self._nudge()
                if self.pos < len(self.data):
                    self._byte()  # embell type
            elif rtype == 7:  # RULER
                self._ruler()
            elif rtype == 8:  # FONT: tface + style + name
                if self.pos + 1 < len(self.data):
                    tface = self._byte()
                    style = self._byte()
                    name = self._null_string()
                    self._font_map[tface] = name
            elif rtype == 9:  # SIZE
                self._size()
            elif rtype in (10, 11, 12, 13, 14):  # typesize
                pass
            else:
                self.unknown_records += 1
                self.unsupported.append(f"record:{rtype}")
                return "".join(out)
        return "".join(out)

    def _char(self, opts: int) -> str:
        if opts & 0x08:
            self._nudge()
        # typeface + 128
        tface_raw = self._byte() if self.pos < len(self.data) else 0
        typeface = tface_raw - 128
        char16 = self._word()
        ch = chr(char16) if char16 else ""
        # 符号字体：Symbol 索引映射，未命中用 Unicode 数学符号兜底
        if typeface == FN_SYMBOL:
            if char16 in _SYMBOL_FONT:
                ch = _SYMBOL_FONT[char16]
            elif char16 in _UNICODE_MATH_FALLBACK:
                ch = _UNICODE_MATH_FALLBACK[char16]
            elif char16 and char16 > 0x7E:
                ch = f"\\char{char16} "
            else:
                ch = chr(char16) if char16 else ""
        elif typeface == FN_VARIABLE:
            pass  # 数学变量：Latin/Unicode 直出
        elif typeface in (FN_LCGREEK, FN_UCGREEK):
            ch = _GREEK_LC.get(ch) or _GREEK_UC.get(ch) or ch
        elif typeface == FN_VECTOR:
            ch = r"\vec{" + ch + "}"
        # 显式字体（负 typeface）→ 记录但不改写
        if opts & 0x02:  # xfEMBELL: embellishment list
            if self.pos < len(self.data):
                self._embell_list()
        return ch

    def _embell_list(self) -> None:
        while self.pos < len(self.data):
            tag = self._byte()
            rtype = tag & 0x0F
            opts = tag >> 4
            if rtype == 0:
                return
            if rtype == 6:  # EMBELL
                if opts & 0x08:
                    self._nudge()
                if self.pos < len(self.data):
                    self._byte()
            else:
                return

    def _ruler(self) -> None:
        if self.pos >= len(self.data):
            return
        n = self._byte()
        self._skip(n * 3)

    def _size(self) -> None:
        if self.pos >= len(self.data):
            return
        lsize = self._byte()
        if lsize == 101:  # 显式磅值
            self._word()
        elif lsize == 100:  # 大增量
            self._skip(3)
        else:
            self._skip(1)

    def _tmpl(self, opts: int) -> str:
        self._enter()
        try:
            return self._tmpl_inner(opts)
        finally:
            self._leave()

    def _tmpl_inner(self, opts: int) -> str:
        if opts & 0x08:
            self._nudge()
        if self.pos >= len(self.data):
            self.errors.append("tmpl selector missing")
            return ""
        sel = self._byte()
        # variation（v3 实测 1 字节；0x80 标志 → 2 字节，兼容 v5）
        if self.pos < len(self.data):
            v0 = self.data[self.pos]
            if v0 & 0x80:
                self.pos += 1
                if self.pos < len(self.data):
                    v1 = self._byte()
                    variation = (v0 & 0x7F) | (v1 << 8)
                else:
                    variation = v0 & 0x7F
            else:
                variation = v0
                self.pos += 1
        else:
            variation = 0
        # tmpl-options（v3/v5 模板固定第 4 字节；漏消费会导致槽位错位）
        if self.pos < len(self.data):
            self.pos += 1
        # 槽位：LINE 记录创建新槽；xfNULL LINE → 空槽（不消费对象列表）；
        # 其余记录（CHAR/TMPL/typesize…）追加到当前槽（typesize 无内容不建槽）
        slots: list[str] = []
        while self.pos < len(self.data):
            b = self.data[self.pos]
            rtype = b & 0x0F
            ropts = b >> 4
            if rtype == 0:  # END → 模板结束
                self.pos += 1
                break
            if rtype == 1:  # LINE → 新槽位
                self.pos += 1
                if ropts & 0x04:  # xfLSPACE
                    self._word()
                if ropts & 0x02:  # xfRULER → RULER 记录
                    self._ruler()
                if ropts & 0x01:  # xfNULL 空槽（无对象列表）
                    slots.append("")
                    continue
                slots.append(self._object_list())
            else:
                piece = self._dispatch(b)
                if piece:
                    if not slots:
                        slots.append("")
                    slots[-1] += piece
        return self._template_to_latex(sel, variation, slots)

    def _dispatch(self, b: int) -> str:
        """处理非 LINE 顶层记录（模板子对象内）。b 为已 peek 的 tag 字节。"""
        rtype = b & 0x0F
        ropts = b >> 4
        self.pos += 1  # 消费 tag（b 已 peek）
        if rtype == 0:
            return ""
        if rtype == 2:  # CHAR
            return self._char(ropts)
        if rtype == 3:  # TMPL
            self.template_count += 1
            return self._tmpl(ropts)
        if rtype == 6:  # EMBELL
            if ropts & 0x08:
                self._nudge()
            if self.pos < len(self.data):
                self._byte()
            return ""
        if rtype == 7:
            self._ruler()
            return ""
        if rtype in (10, 11, 12, 13, 14):
            return ""
        if rtype == 8:
            self._skip_font()
            return ""
        if rtype == 9:
            self._size()
            return ""
        if rtype == 4:  # PILE
            if ropts & 0x08:
                self._nudge()
            self._skip(2)
            if ropts & 0x02:
                self._ruler()
            return self._object_list()
        if rtype == 5:  # MATRIX
            if ropts & 0x08:
                self._nudge()
            self._skip(4)
            if self.pos < len(self.data):
                cols = self._byte()
            else:
                cols = 1
            self._skip((cols + 1 + 1 + 1) // 2)
            return self._object_list()
        return ""

    def _template_to_latex(self, sel: int, variation: int, slots: list[str]) -> str:
        """模板 → LaTeX。必需槽位为空时记录 empty_template_slot（不得输出
        空花括号拼出的“语法合法但语义残缺”LaTeX，必须降级 needs_review）。"""
        def req(i: int, name: str) -> str:
            v = slots[i] if i < len(slots) else ""
            if not v.strip():
                self.unsupported.append(f"empty_template_slot:{name}")
                self.empty_slot_count += 1
            return v

        s = lambda i: slots[i] if i < len(slots) else ""
        if sel in (TM_SCRIPT, TM_LSCRIPT):
            # ScrBox 槽位顺序：下标→上标（v5 同）；单槽时按变体归属：
            # var 0=上标 1=下标 2=上下标；仅活动槽为空才标记 empty_template_slot
            sub = slots[0] if len(slots) >= 1 else ""
            sup = slots[1] if len(slots) >= 2 else ""
            if len(slots) == 1 and variation == 0:
                sup, sub = sub, ""
            out = ""
            if variation in (1, 2):
                if sub.strip():
                    out += "_{" + sub + "}"
                else:
                    self.unsupported.append("empty_template_slot:script")
                    self.empty_slot_count += 1
            if variation in (0, 2):
                if sup.strip():
                    # Prime 必须作为后缀修饰符，不得生成“上标的上标”：
                    #   u'  → u'        （prime 后缀）
                    #   u'^2 → u'^{2}   （prime 后缀 + 独立指数）
                    #   u'' → u''       （双 prime）
                    primes, rest = _split_prime_superscript(sup)
                    if primes and not rest:
                        out += primes
                    elif primes:
                        out += primes + "^{" + rest + "}"
                    else:
                        out += "^{" + sup + "}"
                else:
                    self.unsupported.append("empty_template_slot:script")
                    self.empty_slot_count += 1
            return out
        if sel in (TM_FRACT, TM_SLFRACT):
            return f"\\frac{{{req(0, 'frac')}}}{{{req(1, 'frac')}}}"
        if sel == TM_ROOT:
            if variation == 1:
                return f"\\sqrt[{s(1)}]{{{req(0, 'root')}}}"
            return f"\\sqrt{{{req(0, 'root')}}}"
        if sel in (TM_SINT, TM_DINT, TM_TINT, TM_SSINT, TM_DSINT, TM_TSINT):
            intcmd = {TM_SINT: r"\int", TM_DINT: r"\iint", TM_TINT: r"\iiint",
                      TM_SSINT: r"\int", TM_DSINT: r"\iint", TM_TSINT: r"\iiint"}[sel]
            if s(2) or s(1):
                return f"{intcmd}_{{{s(2)}}}^{{{s(1)}}} {s(0)}"
            return f"{intcmd} {s(0)}"
        if sel in (TM_SUM, TM_PROD, TM_COPROD, TM_UNION, TM_INTER,
                   TM_ISUM, TM_IPROD, TM_ICOPROD, TM_IUNION, TM_IINTER,
                   TM_INTOP, TM_SUMOP):
            opc = {TM_SUM: r"\sum", TM_PROD: r"\prod", TM_COPROD: r"\amalg",
                   TM_UNION: r"\bigcup", TM_INTER: r"\bigcap",
                   TM_ISUM: r"\sum", TM_IPROD: r"\prod", TM_ICOPROD: r"\amalg",
                   TM_IUNION: r"\bigcup", TM_IINTER: r"\bigcap",
                   TM_INTOP: r"\int", TM_SUMOP: r"\sum"}[sel]
            # BigOp: 主槽→上槽→下槽（下标 s(2) 上标 s(1)）
            if s(2) or s(1):
                return f"{opc}_{{{s(2)}}}^{{{s(1)}}} {s(0)}"
            return f"{opc} {s(0)}"
        if sel == TM_LIM:
            if s(1):
                return f"\\lim_{{{s(1)}}} {s(0)}"
            return f"\\lim {s(0)}"
        if sel in (TM_PAREN, TM_BRACE, TM_BRACK, TM_ANGLE, TM_BAR, TM_DBAR,
                   TM_FLOOR, TM_CEILING):
            lmap = {TM_PAREN: r"\left(", TM_BRACE: r"\left\{", TM_BRACK: r"\left[",
                    TM_ANGLE: r"\langle", TM_BAR: r"\left|", TM_DBAR: r"\left\|",
                    TM_FLOOR: r"\lfloor", TM_CEILING: r"\lceil"}[sel]
            rmap = {TM_PAREN: r"\right)", TM_BRACE: r"\right\}", TM_BRACK: r"\right]",
                    TM_ANGLE: r"\rangle", TM_BAR: r"\right|", TM_DBAR: r"\right\|",
                    TM_FLOOR: r"\rfloor", TM_CEILING: r"\rceil"}[sel]
            main = req(0, "fence")
            if variation == 1:  # 仅左分隔符 → \right. 空分隔符配平（忠实渲染）
                return f"{lmap} {main} \\right."
            if variation == 2:  # 仅右分隔符 → \left. 空分隔符配平
                return f"\\left. {main} {rmap}"
            return f"{lmap} {main} {rmap}"
        if sel == TM_UBAR:
            return f"\\underline{{{req(0, 'underline')}}}"
        if sel == TM_OBAR:
            return f"\\overline{{{req(0, 'overline')}}}"
        if sel in (TM_LARROW, TM_RARROW, TM_BARROW, TM_UARROW, TM_OARROW,
                   TM_UHBRACE, TM_LHBRACE, TM_LDIV, TM_OARC):
            return s(0)
        if sel == TM_DIRAC:
            return f"{req(0, 'dirac')} \\mid {req(1, 'dirac')}"
        self.unsupported.append(f"template:{_SELECTOR_NAMES.get(sel, sel)}")
        return "".join(slots)

    def _skip_font(self) -> None:
        """FONT 记录：tface + style + name。"""
        if self.pos + 1 < len(self.data):
            self.pos += 2
        self._null_string()


def parse_mtef(data: bytes) -> dict:
    """入口：MTEF 二进制 → {latex, errors, status, template_count, has_cjk, version}。"""
    return MTEFParser(data).parse()
