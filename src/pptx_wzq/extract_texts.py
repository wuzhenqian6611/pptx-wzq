"""extract_texts.py — 逐页提取 PPT 文本框/表格文本，输出 md 清单。

v1.0 设计：
  - 遍历每页 slide 的 <p:sp>（含 grpSp 嵌套）与 <p:graphicFrame>（表格），
    每个文本对象一行；
  - 每页一个 ID 序列：TXT{page}-{seq}（seq 页内递增）；
  - 标题/内容标记：占位符 ph type ∈ {title, ctrTitle}（或形状名含
    title/标题/Title）→ 标记 [标题]，其余 → [内容]；
  - 过滤（默认开启，--no-filter 关闭）：
      1) 页眉/页脚/页码/日期：ph type ∈ {ftr, sldNum, hdr, dt}；
      2) 母版/布局固定文本：与 slideLayout/slideMaster 中形状文本完全
         相同的条目（如"单击此处编辑母版标题样式"、课程标题、单位作者）；
      3) 跨页全局文本：在 ≥90% 页面重复出现的相同文本（每页手动复制的
         课程标题/高校/出版社/作者信息）；
      4) 空/纯空白、长度 < min_len（默认 2）的碎片（如"."、"§"）。
"""
from __future__ import annotations

import json
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from pathlib import Path

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
SP = "{http://schemas.openxmlformats.org/presentationml/2006/main}sp"
GRAPHICFRAME = "{http://schemas.openxmlformats.org/presentationml/2006/main}graphicFrame"
TABLECELL = "{http://schemas.openxmlformats.org/drawingml/2006/main}tc"
TITLE = "{http://schemas.openxmlformats.org/drawingml/2006/main}t"

# 页眉/页脚/页码/日期占位符类型（不进入正文）
HEADER_FOOTER_PH = {"ftr", "sldNum", "hdr", "dt"}
# 标题占位符类型
TITLE_PH = {"title", "ctrTitle"}


def _q(ns, tag):
    return f"{{{NS[ns]}}}{tag}"


def _sp_text(sp) -> str:
    """提取 sp 内 txBody 全部文本（段落用 | 连接，跑马灯 fld 也取文本）。"""
    body = sp.find(_q("p", "txBody"))
    if body is None:
        return ""
    parts = []
    for para in body.iter(_q("a", "p")):
        seg = []
        for t in para.iter(_q("a", "t")):
            if t.text:
                seg.append(t.text)
        for f in para.iter(_q("a", "fld")):
            t = f.find(_q("a", "t"))
            if t is not None and t.text:
                seg.append(t.text)
        parts.append("".join(seg))
    return " | ".join(x.strip() for x in parts if x.strip())


def _sp_meta(sp):
    """返回 (shape_name, ph_type, x, y, w, h)。
    坐标为形状在幻灯片上的位置与尺寸（EMU→px，来自 <a:xfrm> off/ext）。"""
    name = ""
    for cnvpr in sp.iter(_q("p", "cNvPr")):
        name = cnvpr.get("name", "")
        break
    ph_type = ""
    for ph in sp.iter(_q("p", "ph")):
        ph_type = ph.get("type", "obj")
        break
    x = y = w = h = 0
    try:
        xfrm = sp.find(".//" + _q("a", "xfrm"))
        if xfrm is not None:
            off = xfrm.find(_q("a", "off"))
            ext = xfrm.find(_q("a", "ext"))
            if off is not None:
                x = int(off.get("x", 0)) / 914400 * 96
                y = int(off.get("y", 0)) / 914400 * 96
            if ext is not None:
                w = int(ext.get("cx", 0)) / 914400 * 96
                h = int(ext.get("cy", 0)) / 914400 * 96
    except Exception:
        pass
    return name, ph_type, x, y, w, h


def _table_text(gf) -> str:
    """graphicFrame 内表格文本：每行单元格以 | 连接，行之间用 / 连接。"""
    rows = []
    for tr in gf.iter(_q("a", "tr")):
        cells = []
        for tc in tr.findall(TABLECELL):
            parts = []
            for t in tc.iter(TITLE):
                if t.text:
                    parts.append(t.text)
            cells.append("".join(parts).strip())
        rows.append(" | ".join(x for x in cells if x))
    return " / ".join(x for x in rows if x)


def _iter_slide_texts(slide_xml: bytes, with_tables: bool = True):
    """遍历 slide 全部文本对象，产出 (kind, name, ph_type, text, x, y, w, h)。"""
    root = ET.fromstring(slide_xml)
    for sp in root.iter(SP):
        txt = _sp_text(sp)
        if not txt:
            continue
        name, ph_type, x, y, w, h = _sp_meta(sp)
        yield ("sp", name, ph_type, txt, x, y, w, h)
    if with_tables:
        for gf in root.iter(GRAPHICFRAME):
            txt = _table_text(gf)
            if txt:
                name = ""
                for cnvpr in gf.iter(_q("p", "cNvPr")):
                    name = cnvpr.get("name", "")
                    break
                x = y = w = h = 0
                try:
                    xfrm = gf.find(".//" + q("a", "xfrm"))
                    if xfrm is not None:
                        off = xfrm.find(q("a", "off"))
                        ext = xfrm.find(q("a", "ext"))
                        if off is not None:
                            x = int(off.get("x", 0)) / 914400 * 96
                            y = int(off.get("y", 0)) / 914400 * 96
                        if ext is not None:
                            w = int(ext.get("cx", 0)) / 914400 * 96
                            h = int(ext.get("cy", 0)) / 914400 * 96
                except Exception:
                    pass
                yield ("table", name or "表格", "tbl", txt, x, y, w, h)


def _fixed_texts(zf: zipfile.ZipFile) -> set:
    """收集母版/布局里所有形状的文本（每页共享的固定文本，须排除）。"""
    out = set()
    for n in zf.namelist():
        if not (n.startswith("ppt/slideMasters/slideMaster")
                or n.startswith("ppt/slideLayouts/slideLayout")):
            continue
        try:
            root = ET.fromstring(zf.read(n))
        except Exception:
            continue
        for sp in root.iter(SP):
            txt = _sp_text(sp).strip()
            if len(txt) >= 2:
                out.add(txt)
    return out


def extract_texts(pptx_path, out_dir,
                  filter_texts: bool = True, min_len: int = 2,
                  with_tables: bool = True, on_progress=None):
    """逐页提取文本，写 <stem>_texts.md（+ <stem>_text_entries.json 审计）。

    返回统计 dict：pages / total / titles / filtered / excluded / md。
    filter_texts=False 时全量提取（含页眉页脚、母版固定文本），协议不变。
    """
    pptx_path = Path(pptx_path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = pptx_path.stem

    slides_in_order = []
    with zipfile.ZipFile(pptx_path) as zf:
        slides_in_order = sorted(
            (n for n in zf.namelist()
             if re.match(r"ppt/slides/slide\d+\.xml$", n)),
            key=lambda n: int(re.search(r"\d+", n).group()))
        fixed = _fixed_texts(zf) if filter_texts else set()
        n_slides = len(slides_in_order)

        # 第一遍：收集全部页文本（统计跨页重复用）
        all_pages_texts = []          # [{ph, text, ...}]
        for i, slide_path in enumerate(slides_in_order):
            if on_progress is not None:
                try:
                    on_progress(i + 1, n_slides, {"kind": "text"})
                except Exception:
                    pass
            items = []
            for kind, name, ph, txt, x, y, w, h in \
                    _iter_slide_texts(zf.read(slide_path), with_tables):
                items.append({"kind": kind, "name": name, "ph": ph,
                              "text": txt.strip(),
                              "x": x, "y": y, "w": w, "h": h})
            all_pages_texts.append(items)

        # 跨页重复统计（≥90% 页面相同 → 全局固定文本）
        text_pages = Counter()
        for items in all_pages_texts:
            seen = set()
            for it in items:
                if it["text"] not in seen:
                    seen.add(it["text"])
            for t in seen:
                text_pages[t] += 1
        global_texts = {t for t, c in text_pages.items()
                        if c / n_slides >= 0.9} if filter_texts else set()

        # 第二遍：过滤 + 写 md
        lines = [f"# {stem} 文本清单", "",
                 f"> 由 `pptx-text` 提取（共 {n_slides} 页）。"
                 "每个文本对象一行；ID 页内递增；已排除页眉/页脚/页码/"
                 "母版固定文本" + ("" if filter_texts else "（--no-filter 全量）"),
                 "> 格式：`ID | 类型 | 文本 | 坐标`；`[标题]` 来自标题占位符，"
                 "其余为 `[内容]`（含表格行）。坐标=(x,y) 宽x高（px，幻灯片坐标系）。", ""]
        n_title = n_total = n_filtered = 0
        entries = []
        for page_no, items in enumerate(all_pages_texts, start=1):
            lines.append(f"## 第 {page_no} 页")
            lines.append("")
            lines.append("| ID | 类型 | 文本 | 坐标 |")
            lines.append("|---|---|---|---|")
            kept_this = 0
            for it in items:
                txt = it["text"]
                reason = None
                if filter_texts:
                    if it["ph"] in HEADER_FOOTER_PH:
                        reason = f"页眉页脚/页码(ph={it['ph']})"
                    elif txt in fixed:
                        reason = "母版/布局固定文本"
                    elif txt in global_texts:
                        reason = "跨页全局固定文本"
                    elif len(txt) < min_len:
                        reason = f"过短(<{min_len}字符)"
                pos = f"({int(it['x'])},{int(it['y'])}) " \
                      f"{int(it['w'])}x{int(it['h'])}"
                if reason is not None:
                    n_filtered += 1
                    entries.append({"page": page_no, "id": None,
                                    "type": "排除", "text": txt,
                                    "reason": reason,
                                    "shape": it["name"], "ph": it["ph"],
                                    "x": it["x"], "y": it["y"],
                                    "w": it["w"], "h": it["h"]})
                    continue
                kept_this += 1
                n_total += 1
                is_title = it["ph"] in TITLE_PH or \
                    re.search(r"title|标题", it["name"], re.I)
                label = "标题" if is_title else "内容"
                if is_title:
                    n_title += 1
                txt_disp = txt.replace("|", "\\|")
                lines.append(f"| TXT{page_no:03d}-{kept_this:02d} | "
                             f"{label} | {txt_disp} | {pos} |")
                entries.append({"page": page_no,
                                "id": f"TXT{page_no:03d}-{kept_this:02d}",
                                "type": label, "text": txt,
                                "shape": it["name"], "ph": it["ph"],
                                "x": it["x"], "y": it["y"],
                                "w": it["w"], "h": it["h"]})
            lines.append("")

        lines.append("---")
        lines.append("")
        lines.append(f"文本对象总数 **{n_total}** 条"
                     f"（标题 {n_title} / 内容 {n_total - n_title}）")
        if filter_texts:
            lines.append(f"- 已排除：{n_filtered} 条"
                         "（页眉页脚/页码/母版固定/跨页全局/过短）")
            lines.append(f"- 页数：{n_slides}")
        lines.append("")
        md_path = out_dir / f"{stem}_texts.md"
        md_path.write_text("\n".join(lines), encoding="utf-8")
        (out_dir / f"{stem}_text_entries.json").write_text(
            json.dumps(entries, ensure_ascii=False, indent=1), encoding="utf-8")

    return {"pages": n_slides, "total": n_total, "titles": n_title,
            "contents": n_total - n_title,
            "filtered": n_filtered, "md": str(md_path)}
