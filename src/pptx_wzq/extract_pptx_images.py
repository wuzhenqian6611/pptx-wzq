#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract_pptx_images.py — 通用 PPTX 图片提取器（基于 OOXML 标准，v2）
============================================================================

在 v1（独立图片对象 / 页面背景 / 全部媒体）基础上，v2 新增并强化：

  A. 形状填充图（shape fill image）
     除 <p:pic> 之外的形状（矩形、圆、任意 <p:sp>）若用「图片填充」
     （<p:spPr><a:blipFill>），同样会被提取为独立 PNG，kind='fill'。

  B. 布局 / 母版背景图（slide → layout → master 三级回退）
     <p:bg> 可能不在幻灯片本体，而在其 slideLayout 甚至 slideMaster 上。
     本工具沿 sldLayoutId → layout → master 链路追踪背景图，缺省开启。

  C. 矢量图栅格化（EMF / WMF / SVG → PNG）
     Pillow 无法解码矢量图，故 v1 只保留原文件。本版在 --rasterize-vector
     开启时，自动探测 LibreOffice(soffice) 或 Inkscape 把矢量图转成 PNG；
     两者都缺失则保留矢量原文件并在 note 中标注，绝不报错中断。

  D. 【v2 核心】srcRect 裁剪 —— 还原「用户直观所见」
     PPT 里一个 <p:pic> 可能引用「多图合图」底图，靠 <a:srcRect> 只显示其中
     一段。本工具读取 srcRect 并沿它裁剪底图 PNG，使导出图与 PPT 页面显示一致，
     而非把合图里「本页不显示的其他内容」一并带出。可用 --no-crop 退回旧行为。

  E. 【v2 新增】公式对象（<p:oleObj> / OMML）与图表（<p:graphicFrame><c:chart>）
     教材核心内容。无可用渲染工具时，按源文件（oleObjectN.bin / OMML / chartN.xml）
     目录化保留并标注；探测到 LibreOffice 时，整页渲染 PNG 再按 xfrm 裁出对象区域。

  F. 【v2 新增】入参路径跨平台归一化
     同一提取器会被不同外壳 / 智能体以不同风格传参（Windows `C:\\...`、
     Git Bash `/c/...`、WSL `/mnt/c/...`、UNC、`~` 家目录），统一归一化为当前平台
     可识别的绝对路径，不写死任何盘符或特例。

用法
----
  python extract_pptx_images.py input.pptx
  python extract_pptx_images.py input.pptx -o out --all-media
  python extract_pptx_images.py input.pptx --no-fill          # 不提取形状填充图
  python extract_pptx_images.py input.pptx --no-bg-layout     # 不追溯布局/母版背景
  python extract_pptx_images.py input.pptx --rasterize-vector --raster-dpi 200
  python extract_pptx_images.py input.pptx --no-crop          # 退回导出完整媒体
  python extract_pptx_images.py input.pptx --min-crop 48      # 裁剪阈值(px)
  python extract_pptx_images.py input.pptx --latex            # [v4] 额外导出 <名>_formulas.md（三路径公式 LaTeX 汇总）
  python extract_pptx_images.py input.pptx --latex --no-eq3   # [v4] 关闭 EQ3 解析路径
  python extract_pptx_images.py input.pptx --latex --no-ocr   # [v4] 关闭数学 OCR（仍渲染+裁剪）

作者：吴振谦 · wuzhenqian@nbu.edu.cn · QQ：38328063"""
from __future__ import annotations

import argparse
import csv
import json
import os
import posixpath
import re
import shutil
import subprocess
import sys
import zipfile
from collections import Counter
from dataclasses import dataclass, asdict
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET

try:
    from PIL import Image
    HAVE_PIL = True
except Exception:  # pragma: no cover
    HAVE_PIL = False


# --------------------------------------------------------------------------
# OOXML 命名空间
# --------------------------------------------------------------------------
NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}

REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
REL_LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
REL_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
REL_OLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
REL_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"


def q(pre: str, tag: str) -> str:
    return f"{{{NS[pre]}}}{tag}"


PIC = q("p", "pic")
SP = q("p", "sp")
CXNSP = q("p", "cxnSp")
BLIP = q("a", "blip")
CNVPR = q("p", "cNvPr")
BG = q("p", "bg")
OLEOBJ = q("p", "oleObj")
GRAPHICFRAME = q("p", "graphicFrame")


def _sp_text(sp) -> str:
    """提取 <p:sp>/<p:cxnSp> 内全部 <a:t> 文本（含嵌套段落）。"""
    parts = []
    for t in sp.iter(q("a", "t")):
        if t.text:
            parts.append(t.text)
    return "".join(parts).strip()


# --------------------------------------------------------------------------
# 入参路径跨平台归一化（不写死盘符 / 外壳特例）
# --------------------------------------------------------------------------
def normalize_path(p: str) -> str:
    """把不同外壳 / 平台下传入的路径，归一化为当前平台可识别的绝对路径。

    处理：~ 家目录、环境变量、Git Bash 的 /x/...、WSL 的 /mnt/x/...、
    Windows 原生 C:\\...、UNC \\\\host\\...。不假设具体盘符，用变量推导。
    """
    if p is None:
        return p
    p = os.path.expanduser(os.path.expandvars(p))
    if sys.platform.startswith("win"):
        # POSIX 风格（Git Bash / Cygwin）：/x/foo -> X:\\foo
        m = re.match(r"^/([a-zA-Z])/(.*)$", p)
        if m:
            rest = m.group(2).replace('/', '\\')
            return f"{m.group(1).upper()}:\\{rest}"
        # WSL 风格：/mnt/x/foo -> X:\\foo
        m = re.match(r"^/mnt/([a-zA-Z])/(.*)$", p)
        if m:
            rest = m.group(2).replace('/', '\\')
            return f"{m.group(1).upper()}:\\{rest}"
        # 其余（UNC / 原生 Windows 路径 / 无法映射的 POSIX 绝对路径）原样 normpath
    return os.path.normpath(p)


# --------------------------------------------------------------------------
# 数据载体
# --------------------------------------------------------------------------
@dataclass
class ImageRecord:
    page: int
    index: int
    kind: str               # picture | fill | background | formula_ole | formula_omath | chart | visio
    shape_name: str
    source_media: str
    output_file: str
    width: int = 0
    height: int = 0
    original_format: str = ""
    converted_to_png: bool = False
    note: str = ""
    src_rect: tuple = None  # (l,t,r,b) 比例 0..1，来自 <a:srcRect>
    cropped: str = ""       # "" | "yes" | "no" | "full"(裁出过小退化)
    x: float = 0            # 在幻灯片中的位置（px，来自 <a:xfrm><a:off> EMU）
    y: float = 0
    shape_w: float = 0      # 在幻灯片上的显示宽高（px，来自 <a:xfrm><a:ext> EMU）
    shape_h: float = 0
    ole_progid: str = ""    # OLE 对象 progId（识别 Visio/公式编辑器等）
    preview_file: str = ""  # visio 对象关联的预览图文件名（EMF/SVG，可空）


# --------------------------------------------------------------------------
# 关系 / 路径解析
# --------------------------------------------------------------------------
def read_rels(zf: zipfile.ZipFile, rels_path: str) -> dict:
    """返回 {Id: Target}。"""
    if rels_path not in zf.namelist():
        return {}
    root = ET.fromstring(zf.read(rels_path))
    out = {}
    for rel in root:
        rid = rel.get("Id")
        target = rel.get("Target")
        if rid and target:
            out[rid] = target
    return out


def read_rels_full(zf: zipfile.ZipFile, rels_path: str) -> dict:
    """返回 {Id: {'Target':..., 'Type':...}}，用于识别 layout/master 关系。"""
    if rels_path not in zf.namelist():
        return {}
    root = ET.fromstring(zf.read(rels_path))
    out = {}
    for rel in root:
        rid = rel.get("Id")
        if rid:
            out[rid] = {"Target": rel.get("Target"), "Type": rel.get("Type")}
    return out


def resolve_target(base_dir: str, target: str) -> str:
    if target.startswith("/"):
        # 包内绝对路径（以 / 开头，指向包根）
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(base_dir, target))


def list_slides_in_order(zf: zipfile.ZipFile) -> list:
    """按 presentation.xml 的 <p:sldIdLst> 顺序返回 slide 路径列表（还原真实页码）。"""
    pres_path = "ppt/presentation.xml"
    if pres_path not in zf.namelist():
        return sorted(n for n in zf.namelist()
                      if n.startswith("ppt/slides/slide") and n.endswith(".xml"))

    root = ET.fromstring(zf.read(pres_path))
    sld_id_lst = root.find(q("p", "sldIdLst"))
    if sld_id_lst is None:
        return sorted(n for n in zf.namelist()
                      if n.startswith("ppt/slides/slide") and n.endswith(".xml"))

    pres_rels = read_rels(zf, "ppt/_rels/presentation.xml.rels")
    slides = []
    for sld_id in sld_id_lst.findall(q("p", "sldId")):
        rid = sld_id.get(q("r", "id"))
        target = pres_rels.get(rid)
        if not target:
            continue
        path = resolve_target("ppt", target)
        if path in zf.namelist():
            slides.append(path)
    return slides


# --------------------------------------------------------------------------
# srcRect 解析
# --------------------------------------------------------------------------
def parse_src_rect(blip_fill) -> tuple | None:
    """从 <p:blipFill> 下读取 <a:srcRect>，返回 (l,t,r,b) 比例(0..1) 或 None。
    OOXML 单位：1/100000，仅被指定的边才有属性，缺省=0。"""
    sr = blip_fill.find(q("a", "srcRect"))
    if sr is None:
        return None
    f = lambda k: (int(sr.get(k)) / 100000.0) if sr.get(k) is not None else 0.0
    return (f("l"), f("t"), f("r"), f("b"))


def _find_xfrm(elem):
    """从形状/图框元素里取 <a:off>/<a:ext>（EMU）。返回 (ox,oy,cx,cy) 或 None。"""
    off = elem.find(".//" + q("a", "off"))
    ext = elem.find(".//" + q("a", "ext"))
    if off is not None and ext is not None:
        return (off.get("x"), off.get("y"), ext.get("cx"), ext.get("cy"))
    return None


# --------------------------------------------------------------------------
# 形状遍历：独立图片对象 / 形状填充图 / 背景链
# --------------------------------------------------------------------------
def _xfrm_xy(el) -> tuple:
    """从形状元素解析 <a:xfrm><a:off> 左上角位置（EMU→px）。"""
    try:
        xfrm = el.find(".//" + q("a", "xfrm"))
        if xfrm is None:
            return 0.0, 0.0
        off = xfrm.find(q("a", "off"))
        if off is None:
            return 0.0, 0.0
        return (int(off.get("x", 0)) / 914400 * 96,
                int(off.get("y", 0)) / 914400 * 96)
    except Exception:
        return 0.0, 0.0


def _xfrm_wh(el) -> tuple:
    """从形状元素解析 <a:xfrm><a:ext> 显示宽高（EMU→px）。"""
    try:
        xfrm = el.find(".//" + q("a", "xfrm"))
        if xfrm is None:
            return 0.0, 0.0
        ext = xfrm.find(q("a", "ext"))
        if ext is None:
            return 0.0, 0.0
        return (int(ext.get("cx", 0)) / 914400 * 96,
                int(ext.get("cy", 0)) / 914400 * 96)
    except Exception:
        return 0.0, 0.0


def iter_pictures(slide_xml: bytes):
    """独立图片对象 <p:pic> → (shape_name, embed_rid, src_rect, x, y, w, h)。"""
    root = ET.fromstring(slide_xml)
    for pic in root.iter(PIC):
        name = ""
        for cnvpr in pic.iter(CNVPR):
            name = cnvpr.get("name", "")
            break
        blip_fill = pic.find(q("p", "blipFill"))
        if blip_fill is None:
            continue
        blip = blip_fill.find(BLIP)
        rid = None
        if blip is not None:
            rid = blip.get(q("r", "embed")) or blip.get(q("r", "link"))
        src_rect = parse_src_rect(blip_fill)
        x, y = _xfrm_xy(pic)
        w, h = _xfrm_wh(pic)
        yield name, rid, src_rect, x, y, w, h


def iter_fill_images(slide_xml: bytes):
    """形状填充图：<p:sp> 内部 <p:spPr><a:blipFill> 且非 <p:pic>。
    返回 (shape_name, embed_rid, src_rect, x, y, w, h)。"""
    root = ET.fromstring(slide_xml)
    for sp in root.iter(SP):
        sp_pr = sp.find(q("p", "spPr"))
        if sp_pr is None:
            continue
        blip_fill = sp_pr.find(q("a", "blipFill"))
        if blip_fill is None:
            continue
        blip = blip_fill.find(BLIP)
        if blip is None:                       # 极少数 blip 更深一层
            blip = blip_fill.find(".//" + BLIP)
        if blip is None:
            continue
        rid = blip.get(q("r", "embed")) or blip.get(q("r", "link"))
        if not rid:
            continue
        name = ""
        for cnvpr in sp.iter(CNVPR):
            name = cnvpr.get("name", "")
            break
        src_rect = parse_src_rect(blip_fill)
        x, y = _xfrm_xy(sp)
        w, h = _xfrm_wh(sp)
        yield name or "FillShape", rid, src_rect, x, y, w, h


def _bg_blips(xml_bytes: bytes):
    """从任意 xml（slide/layout/master）里取 <p:bg> 内的 (rid, src_rect)。"""
    root = ET.fromstring(xml_bytes)
    bg = root.find(".//" + BG)
    if bg is None:
        return
    for blip_fill in bg.iter(q("a", "blipFill")):
        rid = None
        blip = blip_fill.find(BLIP)
        if blip is not None:
            rid = blip.get(q("r", "embed")) or blip.get(q("r", "link"))
        if rid:
            yield rid, parse_src_rect(blip_fill)


def iter_background_chain(zf: zipfile.ZipFile, slide_path: str, slide_rels_full: dict):
    """沿 幻灯片 → 布局 → 母版 追溯背景图，产出 (name, rid, base_dir, rels, src_rect)。
    按出现优先级返回，调用方负责去重。"""
    seen = set()
    slide_dir = posixpath.join("", posixpath.dirname(slide_path))
    slide_rels = {rid: v["Target"] for rid, v in slide_rels_full.items()}
    for rid, src_rect in _bg_blips(zf.read(slide_path)):
        key = (slide_dir, rid)
        if key in seen:
            continue
        seen.add(key)
        yield "SlideBackground", rid, slide_dir, slide_rels, src_rect

    layout_target = next((v["Target"] for v in slide_rels_full.values()
                          if v["Type"] == REL_LAYOUT), None)
    if layout_target:
        layout_path = resolve_target(slide_dir, layout_target)
        if layout_path in zf.namelist():
            layout_dir = posixpath.dirname(layout_path)
            layout_rels_full = read_rels_full(
                zf, posixpath.join(layout_dir, "_rels",
                                   posixpath.basename(layout_path) + ".rels"))
            layout_rels = {rid: v["Target"] for rid, v in layout_rels_full.items()}
            for rid, src_rect in _bg_blips(zf.read(layout_path)):
                key = (layout_dir, rid)
                if key in seen:
                    continue
                seen.add(key)
                yield "LayoutBackground", rid, layout_dir, layout_rels, src_rect

            master_target = next((v["Target"] for v in layout_rels_full.values()
                                  if v["Type"] == REL_MASTER), None)
            if master_target:
                master_path = resolve_target(layout_dir, master_target)
                if master_path in zf.namelist():
                    master_dir = posixpath.dirname(master_path)
                    master_rels = read_rels(
                        zf, posixpath.join(master_dir, "_rels",
                                           posixpath.basename(master_path) + ".rels"))
                    for rid, src_rect in _bg_blips(zf.read(master_path)):
                        key = (master_dir, rid)
                        if key in seen:
                            continue
                        seen.add(key)
                        yield "MasterBackground", rid, master_dir, master_rels, src_rect


# --------------------------------------------------------------------------
# 公式对象（oleObj / OMML）与图表（graphicFrame）遍历
# --------------------------------------------------------------------------
def iter_ole_formulas(slide_xml: bytes):
    """扫描 OLE 对象（公式编辑器 / 嵌入对象），返回 (shape_name, embed_rid, xfrm, prog_id)。
    真实课件中 OLE 常位于 <p:graphicFrame> 内（少数在 <p:sp>），且 r:id 直接挂在
    <p:oleObj> 上（不一定有嵌套 <p:oleObjEmbed>）。progId 含 'equation' 即公式编辑器。"""
    root = ET.fromstring(slide_xml)
    for parent in (GRAPHICFRAME, SP):
        for el in root.iter(parent):
            ole = el.find(".//" + OLEOBJ)
            if ole is None:
                continue
            rid = ole.get(q("r", "id"))          # r:id 直接挂在 <p:oleObj>
            if not rid:
                embed = ole.find(q("p", "oleObjEmbed"))
                rid = embed.get(q("r", "id")) if embed is not None else None
            if not rid:
                continue
            prog_id = (ole.get("progId") or "").lower()
            name = ""
            for cnvpr in el.iter(CNVPR):
                name = cnvpr.get("name", "")
                break
            if not name:
                name = "Equation" if "equation" in prog_id else "OLEObject"
            xfrm = _find_xfrm(el)
            yield name, rid, xfrm, prog_id


def iter_omath(slide_xml: bytes):
    """扫描 slide 内所有 <m:oMath>，返回 OMML 片段字符串列表（每个独立公式）。"""
    root = ET.fromstring(slide_xml)
    out = []
    for om in root.iter(q("m", "oMath")):
        out.append(ET.tostring(om, encoding="unicode"))
    return out


def iter_charts(slide_xml: bytes):
    """扫描 <p:graphicFrame> 内含 <c:chart r:id>。返回 (shape_name, chart_rid, xfrm)。"""
    root = ET.fromstring(slide_xml)
    for gf in root.iter(GRAPHICFRAME):
        chart = gf.find(".//" + q("c", "chart"))
        if chart is None:
            continue
        rid = chart.get(q("r", "id"))
        if not rid:
            continue
        name = ""
        for cnvpr in gf.iter(CNVPR):
            name = cnvpr.get("name", "")
            break
        xfrm = _find_xfrm(gf)
        yield name or "Chart", rid, xfrm


# --------------------------------------------------------------------------
# 原子对象：原生 shape / 连接符 / 表格 / 文本框（可视逻辑块的数据地基）
# --------------------------------------------------------------------------
def iter_native_shapes(slide_xml: bytes):
    """扫描全部 <p:sp>（含文本框/形状/图形），返回 (shape_name, text, xfrm, z_index, ph_type)。
    用于可视逻辑块的空间聚类；注意过滤掉已被 <p:pic>/<p:oleObj>/<p:graphicFrame> 覆盖的对象。
    不含图片填充的纯形状（如箭头、矩形、圆）也会输出。ph_type 用于识别标题/正文占位符
    （title/ctrTitle/subTitle/body 等），这些是页面文本区而非可视逻辑块。"""
    root = ET.fromstring(slide_xml)
    for idx, sp in enumerate(root.iter(SP)):
        # 跳过已被独立图片对象处理的形状（<p:pic> 内部也有 <p:sp>）
        if sp.find(q("p", "pic")) is not None:
            continue
        name = ""
        for cnvpr in sp.iter(CNVPR):
            name = cnvpr.get("name", "")
            break
        text = _sp_text(sp)
        xfrm = _find_xfrm(sp)
        if xfrm is None:
            continue
        ph_type = ""
        for ph in sp.iter(q("p", "ph")):
            ph_type = ph.get("type", "obj")
            break
        yield name or f"Shape{idx}", text, xfrm, idx, ph_type


def iter_connectors(slide_xml: bytes):
    """扫描 <p:cxnSp>（连接符/箭头线）。返回 (shape_name, text, xfrm, z_index, start_id, end_id)。
    start_id/end_id 为 <a:stCxn>/<a:endCxn> 的 id（指向被连接 shape 的 id），可能为空。"""
    root = ET.fromstring(slide_xml)
    for idx, cxn in enumerate(root.iter(CXNSP)):
        name = ""
        for cnvpr in cxn.iter(CNVPR):
            name = cnvpr.get("name", "")
            break
        text = _sp_text(cxn)
        xfrm = _find_xfrm(cxn)
        start_id = end_id = ""
        for c in cxn.iter(q("a", "stCxn")):
            start_id = c.get("id", "")
        for c in cxn.iter(q("a", "endCxn")):
            end_id = c.get("id", "")
        if xfrm is None:
            continue
        yield name or f"Connector{idx}", text, xfrm, idx, start_id, end_id


def iter_tables(slide_xml: bytes):
    """扫描 <p:graphicFrame> 内含 <a:tbl> 的表格对象。返回 (shape_name, rows, xfrm, z_index)。
    rows = [[cell_text, ...], ...]（按行组织）。"""
    root = ET.fromstring(slide_xml)
    for idx, gf in enumerate(root.iter(GRAPHICFRAME)):
        tbl = gf.find(".//" + q("a", "tbl"))
        if tbl is None:
            continue
        name = ""
        for cnvpr in gf.iter(CNVPR):
            name = cnvpr.get("name", "")
            break
        rows = []
        for tr in tbl.iter(q("a", "tr")):
            row = []
            for tc in tr.iter(q("a", "tc")):
                cells = [t.text or "" for t in tc.iter(q("a", "t"))]
                row.append("".join(cells).strip())
            rows.append(row)
        xfrm = _find_xfrm(gf)
        yield name or f"Table{idx}", rows, xfrm, idx


def _xfrm_to_bbox(xfrm) -> dict:
    """把 (ox, oy, cx, cy) EMU 元组转成 px bbox dict；异常返回空 bbox。"""
    try:
        if xfrm is None:
            return {"x": 0, "y": 0, "w": 0, "h": 0}
        return {"x": int(xfrm[0]) / 914400 * 96,
                "y": int(xfrm[1]) / 914400 * 96,
                "w": int(xfrm[2]) / 914400 * 96,
                "h": int(xfrm[3]) / 914400 * 96}
    except Exception:
        return {"x": 0, "y": 0, "w": 0, "h": 0}


def _collect_atomic_objects(page_no: int, slide_xml: bytes,
                            page_records: list) -> list:
    """把本页的图片记录（ImageRecord）与 shape/connector/table 统一映射为
    原子对象 dict 列表（按 z_index 排序）。图片记录缺 z_index 时排在 shape 之后。"""
    objs = []
    # 图片类记录（已有 x/y/shape_w/h）
    for rec in page_records:
        if not rec.output_file:
            continue
        kind = rec.kind
        if kind == "picture":
            fmt = (rec.original_format or "").lower()
            a_kind = "vector" if fmt in VECTOR else "raster"
            if fmt in ("vsdx", "vsd"):
                a_kind = "visio"
            bbox = {"x": rec.x, "y": rec.y, "w": rec.shape_w, "h": rec.shape_h}
            objs.append({
                "obj_id": f"s{page_no:02d}_r{rec.index:02d}",
                "page": page_no, "kind": a_kind,
                "shape_name": rec.shape_name or "",
                "text": "", "bbox": bbox, "z_index": 1000 + rec.index,
                "source_media": rec.source_media or "",
                "output_file": rec.output_file or "",
                "original_format": fmt,
                "children": [],
            })
        elif kind == "chart":
            bbox = {"x": rec.x, "y": rec.y, "w": rec.shape_w, "h": rec.shape_h}
            objs.append({
                "obj_id": f"s{page_no:02d}_c{rec.index:02d}",
                "page": page_no, "kind": "chart",
                "shape_name": rec.shape_name or "",
                "text": "", "bbox": bbox, "z_index": 1000 + rec.index,
                "source_media": rec.source_media or "",
                "output_file": rec.output_file or "",
                "original_format": rec.original_format or "",
                "children": [],
            })
        elif kind == "formula_ole":
            bbox = {"x": rec.x, "y": rec.y, "w": rec.shape_w, "h": rec.shape_h}
            objs.append({
                "obj_id": f"s{page_no:02d}_f{rec.index:02d}",
                "page": page_no, "kind": "formula",
                "shape_name": rec.shape_name or "",
                "text": "", "bbox": bbox, "z_index": 1000 + rec.index,
                "source_media": rec.source_media or "",
                "output_file": rec.output_file or "",
                "original_format": rec.original_format or "",
                "children": [],
            })
        elif kind == "formula_omath":
            bbox = {"x": rec.x, "y": rec.y, "w": rec.shape_w, "h": rec.shape_h}
            objs.append({
                "obj_id": f"s{page_no:02d}_m{rec.index:02d}",
                "page": page_no, "kind": "formula",
                "shape_name": rec.shape_name or "",
                "text": "", "bbox": bbox, "z_index": 1000 + rec.index,
                "source_media": rec.source_media or "",
                "output_file": rec.output_file or "",
                "original_format": rec.original_format or "",
                "children": [],
            })
    # 原生 shape（文本框/图形）
    # 标题/正文/副标题等占位符 → kind="text_region"（页面文本区，不参与
    # 可视逻辑块聚类；它们的内容已由文本提取步骤完整保留）
    TEXT_REGION_PH = {"title", "ctrTitle", "subTitle", "body", "obj"}
    for name, text, xfrm, z, ph_type in iter_native_shapes(slide_xml):
        if not text and not name:
            continue
        kind = "shape"
        if ph_type in TEXT_REGION_PH:
            kind = "text_region"
        objs.append({
            "obj_id": f"s{page_no:02d}_sp{z:03d}",
            "page": page_no, "kind": kind,
            "shape_name": name, "text": text,
            "bbox": _xfrm_to_bbox(xfrm), "z_index": z,
            "source_media": "", "output_file": "",
            "original_format": "", "children": [],
            "ph_type": ph_type,
        })
    # 连接符（箭头线）
    for name, text, xfrm, z, s_id, e_id in iter_connectors(slide_xml):
        objs.append({
            "obj_id": f"s{page_no:02d}_cn{z:03d}",
            "page": page_no, "kind": "connector",
            "shape_name": name, "text": text or "",
            "bbox": _xfrm_to_bbox(xfrm), "z_index": z,
            "source_media": "", "output_file": "",
            "original_format": "", "children": [],
            "start_shape_id": s_id, "end_shape_id": e_id,
        })
    # 表格
    for name, rows, xfrm, z in iter_tables(slide_xml):
        objs.append({
            "obj_id": f"s{page_no:02d}_tb{z:03d}",
            "page": page_no, "kind": "table",
            "shape_name": name, "text": "",
            "bbox": _xfrm_to_bbox(xfrm), "z_index": z,
            "source_media": "", "output_file": "",
            "original_format": "", "children": rows,
        })
    objs.sort(key=lambda o: o["z_index"])
    return objs


# --------------------------------------------------------------------------
# 字节落盘（栅格→PNG 归一化；矢量保留原文件；支持 srcRect 裁剪）
# --------------------------------------------------------------------------
RASTER = {"jpg", "jpeg", "gif", "bmp", "tif", "tiff", "webp"}
VECTOR = {"emf", "wmf", "svg"}


def ext_of(name: str) -> str:
    return Path(name).suffix.lower().lstrip(".")


def _crop_im(im, src_rect, min_crop):
    """按 src_rect 比例裁剪 PIL 图像。返回 (im, flag)，flag: 'yes'|'full'。"""
    l, t, r, b = src_rect
    W, H = im.size
    x0, y0 = int(round(l * W)), int(round(t * H))
    x1, y1 = int(round((1 - r) * W)), int(round((1 - b) * H))
    if x1 > x0 and y1 > y0 and (x1 - x0) >= min_crop and (y1 - y0) >= min_crop:
        return im.crop((x0, y0, x1, y1)), "yes"
    return im, "full"  # 裁出过小，保存完整图


def save_image(data: bytes, out_path: Path, original_name: str, convert: bool,
               src_rect=None, min_crop: int = 64):
    """写成输出文件，返回 (written_path, converted, w, h, note, cropped)。
    - PNG/JPEG/... 栅格：按 src_rect 裁剪（若适用），栅格图统一转 PNG。
    - 矢量（emf/wmf/svg）：Pillow 无法解码，保留原文件，cropped=''。
    - 无 Pillow 或 convert=False：原样落盘，cropped=''。
    """
    ext = ext_of(original_name)
    note = ""
    cropped = ""

    if ext == "png":
        out = out_path.with_suffix(".png")
        if HAVE_PIL and src_rect is not None:
            try:
                im = Image.open(BytesIO(data))
                im.load()
                im, flag = _crop_im(im, src_rect, min_crop)
                if flag == "yes":
                    im.save(out, "PNG")
                    return out, True, im.width, im.height, "", "yes"
                # flag == 'full'：裁出过小，保存完整图
                out.write_bytes(data)
                return out, True, 0, 0, "srcRect 裁出过小，保存完整图", "full"
            except Exception as e:  # pragma: no cover
                note = f"Pillow 处理失败({e})，已保留原格式"
                out.write_bytes(data)
                return out, True, 0, 0, note, ""
        out.write_bytes(data)
        w = h = 0
        if HAVE_PIL:
            try:
                im = Image.open(BytesIO(data))
                w, h = im.size
            except Exception:
                pass
        return out, True, w, h, note, ""

    if convert and ext in RASTER and HAVE_PIL:
        try:
            im = Image.open(BytesIO(data))
            if src_rect is not None:
                im, flag = _crop_im(im, src_rect, min_crop)
                cropped = flag  # 'yes' | 'full'
            im = im.convert("RGBA")
            out = out_path.with_suffix(".png")
            im.save(out, "PNG")
            return out, True, im.width, im.height, note, cropped
        except Exception as e:  # pragma: no cover
            note = f"Pillow 转换失败({e})，已保留原格式"
            out = out_path.with_suffix("." + ext)
            out.write_bytes(data)
            return out, False, 0, 0, note, ""

    if ext in VECTOR:
        note = "矢量格式(Pillow 无法栅格化)，保留原文件（建议 --rasterize-vector）"
    out = out_path.with_suffix("." + (ext or "bin"))
    out.write_bytes(data)
    return out, False, 0, 0, note, cropped


# --------------------------------------------------------------------------
# 矢量 / 整页 栅格化（LibreOffice / Inkscape 自动探测）
# --------------------------------------------------------------------------
def _detect_rasterizer(prefer: str):
    """返回 ('soffice', exe) / ('inkscape', exe) / ('pdftoppm', exe) / None。

    v4 加法：在 PATH 探测之外，**额外**探测 LibreOffice 标准安装目录
    （Windows 用 %ProgramFiles% 环境变量推导，不写死盘符；macOS/linux 用
    常规路径）。解决「装了但不在 PATH 上」导致路径3 渲染前置缺失的问题。
    不改变既有 PATH 探测行为。
    """
    found = []
    if prefer in ("auto", "soffice", "libreoffice"):
        # 1) PATH（既有行为）
        for c in ("soffice", "libreoffice", "libreoffice24.2", "libreoffice7.6"):
            p = shutil.which(c)
            if p:
                found.append(("soffice", p))
                break
        # 2) 标准安装目录（加法，不写死盘符）
        if not any(t[0] == "soffice" for t in found):
            cand_dirs = []
            if sys.platform.startswith("win"):
                pf = os.environ.get("ProgramFiles", r"C:\Program Files")
                pf86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
                cand_dirs = [Path(pf) / "LibreOffice" / "program",
                             Path(pf86) / "LibreOffice" / "program"]
            elif sys.platform == "darwin":
                cand_dirs = [Path("/Applications/LibreOffice.app/Contents/MacOS")]
            else:
                cand_dirs = [Path("/usr/bin"), Path("/opt/libreoffice/program"),
                             Path("/snap/bin")]
            # Windows 优先 soffice.exe（启动器会设置环境；soffice.bin 直调
            # 可能静默失败不产出文件），其余平台用 soffice
            exe_names = ("soffice.exe", "soffice", "soffice.bin") \
                if sys.platform.startswith("win") else ("soffice", "soffice.bin")
            for d in cand_dirs:
                for exe in exe_names:
                    p = d / exe
                    if p.is_file():
                        found.append(("soffice", str(p)))
                        break
                if any(t[0] == "soffice" for t in found):
                    break
    if prefer in ("auto", "inkscape"):
        p = shutil.which("inkscape")
        if p:
            found.append(("inkscape", p))
    p = shutil.which("pdftoppm")
    if p:
        found.append(("pdftoppm", p))
    return found if found else None


def rasterize_vector(in_path: Path, out_png: Path, dpi: int = 150, prefer: str = "auto"):
    """把矢量图转 PNG。返回 (ok, note)。无可用工具时 ok=False 但不抛异常。"""
    tool = None
    for t in _detect_rasterizer(prefer):
        if t[0] in ("soffice", "inkscape"):
            tool = t
            break
    if not tool:
        return False, "未检测到 LibreOffice/Inkscape，矢量图保留原文件（可手动转换）"
    engine, exe = tool
    try:
        if engine == "soffice":
            out_dir = str(out_png.parent)
            subprocess.run([exe, "--headless", "--convert-to", "png",
                            "--outdir", out_dir, str(in_path)],
                           check=True, capture_output=True, timeout=180)
            generated = Path(out_dir) / (in_path.stem + ".png")
            if generated.exists() and generated.resolve() != out_png.resolve():
                generated.replace(out_png)
            if not out_png.exists():
                return False, "LibreOffice 未生成预期 PNG，保留矢量原文件"
            return True, "经 LibreOffice 栅格化"
        else:
            subprocess.run([exe, str(in_path), "--export-type=png",
                            f"--export-filename={out_png}", f"--export-dpi={dpi}"],
                           check=True, capture_output=True, timeout=180)
            return True, f"经 Inkscape 栅格化(DPI={dpi})"
    except Exception as e:  # pragma: no cover
        return False, f"栅格化失败({e})，保留矢量原文件"


def _count_pptx_slides(pptx_path: str) -> int:
    """统计 pptx 页数（zip 内 slideN.xml 数量）。"""
    try:
        import zipfile as _zf
        with _zf.ZipFile(pptx_path) as z:
            return sum(1 for n in z.namelist()
                       if re.match(r"ppt/slides/slide\d+\.xml$", n))
    except Exception:
        return 0


def render_pptx_pages(pptx_path: str, cache_dir: Path, dpi: int = 150):
    """用 LibreOffice 把整份 pptx 渲染为逐页 PNG（index 0=第1页）。
    依赖 soffice + pdftoppm，缺失则返回 None（优雅降级）。
    缓存复用：cache_dir 已有完整页数（与 pptx 页数一致）时直接返回，
    避免每次重跑 soffice/pdftoppm（大课件渲染可达数分钟）。"""
    soffice = None
    for t in _detect_rasterizer("soffice"):
        if t[0] == "soffice":
            soffice = t[1]
            break
    pdftoppm = None
    for t in _detect_rasterizer("auto"):
        if t[0] == "pdftoppm":
            pdftoppm = t[1]
            break
    if not soffice or not pdftoppm:
        return None
    n_slides = _count_pptx_slides(pptx_path)
    existing = sorted(cache_dir.glob("page-*.png")) \
        if cache_dir.is_dir() else []
    if n_slides and len(existing) >= n_slides:
        return existing
    try:
        cache_dir.mkdir(parents=True, exist_ok=True)
        # 清掉不完整的旧 png，避免混用
        for old in cache_dir.glob("page-*.png"):
            try:
                old.unlink()
            except OSError:
                pass
        subprocess.run([soffice, "--headless", "--convert-to", "pdf",
                        "--outdir", str(cache_dir), str(pptx_path)],
                       check=True, capture_output=True, timeout=600)
        pdfs = sorted(cache_dir.glob("*.pdf"))
        if not pdfs:
            return None
        pdf_path = pdfs[0]
        subprocess.run([pdftoppm, "-r", str(dpi), "-png",
                        str(pdf_path), str(cache_dir / "page")],
                       check=True, capture_output=True, timeout=600)
        pages = sorted(cache_dir.glob("page-*.png"))
        return pages or None
    except Exception:  # pragma: no cover
        return None


def emu_to_px(v, dpi):
    return int(round(int(v) * dpi / 914400.0))


def read_sld_size(pptx_path) -> tuple:
    """从 pptx 的 ppt/presentation.xml 读取页面尺寸 <p:sldSz cx cy>（EMU）。
    返回 (cx, cy)；解析失败返回 (None, None)。crop_page_png 的裁剪换算
    必须用真实页面尺寸（16:9 等非常规比例否则左侧会被切）。"""
    try:
        import re as _re
        import zipfile
        with zipfile.ZipFile(str(pptx_path)) as z:
            xml = z.read("ppt/presentation.xml").decode("utf-8", "ignore")
        m = _re.search(r'<p:sldSz[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"', xml)
        if not m:
            m = _re.search(r'<p:sldSz[^>]*\bcy="(\d+)"[^>]*\bcx="(\d+)"', xml)
        if not m:
            return None, None
        return int(m.group(1)), int(m.group(2))
    except Exception:  # pragma: no cover
        return None, None


def crop_page_png(page_png: Path, xfrm, dpi, out_png: Path,
                  sld_cx=None, sld_cy=None):
    """用 xfrm(EMU) 在整页渲染图上裁出对象区域，返回 (ok, w, h)。

    sx/sy 用「渲染图实际像素 / 真实幻灯片尺寸(EMU)」换算。v4 加法：
    sld_cx/sld_cy 缺省时退回标准 10in×7.5in（9144000×6858000 EMU），
    保证既有图片侧调用（不传参）行为不变；公式路径3 传入真实 <p:sldSz> 后
    对 16:9 等非常规比例幻灯片消除裁剪偏差。
    """
    if xfrm is None or not HAVE_PIL:
        return False, 0, 0
    try:
        ox, oy, cx, cy = xfrm
        im = Image.open(page_png).convert("RGBA")
        W, H = im.size
        sx = W / (sld_cx if sld_cx else 9144000.0)
        sy = H / (sld_cy if sld_cy else 6858000.0)
        x0 = int(round(int(ox) * sx))
        y0 = int(round(int(oy) * sy))
        x1 = x0 + int(round(int(cx) * sx))
        y1 = y0 + int(round(int(cy) * sy))
        x0, y0 = max(0, x0), max(0, y0)
        x1, y1 = min(W, x1), min(H, y1)
        if x1 > x0 and y1 > y0:
            im.crop((x0, y0, x1, y1)).save(out_png, "PNG")
            return True, x1 - x0, y1 - y0
        return False, 0, 0
    except Exception:  # pragma: no cover
        return False, 0, 0


# --------------------------------------------------------------------------
# 主提取流程
# --------------------------------------------------------------------------
def extract(pptx_path: str, out_dir: str, convert: bool = True,
            all_media: bool = False, with_fill: bool = True,
            with_bg_layout: bool = True, rasterize: bool = False,
            raster_dpi: int = 150, raster_prefer: str = "auto",
            crop: bool = True, min_crop: int = 64,
            with_atomic: bool = True,
            on_progress=None):
    """主提取流程。on_progress(page_no, n_slides, info) 可选进度回调（默认无）。
    返回 (records, atomic_objects)；with_atomic=False 时 atomic_objects=[]。"""
    pptx_path = Path(pptx_path)
    out_dir = Path(out_dir)
    by_page = out_dir / "by_page"
    by_page.mkdir(parents=True, exist_ok=True)
    if all_media:
        all_dir = out_dir / "all_media"
        all_dir.mkdir(parents=True, exist_ok=True)

    records: list = []
    atomic_objects: list = []
    n_vector_skipped = 0
    # 整页渲染缓存（惰性）：首次遇到需渲染的公式/图表时才调用 LibreOffice
    _sld_cx0, _sld_cy0 = read_sld_size(str(pptx_path))
    render_ctx = {"pages": None, "pptx": str(pptx_path),
                  "out": out_dir, "dpi": raster_dpi,
                  "sld_cx": _sld_cx0, "sld_cy": _sld_cy0}

    with zipfile.ZipFile(pptx_path) as zf:
        if all_media:
            for name in zf.namelist():
                if name.startswith("ppt/media/") and not name.endswith("/"):
                    (all_dir / Path(name).name).write_bytes(zf.read(name))

        slides = list_slides_in_order(zf)
        for page_no, slide_path in enumerate(slides, start=1):
            slide_dir = posixpath.dirname(slide_path)
            rels_path = posixpath.join(
                slide_dir, "_rels", posixpath.basename(slide_path) + ".rels")
            rels = read_rels(zf, rels_path)
            rels_full = read_rels_full(zf, rels_path)
            slide_xml = zf.read(slide_path)
            rec_start = len(records)

            # 1) 独立图片对象 <p:pic>
            for idx, (name, rid, src_rect, x, y, w, h) in enumerate(
                    iter_pictures(slide_xml), start=1):
                rec = _emit(zf, rels, slide_dir, page_no, idx, "picture",
                            name or f"Picture{idx}", rid, by_page, convert,
                            suffix="pic", src_rect=src_rect if crop else None,
                            min_crop=min_crop)
                rec.x, rec.y = x, y
                rec.shape_w, rec.shape_h = w, h
                records.append(rec)

            # 2) 形状填充图 <p:sp><p:spPr><a:blipFill>
            if with_fill:
                for idx, (name, rid, src_rect, x, y, w, h) in enumerate(
                        iter_fill_images(slide_xml), start=1):
                    rec = _emit(zf, rels, slide_dir, page_no, idx, "fill",
                                name, rid, by_page, convert, suffix="fill",
                                src_rect=src_rect if crop else None,
                                min_crop=min_crop)
                    rec.x, rec.y = x, y
                    rec.shape_w, rec.shape_h = w, h
                    records.append(rec)

            # 3) 背景图（幻灯片 → 布局 → 母版）
            bg_items = list(iter_background_chain(zf, slide_path, rels_full)) \
                if with_bg_layout else \
                [("SlideBackground", rid, slide_dir, rels, sr)
                 for rid, sr in _bg_blips(slide_xml)]
            for idx, (name, rid, base_dir, base_rels, src_rect) in enumerate(bg_items, start=1):
                rec = _emit(zf, base_rels, base_dir, page_no, idx, "background",
                            name, rid, by_page, convert, suffix="bg",
                            src_rect=src_rect if crop else None, min_crop=min_crop)
                records.append(rec)

            # 4) 公式对象（oleObj → embeddings/oleObjectN.bin；Visio → .vsdx/.vsd）
            for idx, (name, rid, xfrm, prog_id) in enumerate(iter_ole_formulas(slide_xml), start=1):
                rec = _emit_ole(zf, rels, slide_dir, page_no, idx, name, rid,
                                xfrm, by_page, render_ctx, prog_id)
                if xfrm is not None:
                    try:
                        rec.x = int(xfrm[0]) / 914400 * 96
                        rec.y = int(xfrm[1]) / 914400 * 96
                        rec.shape_w = int(xfrm[2]) / 914400 * 96
                        rec.shape_h = int(xfrm[3]) / 914400 * 96
                    except Exception:
                        pass
                records.append(rec)

            # 5) OMML 公式（<m:oMath>）
            for idx, omxml in enumerate(iter_omath(slide_xml), start=1):
                rec = _emit_omath(page_no, idx, omxml, by_page)
                records.append(rec)

            # 6) 图表（<p:graphicFrame><c:chart>）
            for idx, (name, rid, xfrm) in enumerate(iter_charts(slide_xml), start=1):
                rec = _emit_chart(zf, rels, slide_dir, page_no, idx, name, rid,
                                  xfrm, by_page, render_ctx)
                if xfrm is not None:
                    try:
                        rec.x = int(xfrm[0]) / 914400 * 96
                        rec.y = int(xfrm[1]) / 914400 * 96
                        rec.shape_w = int(xfrm[2]) / 914400 * 96
                        rec.shape_h = int(xfrm[3]) / 914400 * 96
                    except Exception:
                        pass
                records.append(rec)

            # 7) 原子对象收集（可视逻辑块的数据地基：shape/connector/table +
            #    图片记录映射）。每页按 z_index 顺序输出。
            if with_atomic:
                atomic_objects.extend(_collect_atomic_objects(
                    page_no, slide_xml, records[rec_start:]))

            # 进度回调（可选，默认 None 不改变任何行为）
            if on_progress is not None:
                try:
                    page_items = [{
                        "file": r.output_file, "kind": r.kind,
                        "w": r.width, "h": r.height,
                        "x": int(r.x), "y": int(r.y),
                        "shape": r.shape_name,
                    } for r in records[rec_start:]]
                    on_progress(page_no, len(slides),
                                {"kind": "img", "objects": len(records),
                                 "page_items": page_items})
                except Exception:
                    pass

    # 7) 矢量栅格化（后处理）
    if rasterize:
        for rec in records:
            if rec.original_format in VECTOR and rec.output_file:
                vec_path = by_page / rec.output_file
                if vec_path.exists():
                    png_path = vec_path.with_suffix(".png")
                    ok, note = rasterize_vector(vec_path, png_path,
                                                dpi=raster_dpi, prefer=raster_prefer)
                    if ok:
                        try:
                            im = Image.open(png_path)
                            rec.width, rec.height = im.size
                        except Exception:
                            pass
                        rec.output_file = png_path.name
                        rec.converted_to_png = True
                        rec.note = (rec.note + "；" if rec.note else "") + note
                    else:
                        n_vector_skipped += 1
                        rec.note = (rec.note + "；" if rec.note else "") + note
    else:
        n_vector_skipped = sum(1 for rec in records
                               if rec.original_format in VECTOR and not rec.converted_to_png)

    _write_manifest(out_dir, records, len(slides), n_vector_skipped)

    # 原子对象落盘（供可视逻辑块聚类使用）
    if with_atomic and atomic_objects:
        try:
            (out_dir / "atomic_objects.json").write_text(
                json.dumps(atomic_objects, ensure_ascii=False, indent=1),
                encoding="utf-8")
        except OSError:
            pass
    return records, atomic_objects


def _emit(zf, rels, base_dir, page_no, idx, kind, name, rid, by_page, convert,
          suffix="pic", src_rect=None, min_crop=64):
    """解析单个关系并落盘，返回 ImageRecord。rels/base_dir 指向【媒体所在】层级。"""
    if not rid or rid not in rels:
        return ImageRecord(page_no, idx, kind, name, "(missing)", "",
                           note="找不到对应媒体关系")
    target = resolve_target(base_dir, rels[rid])
    if target not in zf.namelist():
        return ImageRecord(page_no, idx, kind, name, target, "",
                           note="媒体文件在包内缺失")
    data = zf.read(target)
    out_base = by_page / f"slide_{page_no:02d}_{suffix}_{idx:02d}"
    written, conv, w, h, note, cropped = save_image(
        data, out_base, target, convert, src_rect=src_rect, min_crop=min_crop)
    return ImageRecord(page_no, idx, kind, name, posixpath.basename(target),
                       written.name, w, h, ext_of(target), conv, note,
                       src_rect, cropped)


def _emit_ole(zf, rels, slide_dir, page_no, idx, name, rid, xfrm, by_page,
              render_ctx, prog_id=""):
    """oleObj → embeddings/oleObjectN.bin。有 LO 整页渲染则裁出 PNG，否则保留 bin。
    prog_id 用于标注对象类型（如 Equation.DSMT4 = 公式编辑器；
    Visio.Drawing.* = Visio 矢量图，走 _emit_visio 存 .vsdx/.vsd）。"""
    if "visio" in (prog_id or "").lower():
        return _emit_visio(zf, rels, slide_dir, page_no, idx, name, rid,
                           xfrm, by_page, prog_id)
    is_eq = "equation" in (prog_id or "").lower()
    target = resolve_target(slide_dir, rels.get(rid, ""))
    if not target or target not in zf.namelist():
        return ImageRecord(page_no, idx, "formula_ole", name, target or "(missing)",
                           "", note="OLE 关系缺失")
    data = zf.read(target)
    out_base = by_page / f"slide_{page_no:02d}_formula_{idx:02d}"
    # 惰性整页渲染：首次需要时才调用 LibreOffice
    pages = _ensure_rendered(render_ctx)
    if pages is not None and (page_no - 1) < len(pages):
        png_path = out_base.with_suffix(".png")
        ok, w, h = crop_page_png(pages[page_no - 1], xfrm, render_ctx["dpi"],
                                 png_path, sld_cx=render_ctx.get("sld_cx"),
                                 sld_cy=render_ctx.get("sld_cy"))
        if ok:
            return ImageRecord(page_no, idx, "formula_ole", name,
                               posixpath.basename(target), png_path.name,
                               w, h, "ole", True,
                               note="经 LibreOffice 整页渲染+xfrm 裁剪"
                                    + ("（公式编辑器）" if is_eq else ""))
    out = out_base.with_suffix(".bin")
    out.write_bytes(data)
    label = "公式编辑器(Equation)" if is_eq else "嵌入对象(OLE)"
    return ImageRecord(page_no, idx, "formula_ole", name, posixpath.basename(target),
                       out.name, 0, 0, "ole", False,
                       note=f"{label}，保留源；建议装 LibreOffice 渲染")


def _emit_visio(zf, rels, slide_dir, page_no, idx, name, rid, xfrm, by_page,
                prog_id=""):
    """Visio OLE 对象 → 按容器类型落盘 .vsdx（zip）或 .vsd（OLE 复合文档）。

    识别规则（按文件头魔数）：
      - b'PK\\x03\\x04'（zip）且含 visio/document.xml → 新版 .vsdx；
      - b'\\xd0\\xcf\\x11\\xe0'（OLE2 复合文档）→ 旧版 .vsd。
    二者均为 Visio 原生矢量格式，直接落盘供编辑/二次加工。
    """
    target = resolve_target(slide_dir, rels.get(rid, ""))
    if not target or target not in zf.namelist():
        return ImageRecord(page_no, idx, "visio", name, target or "(missing)",
                           "", ole_progid=prog_id, note="Visio 关系缺失")
    data = zf.read(target)
    head = data[:4]
    if head[:2] == b"PK":
        ext = ".vsdx"
    elif head == b"\xd0\xcf\x11\xe0":
        ext = ".vsd"
    else:
        ext = ".bin"
    out = by_page / f"slide_{page_no:02d}_visio_{idx:02d}{ext}"
    out.write_bytes(data)
    fmt = "Visio 矢量图(.vsdx)" if ext == ".vsdx" else \
        ("Visio 矢量图(.vsd)" if ext == ".vsd" else "Visio OLE(容器未知)")
    return ImageRecord(page_no, idx, "visio", name,
                       posixpath.basename(target), out.name, 0, 0,
                       ext.lstrip("."), False,
                       ole_progid=prog_id,
                       note=f"{fmt}，直接存原文件；预览图通常以同页 "
                            "EMF/SVG 图片形式另行提取")


def _emit_omath(page_no, idx, omxml, by_page):
    """OMML → 保存 omml 片段 + 扁平文本（便于索引/转 LaTeX）。"""
    out = by_page / f"slide_{page_no:02d}_formula_{idx:02d}.omml.xml"
    out.write_text('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                   + omxml, encoding="utf-8")
    # 扁平化文字（公式内字面量），供检索；完整 LaTeX 转换需专业库（如 omml2latex）
    try:
        root = ET.fromstring(omxml)
        text = "".join(t.text or "" for t in root.iter(q("m", "t")) if t.text)
    except Exception:
        text = ""
    return ImageRecord(page_no, idx, "formula_omath", "OMML公式",
                       "inline-math", out.name, 0, 0, "omml", False,
                       note=f"OMML 公式（文本片段：{text[:40]}）；转 LaTeX 需专业库")


def _emit_chart(zf, rels, slide_dir, page_no, idx, name, rid, xfrm, by_page,
                render_ctx):
    """c:chart r:id → charts/chartN.xml。有 LO 整页渲染则裁出 PNG，否则保留 xml。"""
    target = resolve_target(slide_dir, rels.get(rid, ""))
    if not target or target not in zf.namelist():
        return ImageRecord(page_no, idx, "chart", name, target or "(missing)",
                           "", note="图表关系缺失")
    data = zf.read(target)
    out_base = by_page / f"slide_{page_no:02d}_chart_{idx:02d}"
    pages = _ensure_rendered(render_ctx)
    if pages is not None and (page_no - 1) < len(pages):
        png_path = out_base.with_suffix(".png")
        ok, w, h = crop_page_png(pages[page_no - 1], xfrm, render_ctx["dpi"],
                                 png_path, sld_cx=render_ctx.get("sld_cx"),
                                 sld_cy=render_ctx.get("sld_cy"))
        if ok:
            return ImageRecord(page_no, idx, "chart", name,
                               posixpath.basename(target), png_path.name,
                               w, h, "chart", True,
                               note="经 LibreOffice 整页渲染+xfrm 裁剪")
    out = out_base.with_suffix(".xml")
    out.write_bytes(data)
    return ImageRecord(page_no, idx, "chart", name, posixpath.basename(target),
                       out.name, 0, 0, "chart", False,
                       note="图表以 chartN.xml 保留（建议装 LibreOffice 渲染）")


def _ensure_rendered(render_ctx):
    """惰性渲染整份 pptx 为逐页 PNG（仅一次）；无工具时返回 None（降级）。"""
    if render_ctx["pages"] is not None:
        return render_ctx["pages"]
    render_ctx["pages"] = render_pptx_pages(
        render_ctx["pptx"], render_ctx["out"] / "_render_cache",
        dpi=render_ctx["dpi"])
    return render_ctx["pages"]


def _write_manifest(out_dir: Path, records: list, n_slides: int, n_vector_skipped: int = 0):
    csv_path = out_dir / "manifest.csv"
    json_path = out_dir / "manifest.json"
    cols = ["page", "index", "kind", "shape_name", "source_media", "output_file",
            "width", "height", "original_format", "converted_to_png", "note",
            "src_rect_l", "src_rect_t", "src_rect_r", "src_rect_b", "cropped",
            "x", "y", "shape_w", "shape_h", "ole_progid", "preview_file"]
    with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
        wcsv = csv.writer(f)
        wcsv.writerow(cols)
        for r in records:
            sr = r.src_rect or ("", "", "", "")
            wcsv.writerow([r.page, r.index, r.kind, r.shape_name, r.source_media,
                           r.output_file, r.width, r.height, r.original_format,
                           r.converted_to_png, r.note,
                           sr[0], sr[1], sr[2], sr[3], r.cropped,
                           int(r.x), int(r.y), int(r.shape_w), int(r.shape_h),
                           r.ole_progid, r.preview_file])
    json_path.write_text(json.dumps([asdict(r) for r in records],
                                    ensure_ascii=False, indent=2), encoding="utf-8")

    pics = [r for r in records if r.kind == "picture"]
    fills = [r for r in records if r.kind == "fill"]
    bgs = [r for r in records if r.kind == "background"]
    f_ole = [r for r in records if r.kind == "formula_ole"]
    f_om = [r for r in records if r.kind == "formula_omath"]
    charts = [r for r in records if r.kind == "chart"]
    visios = [r for r in records if r.kind == "visio"]
    cropped = [r for r in records if r.cropped == "yes"]
    # 跨页复用：同一 source_media 被多个记录引用
    media_cnt = Counter(r.source_media for r in records
                        if r.source_media and r.source_media not in ("(missing)", ""))
    cross = sum(1 for m, c in media_cnt.items() if c > 1)

    print(f"[OK] 处理完成：{n_slides} 页")
    print(f"     独立图片对象 {len(pics)} · 形状填充图 {len(fills)} · 背景图 {len(bgs)}")
    print(f"     公式 OLE {len(f_ole)} · OMML 公式 {len(f_om)} · 图表 {len(charts)}"
          + (f" · Visio 矢量 {len(visios)}" if visios else ""))
    print(f"     其中 已按 srcRect 裁剪 {len(cropped)} 张")
    print(f"     矢量未栅格化 {n_vector_skipped} 张（已保留原文件）")
    print(f"     同媒体被多 shape 引用（跨页复用）{cross} 次")
    print(f"     输出目录：{out_dir}")
    print(f"     清单：{csv_path.name} / {json_path.name}")
    if not HAVE_PIL:
        print("[提示] 未检测到 Pillow，栅格图按原格式落盘（未统一转 PNG）。")


# --------------------------------------------------------------------------
# v3 公式 LaTeX 提取（完全附加；不改动任何图片链路）
# --------------------------------------------------------------------------
def _make_latex_converter():
    """返回 omml2latex.convert_omml（接受 Element，返回 LaTeX 字符串），缺失则返回 None。"""
    try:
        from omml2latex import convert_omml
        return convert_omml
    except Exception:
        return None


def _wrap_omml_for_convert(omml_str: str):
    """iter_omath 返回的 OMML 片段不含命名空间声明，omml2latex 需要自包含 XML。
    包一层 <m:oMathPara> 并声明 m/a/w 命名空间，返回 ElementTree Element。"""
    M, A = NS["m"], NS["a"]
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    wrapped = ('<m:oMathPara xmlns:m="%s" xmlns:a="%s" xmlns:w="%s">%s</m:oMathPara>'
               % (M, A, W, omml_str))
    return ET.fromstring(wrapped)


def _latex_to_block(latex: str, display: bool) -> str:
    """omml2latex 总是输出 $$...$$ 块；按 display 决定保留为独立公式或压成 $...$。"""
    s = (latex or "").strip()
    if s.startswith("$$") and s.endswith("$$"):
        inner = s[2:-2].strip()
        if display:
            return f"$$\n{inner}\n$$"
        return f"${inner}$"
    return s


def _convert_one(converter, omml_str):
    """单条 OMML→LaTeX。返回 (latex, ok)。converter 为 None 或异常时 ok=False。"""
    if converter is None:
        return "", False
    try:
        return converter(_wrap_omml_for_convert(omml_str)), True
    except Exception as e:
        print(f"[警告] OMML→LaTeX 转换失败（{e}），该公式将以占位进入 md。",
              file=sys.stderr)
        return "", False


def _read_slide_size(zf: zipfile.ZipFile):
    """从 presentation.xml 读 <p:sldSz cx cy>（EMU），返回 (cx, cy) 或 (None,None)。"""
    try:
        root = ET.fromstring(zf.read("ppt/presentation.xml"))
        sz = root.find(".//" + q("p", "sldSz"))
        if sz is not None:
            return int(sz.get("cx") or 0) or None, int(sz.get("cy") or 0) or None
    except Exception:
        pass
    return None, None


def _make_eq3_converter():
    """返回 eq3_to_latex(ole_bytes)->(latex, ok) 或 None（依赖缺失时）。

    路径2：Equation.3 OLE 的二进制在 embeddings/oleObjectN.bin（OLE2 复合文档），
    公式本体在名为 "Equation Native" 的流里（MTEF 二进制）。

    v4 修正：接入通用 MTEF→LaTeX 解码器（mtef_decoder 包，vendor 自
    pptkb.formula_math.mtef，只读参考未改原文件）。该解码器从 OLE2 复合文档
    抽取 Equation Native 流并按 MTEF v3 规范解析为 LaTeX；needs_review
    （best-effort，如存在未消费尾字节但公式内容已完整）也视为可用。
    若解码器模块或 olefile 缺失，返回 None，交由路径3（LO 渲染+数学 OCR）
    兜底，整体不崩溃（与既有优雅降级行为一致）。
    """
    try:
        from pptx_wzq.mtef_decoder import eq3_ole_bytes_to_latex
    except Exception:
        # 兜底：确保脚本同目录在 sys.path（以脚本方式运行时通常已在）
        import os
        import sys
        _here = os.path.dirname(os.path.abspath(__file__))
        if _here not in sys.path:
            sys.path.insert(0, _here)
        try:
            from pptx_wzq.mtef_decoder import eq3_ole_bytes_to_latex
        except Exception:
            return None

    def eq3_to_latex(ole_bytes):
        latex, ok, _status = eq3_ole_bytes_to_latex(ole_bytes)
        return latex, ok

    return eq3_to_latex


def _make_math_ocr(engine: str = "auto"):
    """返回 math_ocr(image_path)->(latex, ok) 或 None（无可用 OCR 引擎）。

    运行时探测（均缺失即返回 None，优雅降级）：
      1. Mathpix：需环境变量 MATHPIX_APP_ID / MATHPIX_APP_KEY；
      2. 本地 LaTeX-OCR（pix2tex）：需已安装 pix2tex 且模型权重可用（pip install pix2tex）。
    """
    app_id = os.environ.get("MATHPIX_APP_ID")
    app_key = os.environ.get("MATHPIX_APP_KEY")
    if (engine in ("auto", "mathpix")) and app_id and app_key:
        def mathpix(path):
            import base64
            import json
            import urllib.request
            try:
                with open(path, "rb") as f:
                    b64 = base64.b64encode(f.read()).decode("ascii")
                req = urllib.request.Request(
                    "https://api.mathpix.com/v3/text",
                    data=json.dumps({"src": b64,
                                     "formats": ["latex_styled"]}).encode("utf-8"),
                    headers={"app_id": app_id, "app_key": app_key,
                             "Content-Type": "application/json"},
                    method="POST")
                with urllib.request.urlopen(req, timeout=30) as resp:
                    data = json.loads(resp.read().decode("utf-8"))
                latex = data.get("latex_styled") or data.get("latex") or ""
                return latex, bool(latex)
            except Exception:
                return "", False
        return mathpix
    if engine in ("auto", "pix2tex"):
        try:
            from pix2tex.cli import LatexOCR  # 仅探测可用性
        except Exception:
            LatexOCR = None
        if LatexOCR is not None:
            _model = None
            def pix2tex(path):
                nonlocal _model
                try:
                    from PIL import Image
                    if _model is None:
                        _model = LatexOCR()   # 首次调用加载模型/权重（较慢）
                    img = Image.open(path).convert("RGB")
                    latex = _model(img)
                    return (latex or ""), bool(latex)
                except Exception:
                    return "", False
            return pix2tex
    return None


def _fmt_entry(page_no, idx, e):
    """把一条公式条目格式化为 markdown。source ∈ omml|eq3|ocr|placeholder。"""
    head = f"- **(p{page_no}-f{idx})** "
    src = e["source"]
    if src == "omml":
        if e["ok"] and e["latex"].strip():
            body = _latex_to_block(e["latex"], e["display"])
        else:
            body = "_（OMML 已提取，但 omml2latex 未安装/转换失败，详见下方 OMML 源）_"
        fence = "\n\n  ```omml\n" + e["omml"] + "\n  ```\n"
        return head + body + fence
    if src == "eq3":
        body = _latex_to_block(e["latex"], True)
        # 解码器输出为裸 LaTeX，未带数学定界符时补 $$ 以便下游渲染
        if not body.strip().startswith("$$"):
            body = f"$$\n{body.strip()}\n$$"
        return head + body + "\n\n  > 来源：EQ3-MTEF 解析\n"
    if src == "ocr":
        body = _latex_to_block(e["latex"], True)
        if not body.strip().startswith("$$"):
            body = f"$$\n{body.strip()}\n$$"
        crop = e.get("crop_png") or ""
        return (head + body +
                f"\n\n  > 来源：OCR（{e.get('ocr_engine', '')}）；"
                f"裁剪图：`formula_crops/{crop}`\n")
    # placeholder
    name = e.get("name", "")
    is_eq = e.get("is_eq", False)
    crop = e.get("crop_png") or ""
    hint = (f"`formula_crops/{crop}`" if crop
            else f"`by_page/slide_{page_no:02d}_formula_*.bin`")
    if is_eq:
        lead = (f"公式对象（`{name}`）旧式 Equation OLE，未获 LaTeX；"
                f"源见 {hint}（可经 EQ3-MTEF 解析或 LO 渲染+数学 OCR）。")
    else:
        lead = f"OLE 嵌入对象（`{name}`）未获 LaTeX；源见 {hint}。"
    return head + lead


# --------------------------------------------------------------------------
# 公式符号判据过滤（v5）：无数学符号的碎片舍弃
# --------------------------------------------------------------------------
# 教学公式判据：真实公式一般不只 1-2 个符号，会带关系/比较/逻辑/运算符号。
# 命中以下任一记号即视为「有价值的数学公式」；未命中视为误识别碎片。
_FORMULA_SYMBOL_RE = re.compile(
    r"(?:\\(?:neq|ne|approx|approxeq|equiv|cong|leq|geq|le|ge|pm|mp|times|div|"
    r"cdot|ast|star|sum|prod|int|iint|iiint|oint|frac|dfrac|tfrac|sqrt|partial|"
    r"nabla|infty|lim|rightarrow|leftarrow|leftrightarrow|Rightarrow|Leftarrow|"
    r"Leftrightarrow|mapsto|wedge|vee|neg|land|lor|in|notin|subset|subseteq|"
    r"supset|supseteq|cup|cap|setminus|propto|sim|simeq|asymp|doteq|angle|perp|"
    r"parallel|therefore|because|forall|exists|lt|gt|preceq|succeq|big|Big|"
    r"left\(|right\)|begin|end)|[=<>±×÷∧∨¬∈⊂⊆∪∩≈≠≤≥≡∼∝→←↔⇒∀∃∑∫∏√∂∇∞])"
)


def filter_formula_entries(entries):
    """按符号判据过滤一页条目：无关系/比较/逻辑/运算符号 → 视为碎片舍弃。

    占位条目（无 LaTeX 文本，如解析失败）不属于「误识别」，一律保留；
    仅对已产出 LaTeX 的条目（omml/eq3/ocr）执行判据。
    返回 (kept, dropped)。
    """
    kept, dropped = [], []
    for e in entries:
        latex = (e.get("latex") or "").strip()
        if latex and not _FORMULA_SYMBOL_RE.search(latex):
            dropped.append(e)
        else:
            kept.append(e)
    return kept, dropped


def extract_latex(pptx_path, out_dir, do_lo: bool = True,
                  eq3_converter=None, math_ocr=None,
                  no_eq3: bool = False, no_ocr: bool = False, ocr_engine: str = "auto",
                  on_progress=None, filter_formulas: bool = True):
    """v4：把整份 PPTX 的全部公式以 LaTeX 汇总到 <stem>_formulas.md。
    on_progress(page_no, n_slides, info) 可选进度回调（默认无）。
    filter_formulas=True（默认）：写 md 前按符号判据过滤误识别碎片（见
    filter_formula_entries），被过滤条目写入 <stem>_filtered_entries.json 审计。

    三路径 per-object 级联（纯附加，不改动任何图片处理代码/行为）：
      路径1 原生 OMML（<m:oMath>）→ omml2latex；
      路径2 Equation OLE（progId 含 equation）→ eq3_to_latex（MTEF 解析）；
      路径3 LO 渲染整页 + xfrm 裁剪 + 数学 OCR → LaTeX；
      全失败 → 占位（保留 .bin 或裁剪 PNG 溯源）。

    仅 `progId` 含 equation 的 OLE 进入公式清单（照片/Visio 已排除，修 F1）；
    原生 OMML 与 Equation OLE 均计入 md，确保与结构期望数对账一致。
    """
    pptx_path = Path(pptx_path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = pptx_path.stem
    md_path = out_dir / f"{stem}_formulas.md"
    crops_dir = out_dir / "formula_crops"
    crops_dir.mkdir(parents=True, exist_ok=True)

    omml_conv = _make_latex_converter()
    if omml_conv is None:
        print("[警告] 未安装 omml2latex（pip install omml2latex），"
              "原生 OMML 公式将仅以占位进入 md（不崩溃）。", file=sys.stderr)
    if eq3_converter is None and not no_eq3:
        eq3_converter = _make_eq3_converter()
    if math_ocr is None and not no_ocr:
        math_ocr = _make_math_ocr(ocr_engine)

    dpi = 150
    render_ctx = {"pages": None, "pptx": str(pptx_path),
                  "out": out_dir, "dpi": dpi}

    with zipfile.ZipFile(pptx_path) as zf:
        sld_cx, sld_cy = _read_slide_size(zf)
        slides = list_slides_in_order(zf)
        entries_by_page = {}
        pages_with_formula = set()
        counts = {"native": 0, "eq3": 0, "ocr": 0, "placeholder": 0}
        expected_native = 0
        expected_eq = 0

        for page_no, slide_path in enumerate(slides, start=1):
            slide_dir = posixpath.dirname(slide_path)
            rels = read_rels(zf, posixpath.join(
                slide_dir, "_rels", posixpath.basename(slide_path) + ".rels"))
            slide_xml = zf.read(slide_path)
            root = ET.fromstring(slide_xml)
            parent = {c: p for p in root.iter() for c in p}
            page_entries = []

            # --- 路径1：原生 OMML ---
            for om in root.iter(q("m", "oMath")):
                disp = (parent.get(om) is not None
                        and parent[om].tag == q("m", "oMathPara"))
                om_str = ET.tostring(om, encoding="unicode")
                expected_native += 1
                pages_with_formula.add(page_no)
                counts["native"] += 1
                latex, ok = _convert_one(omml_conv, om_str)
                page_entries.append({"source": "omml", "omml": om_str,
                                     "latex": latex, "ok": ok,
                                     "display": disp, "name": "OMML公式"})

            # --- B 类：Equation OLE（仅 progId 含 equation 进公式清单）---
            for (name, rid, xfrm, prog_id) in iter_ole_formulas(slide_xml):
                if "equation" not in (prog_id or "").lower():
                    continue   # 照片/Visio 等非公式 OLE 不进公式清单（修 F1）
                expected_eq += 1
                pages_with_formula.add(page_no)
                target = resolve_target(slide_dir, rels.get(rid, ""))
                ole_bytes = zf.read(target) if target in zf.namelist() else b""

                # 路径2：EQ3-MTEF 解析
                latex, ok = ("", False)
                if eq3_converter is not None:
                    latex, ok = eq3_converter(ole_bytes)
                if ok and latex.strip():
                    counts["eq3"] += 1
                    page_entries.append({"source": "eq3", "latex": latex,
                                         "ok": True, "name": name, "is_eq": True})
                    continue

                # 路径3：LO 渲染整页 + xfrm 裁剪 + 数学 OCR
                crop_png = None
                pages = _ensure_rendered(render_ctx)
                if pages is not None and (page_no - 1) < len(pages):
                    crop_path = (crops_dir /
                                 f"slide_{page_no:02d}_equation_"
                                 f"{len([e for e in page_entries if e['source'] != 'omml']) + 1:02d}.png")
                    ok_crop, _, _ = crop_page_png(pages[page_no - 1], xfrm, dpi,
                                                  crop_path, sld_cx, sld_cy)
                    if ok_crop:
                        crop_png = crop_path.name
                        if math_ocr is not None:
                            ocr_latex, ocr_ok = math_ocr(str(crop_path))
                            if ocr_ok and ocr_latex.strip():
                                counts["ocr"] += 1
                                page_entries.append({
                                    "source": "ocr", "latex": ocr_latex,
                                    "ok": True, "crop_png": crop_png,
                                    "ocr_engine": ocr_engine,
                                    "name": name, "is_eq": True})
                                continue
                # 全失败 → 占位
                counts["placeholder"] += 1
                page_entries.append({"source": "placeholder", "name": name,
                                     "is_eq": True, "crop_png": crop_png})

            if page_entries:
                entries_by_page[page_no] = page_entries

            # 进度回调（可选，默认 None 不改变任何行为）
            if on_progress is not None:
                try:
                    def _brief(e):
                        latex = (e.get("latex") or "").replace("\n", " ")
                        if len(latex) > 48:
                            latex = latex[:48] + "…"
                        return {"source": e.get("source", "?"),
                                "latex": latex,
                                "name": e.get("name", "")}
                    on_progress(page_no, len(slides), {
                        "kind": "formula",
                        "page_entries": len(page_entries),
                        "total": (counts["native"] + counts["eq3"]
                                  + counts["ocr"] + counts["placeholder"]),
                        "page_items": [_brief(e) for e in page_entries],
                    })
                except Exception:
                    pass

    # 写 md（先按符号判据过滤误识别碎片，filter_formulas=False 时跳过）
    n_filtered = 0
    filtered_detail = []
    if filter_formulas:
        for page_no in list(entries_by_page):
            kept, dropped = filter_formula_entries(entries_by_page[page_no])
            if dropped:
                n_filtered += len(dropped)
                for e in dropped:
                    filtered_detail.append({
                        "page": page_no,
                        "source": e.get("source"),
                        "latex": (e.get("latex") or "")[:200],
                        "name": e.get("name"),
                        "reason": "无关系/比较/逻辑/运算符号（判据过滤）",
                    })
            entries_by_page[page_no] = kept
    try:
        if filtered_detail:
            (out_dir / f"{stem}_filtered_entries.json").write_text(
                json.dumps(filtered_detail, ensure_ascii=False, indent=2),
                encoding="utf-8")
    except Exception:
        pass

    lines = [f"# {stem} 公式清单", "",
             "> 由 `extract_pptx_images.py` 自动提取（v4 三路径 + v5 符号判据过滤）。",
             "路径1 原生 OMML→LaTeX；路径2 EQ3-MTEF 解析→LaTeX；"
             "路径3 LO 渲染+xfrm 裁剪+数学 OCR→LaTeX；全失败占位。",
             "仅 `progId` 含 equation 的 OLE 进入公式清单（照片/Visio 已排除）；"
             "无数学符号的碎片条目被过滤（详见 <名>_filtered_entries.json）。", ""]
    total = 0
    for page_no in sorted(entries_by_page):
        entries = entries_by_page[page_no]
        lines.append(f"## 第 {page_no} 页")
        lines.append("")
        for i, e in enumerate(entries, start=1):
            total += 1
            lines.append(_fmt_entry(page_no, i, e))
            lines.append("")
    lines.append("---")
    lines.append("")
    lines.append(f"公式总数 **{total}** 条")
    lines.append(f"- 原生 OMML（路径1）：{counts['native']}")
    lines.append(f"- EQ3-MTEF 解析（路径2）：{counts['eq3']}")
    lines.append(f"- OCR 识别（路径3）：{counts['ocr']}")
    lines.append(f"- 降级占位：{counts['placeholder']}")
    if filter_formulas:
        lines.append(f"- 符号判据过滤舍弃：{n_filtered}")
    lines.append("")
    lines.append("> 结构对账：期望公式条目 = 原生 OMML "
                 f"{expected_native} + Equation OLE {expected_eq} = "
                 f"{expected_native + expected_eq}；实际 {total}"
                 + ("（过滤后）" if n_filtered else "")
                 + "。"
                 + (" ✅ 一致" if total + n_filtered == expected_native + expected_eq
                    else " ⚠️ 不一致"))
    md_path.write_text("\n".join(lines), encoding="utf-8")

    print(f"[OK] 公式 LaTeX 清单已写出：{md_path.name}")
    print(f"     总数 {total}（原生 {counts['native']} / EQ3 {counts['eq3']}"
          f" / OCR {counts['ocr']} / 占位 {counts['placeholder']}"
          + (f" / 判据过滤 {n_filtered}" if filter_formulas else "") + "）")
    return {"total": total, "native": counts["native"], "eq3": counts["eq3"],
            "ocr": counts["ocr"], "placeholder": counts["placeholder"],
            "filtered": n_filtered,
            "expected": expected_native + expected_eq,
            "pages": len(pages_with_formula), "md": str(md_path)}


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------
def main(argv=None):
    ap = argparse.ArgumentParser(
        description="通用 PPTX 图片提取器（OOXML）：独立图片/形状填充/背景/公式/图表 → 独立文件（标注页码）")
    ap.add_argument("pptx", help="输入的 .pptx 文件路径")
    ap.add_argument("-o", "--output", default=None,
                    help="输出目录（默认 <pptx名>_images）")
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
    ap.add_argument("--latex", action="store_true",
                    help="[v4] 额外把全部公式以 LaTeX 汇总导出为 <名>_formulas.md（三路径级联，不改动图片提取）")
    ap.add_argument("--no-eq3", action="store_true",
                    help="[v4] 关闭路径2（EQ3-MTEF 解析），强制 OLE 公式走路径3/占位")
    ap.add_argument("--no-ocr", action="store_true",
                    help="[v4] 关闭路径3 的 OCR（仍渲染+裁剪便于人工核验，但不调 OCR）")
    ap.add_argument("--ocr-engine", default="auto",
                    choices=["auto", "mathpix", "pix2tex"],
                    help="[v4] 指定优先 OCR 引擎（默认 auto）")
    args = ap.parse_args(argv)

    pptx_norm = normalize_path(args.pptx)
    if not Path(pptx_norm).is_file():
        print(f"[错误] 找不到文件：{args.pptx}（归一化后：{pptx_norm}）", file=sys.stderr)
        return 2
    out = args.output or (Path(pptx_norm).stem + "_images")
    out = normalize_path(out)
    extract(pptx_norm, out,
            convert=not args.no_convert,
            all_media=args.all_media,
            with_fill=not args.no_fill,
            with_bg_layout=not args.no_bg_layout,
            rasterize=args.rasterize_vector,
            raster_dpi=args.raster_dpi,
            raster_prefer=args.raster_prefer,
            crop=not args.no_crop,
            min_crop=args.min_crop)
    if args.latex:
        extract_latex(pptx_norm, out, no_eq3=args.no_eq3, no_ocr=args.no_ocr,
                      ocr_engine=args.ocr_engine)
    return 0


def main() -> int:  # console
    return _main()


if __name__ == "__main__":
    raise SystemExit(main())
