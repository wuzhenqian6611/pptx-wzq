#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
img_filter.py — PPTX 提取图的教学用途过滤（规则层 + YOLO 层）
================================================================

承接 extract_pptx_images.py 的产物（by_page/ 下的独立 PNG + ImageRecord 列表），
按教学场景判据对图片取舍。判据（与用户约定一致）：

  「教学 PPT 里的图：照片、框图、电路图、结构图、设计图、波形图
   （各种矢量/像素格式皆可）；绝不会是：纯色图、PPT背景图、
   特别窄/细长、尺寸特别小、某公司 LOGO。」

两层过滤：
  1) 规则层（确定性，不依赖模型）：
     - 尺寸过小：宽或高 < min_size（默认 48px）；
     - 特别细长：宽高比 > max_ratio（默认 10:1）；
     - 纯色/近纯色：缩小采样后颜色种类 ≤ min_colors（默认 4）。
     （矢量图 wmf/emf/svg 与未知尺寸不参与规则层。）
  2) YOLO 层（可选，本地模型）：
     - 规则命中但 YOLO 检测到明确物体（照片/实物）→ 强保留（防误删照片），
       并在 note 标注检测内容；
     - 规则未命中 → 一律保留（框图/电路图/波形图 YOLO 检不出类别，不因
       "无检测结果"删除）。
  未检测到本地权重时 YOLO 层自动跳过（打印提示），规则层照常生效。

被过滤的图移动到 <out>/discarded/ 子目录（可审计、可恢复），
取舍明细写入 <out>/filter_report.json。不修改 manifest 16 列协议。

作者：吴振谦 · wuzhenqian@nbu.edu.cn · QQ：38328063"""
from __future__ import annotations

import json
import re
import shutil
import struct
import sys
from pathlib import Path

VECTOR_EXTS = {"emf", "wmf", "svg"}
# 图片集收录的格式（栅格 + 矢量；公式 bin / chart xml 等不进图片集）
GALLERY_EXTS = {"png", "jpg", "jpeg", "bmp", "gif", "webp",
                "tif", "tiff"} | VECTOR_EXTS


def _vector_size(path: Path):
    """从矢量文件头解析内容逻辑尺寸(px)。返回 (w, h) 或 None。

    - WMF：placeable 头（key=0x9AC6CDD7）bbox 在 offset 6，标准头在 offset 14，
      单位 twips（1/1440 inch）；
    - EMF：EMR_HEADER 的 rclFrame 在 offset 24，单位 0.01mm；
    - SVG：width/height 或 viewBox（按 px 折算）。
    """
    ext = path.suffix.lower()
    try:
        if ext == ".wmf":
            b = path.read_bytes()
            if len(b) >= 22 and struct.unpack("<I", b[0:4])[0] == 0x9AC6CDD7:
                x_min, y_min, x_max, y_max = struct.unpack("<hhhh", b[6:14])
            elif len(b) >= 22:
                x_min, y_min, x_max, y_max = struct.unpack("<hhhh", b[14:22])
            else:
                return None
            w, h = (x_max - x_min) * 96 / 1440, (y_max - y_min) * 96 / 1440
            return (w, h) if w > 0 and h > 0 else None
        if ext == ".emf":
            b = path.read_bytes()[:40]
            if len(b) >= 40:
                l, t, r, bt = struct.unpack("<iiii", b[24:40])
                w, h = (r - l) * 96 / 2540, (bt - t) * 96 / 2540
                return (w, h) if w > 0 and h > 0 else None
        if ext == ".svg":
            s = path.read_text("utf-8", "ignore")[:4000]
            mw = re.search(r'width="([\d.]+)(\w*)"', s)
            mh = re.search(r'height="([\d.]+)(\w*)"', s)
            if mw and mh:
                w, h = float(mw.group(1)), float(mh.group(1))
                u = mw.group(2) or "px"
                if u == "pt":
                    w, h = w * 96 / 72, h * 96 / 72
                elif u == "cm":
                    w, h = w * 96 / 2.54, h * 96 / 2.54
                elif u == "mm":
                    w, h = w * 96 / 25.4, h * 96 / 25.4
                return (w, h) if w > 0 and h > 0 else None
            mv = re.search(r'viewBox="[\d.]+ [\d.]+ ([\d.]+) ([\d.]+)"', s)
            if mv:
                w, h = float(mv.group(1)), float(mv.group(2))
                return (w, h) if w > 0 and h > 0 else None
    except Exception:
        pass
    return None


def _count_colors(path: Path, sample: int = 32):
    """缩小采样统计颜色种类；异常返回 None（不参与纯色判据）。"""
    try:
        from PIL import Image
        im = Image.open(path).convert("RGB").resize((sample, sample))
        return len(im.getcolors(sample * sample) or [])
    except Exception:
        return None


def _ink_ratio(path: Path, sample: int = 64):
    """前景（内容）像素占比：灰度 <200 像素 / 总像素；异常返回 None。

    白底字符碎片图的字符只占 3%~8%，正常教学图（框图/电路图/波形图）
    通常 >8% 且绝对面积极大——与面积判据配合可精准识别孤立符号图。
    """
    try:
        from PIL import Image
        im = Image.open(path).convert("L")
        w, h = im.size
        if not w or not h:
            return None
        sw = min(w, sample)
        sh = max(1, int(sample * h / w)) if w > sample else h
        a = im.resize((sw, sh))
        px = a.load()
        dark = 0
        for y in range(sh):
            for x in range(sw):
                if px[x, y] < 200:
                    dark += 1
        return dark / (sw * sh)
    except Exception:
        return None


def _wmf_bitmap_wrapper(path: Path) -> bool:
    """WMF 记录流检测：是否"位图封装型"（公式/符号渲染成位图嵌入 WMF）。

    判定：存在内嵌位图记录（STRETCHBLT/DIBBLT 等）且**无任何曲线记录**
    （Polyline/Arc/Ellipse）→ 说明内容是位图（公式/符号），非真矢量图。
    实测课件：位图封装型 WMF 的 polyline 恒为 0、bitblt ≥ 1（公式 WMF）；
    真矢量电路图/波形图 WMF 含 polyline/arc 曲线记录。
    """
    if path.suffix.lower() != ".wmf":
        return False
    try:
        b = path.read_bytes()
        off = 22 if len(b) >= 22 and \
            struct.unpack("<I", b[0:4])[0] == 0x9AC6CDD7 else 0
        p = off + 18
        bitblt = polyline = arc = ellipse = 0
        while p + 6 <= len(b):
            rec_size, rec_type = struct.unpack("<IH", b[p:p + 6])
            if rec_size < 3:
                break
            rb = rec_size * 2
            if p + rb > len(b):
                break
            if rec_type in (0x0A32, 0x0329, 0x032A, 0x032B):
                bitblt += 1
            elif rec_type in (0x0623, 0x0325):
                polyline += 1
            elif rec_type in (0x0628, 0x0631, 0x0633):
                arc += 1
            elif rec_type == 0x0613:
                ellipse += 1
            p += rb
        return bitblt > 0 and polyline == 0 and arc == 0 and ellipse == 0
    except Exception:
        return False


def sparse_filter(rec, by_page: Path, min_area: int = 40000,
                  max_sparse_ink: float = 0.20,
                  vec_min_area: int = 10000,
                  keep_vec_bitmap: bool = False):
    """孤立字符/碎片图判据（优先级最高，命中后跳过 YOLO 强保留）。

    栅格图判定：面积 < min_area（默认 200x200=40000px²）
               且 前景占比 < max_sparse_ink（默认 20%）。
    矢量图判定：无像素可算前景，改为解析矢量文件头的内容尺寸
               （WMF bbox/EMF frame/SVG viewBox），面积 < vec_min_area
               （默认 100x100=10000px²）即视为碎片——实测课件中碎片
               符号矢量图面积 ≤1147px²，正常矢量图 ≥28000px²，空档大；
               另：WMF 位图封装型（公式/符号渲染成位图嵌入，无曲线记录）
               直接视为公式矢量版过滤（keep_vec_bitmap=True 保留）。
    典型目标：PPT 里"插入的单个公式符号"（栅格或矢量形式）。
    说明：符号可能被 YOLO 误检（如 traffic light），故本判据先于 YOLO
          强保留执行，直接过滤。返回 reason(str) 或 None（保留）。
    """
    w, h = rec.width, rec.height
    ext = (rec.output_file or "").rsplit(".", 1)[-1].lower() \
        if rec.output_file and "." in rec.output_file else ""
    if not w or not h:
        # 矢量/未知尺寸：从头解析内容尺寸（无像素可算前景）
        if ext in VECTOR_EXTS and rec.output_file:
            p = by_page / rec.output_file
            if ext == "wmf" and not keep_vec_bitmap \
                    and _wmf_bitmap_wrapper(p):
                return "位图封装矢量图(公式/符号)"
            vs = _vector_size(p)
            if vs is not None:
                vw, vh = vs
                if vw * vh < vec_min_area:
                    return f"矢量碎片图({vw:.0f}x{vh:.0f})"
        return None
    if w * h >= min_area:
        return None
    if rec.output_file:
        ink = _ink_ratio(by_page / rec.output_file)
        if ink is not None and ink < max_sparse_ink:
            return f"孤立字符/碎片图({w}x{h}, 前景{ink:.0%})"
    return None


def rule_filter(rec, by_page: Path, min_size: int = 48,
                max_ratio: int = 10, min_colors: int = 4):
    """确定性规则层。返回 reason(str) 或 None（保留）。
    rec: ImageRecord；by_page: 图片所在目录。"""
    w, h = rec.width, rec.height
    if not w or not h:
        return None                      # 矢量/未知尺寸不按规则删
    if min(w, h) < min_size:
        return f"尺寸过小({w}x{h})"
    if max(w, h) / min(w, h) > max_ratio:
        return f"宽高比细长({w}x{h})"
    if rec.output_file:
        colors = _count_colors(by_page / rec.output_file)
        if colors is not None and colors <= min_colors:
            return f"纯色/近纯色({colors}色)"
    return None


def load_yolo(model_path="auto"):
    """加载本地 YOLO 模型（ultralytics 引擎，兼容 YOLOv5 权重）。
    model_path='auto' 时探测本地权重文件；无本地权重返回 None（不自动下载）。"""
    try:
        from ultralytics import YOLO
    except Exception:
        return None
    try:
        if model_path == "auto":
            # 优先使用随包内置权重（src/pptx_wzq/weights/），再探测外部文件
            bundled = Path(__file__).parent / "weights" / "yolov5su.pt"
            cands = ([bundled] if bundled.is_file() else []) + [
                     Path("yolov5su.pt"), Path("yolov5s.pt"), Path("yolov5n.pt"),
                     Path.home() / "weights" / "yolov5s.pt",
                     Path.home() / "weights" / "yolov5n.pt"]
            hit = next((c for c in cands if c.is_file()), None)
            if hit is None:
                return None
            model_path = str(hit)
        return YOLO(model_path)
    except Exception:
        return None


def detect_objects(model, img_path: Path, conf: float = 0.25):
    """单图检测，返回 [(class_name, conf), ...]；失败返回 []。"""
    try:
        res = model(str(img_path), conf=conf, verbose=False)[0]
        names = model.names
        out = []
        for box in (res.boxes or []):
            cls = int(box.cls[0])
            c = float(box.conf[0])
            out.append((names.get(cls, f"class{cls}"), round(c, 3)))
        return out
    except Exception:
        return []


def filter_images(records, by_page: Path, out_dir: Path,
                  min_size: int = 48, max_ratio: int = 10, min_colors: int = 4,
                  min_area: int = 40000, max_sparse_ink: float = 0.20,
                  vec_min_area: int = 10000,
                  keep_background: bool = False, keep_fill: bool = False,
                  keep_vec_bitmap: bool = False,
                  yolo_model="auto", yolo_conf: float = 0.25,
                  on_progress=None):
    """对 extract() 产物做取舍后处理。

    返回 (kept_records, filtered_entries, yolo_used)：
      - kept_records：保留的记录（note 已追加过滤/检测信息）；
      - filtered_entries：被弃明细 [{file, reason, page, shape, note}]；
      - yolo_used：本次是否实际启用了 YOLO 层。
    被弃图移动到 out_dir/discarded/，取舍明细写 out_dir/filter_report.json。

    判据优先级：
      0) kind 判据（结构背景图）→ 直接过滤，跳过 YOLO 强保留
         background：母版/布局背景图（<p:bg> 引用），keep_background=True 保留；
         fill：形状/文本框填充图（<p:sp><p:spPr><a:blipFill>），keep_fill=True 保留；
         （教学场景下两者均为装饰底图，非教学内容图）
      1) sparse_filter（孤立字符/碎片图）→ 直接过滤，跳过 YOLO 强保留
         （孤立符号常被 YOLO 误检为物体，如 traffic light）；
      2) rule_filter（尺寸/宽高比/纯色）→ 命中后 YOLO 检出物体则强保留；
      3) 其余保留。
    """
    discarded = out_dir / "discarded"
    model = load_yolo(yolo_model)
    kept, filtered = [], []
    for i, rec in enumerate(records):
        if on_progress is not None:
            try:
                on_progress(i + 1, len(records), {"kind": "filter"})
            except Exception:
                pass
        src = (by_page / rec.output_file) if rec.output_file else None
        # 判据0：结构背景图（背景/形状填充）→ 直接过滤，绕过 YOLO
        reason = None
        force_drop = False
        if rec.kind == "background" and not keep_background:
            reason = "PPT背景图(母版/布局)"
            force_drop = True
        elif rec.kind == "fill" and not keep_fill:
            reason = "形状/文本框填充图"
            force_drop = True
        # 判据1：孤立字符/碎片图（栅格：面积小+内容稀疏；矢量：头解析尺寸小）
        if reason is None:
            reason = sparse_filter(rec, by_page, min_area, max_sparse_ink,
                                   vec_min_area, keep_vec_bitmap)
            force_drop = reason is not None
        # 判据2：常规规则层
        if reason is None:
            reason = rule_filter(rec, by_page, min_size, max_ratio, min_colors)
        if reason is None:
            kept.append(rec)
            continue
        dets = []
        if model is not None and src is not None and src.is_file():
            dets = detect_objects(model, src, yolo_conf)
        if dets and not force_drop:
            # YOLO 检测到明确物体 → 强保留（照片等实物图不被误删）
            rec.note = (rec.note + "；" if rec.note else "") + \
                f"YOLO检测到: {','.join(n for n, _ in dets)}"
            kept.append(rec)
            continue
        # 确认过滤：移到 discarded/（同盘 replace，安全可审计）
        if src is not None and src.is_file():
            try:
                dst = discarded / src.name
                dst.parent.mkdir(parents=True, exist_ok=True)
                src.replace(dst)
            except Exception as e:
                filtered.append({"file": rec.output_file, "reason": reason,
                                 "page": rec.page, "shape": rec.shape_name,
                                 "note": f"移动失败({e})，未删除"})
                kept.append(rec)
                continue
        rec.note = (rec.note + "；" if rec.note else "") + f"已过滤: {reason}"
        filtered.append({"file": rec.output_file, "reason": reason,
                         "page": rec.page, "shape": rec.shape_name,
                         "kind": rec.kind, "note": rec.note})
    # 取舍明细落盘（不修改 manifest 16 列协议）
    try:
        (out_dir / "filter_report.json").write_text(
            json.dumps(filtered, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass
    return kept, filtered, model is not None


def print_env(prefix="[环境]"):
    """YOLO 环境检查摘要（stderr）。返回 (yolo_ok, reason)。"""
    try:
        from ultralytics import YOLO  # noqa
        yolo_ok = True
        reason = "ultralytics 已安装"
    except Exception:
        yolo_ok = False
        reason = "未安装 ultralytics（pip install ultralytics）"
    try:
        from PIL import Image  # noqa
        pil_ok = True
    except Exception:
        pil_ok = False
    print(f"{prefix} Pillow: {'OK' if pil_ok else '缺失（图片无法转 PNG）'}", file=sys.stderr)
    if yolo_ok:
        m = load_yolo("auto")
        if m is not None:
            print(f"{prefix} YOLO: OK（本地权重已加载）", file=sys.stderr)
        else:
            print(f"{prefix} YOLO: 引擎已装，但未找到本地权重文件"
                  "（放置 yolov5su.pt 或用 --yolo-model 指定；YOLO 层将跳过）",
                  file=sys.stderr)
    else:
        print(f"{prefix} YOLO: {reason}（YOLO 层跳过，规则层仍生效）", file=sys.stderr)
    return yolo_ok, reason


def _no_detect_desc():
    return "未检出明显目标（可能为框图/电路图/波形图等教学图形）"


def build_image_gallery(records, by_page: Path, out_dir: Path,
                        stem: str, yolo_model="auto", yolo_conf: float = 0.25,
                        on_progress=None):
    """把保留图片复制到 out/images/，并生成 out/images.md 清单。

    每张图定一个 ID（IMG0001 递增，按 页码→序号 排序），附 YOLO 简要说明：
      检测到物体 → "类别(置信度)；类别(置信度)…"；
      无检测结果 → "未检出明显目标（可能为框图/电路图/波形图等教学图形）"。
    图片为**复制**（by_page 与 manifest 16 列协议零改动，可审计溯源）。

    返回 (n_images, md_path)。
    """
    images_dir = out_dir / "images"
    images_dir.mkdir(parents=True, exist_ok=True)
    model = load_yolo(yolo_model)
    rows = []
    ordered = sorted(records, key=lambda r: (r.page, r.index))
    for i, rec in enumerate(ordered):
        if on_progress is not None:
            try:
                on_progress(i + 1, len(ordered), {"kind": "gallery"})
            except Exception:
                pass
        if not rec.output_file:
            continue
        ext = rec.output_file.rsplit(".", 1)[-1].lower() if "." in rec.output_file else ""
        if ext not in GALLERY_EXTS:
            continue                       # 公式 bin / chart xml 不进图片集
        src = by_page / rec.output_file
        if not src.is_file():
            continue
        dst = images_dir / rec.output_file
        if dst.exists():                   # 同名冲突加序号
            dst = images_dir / f"{Path(rec.output_file).stem}_{i:03d}{Path(rec.output_file).suffix}"
        try:
            shutil.copy2(src, dst)
        except Exception:
            continue
        img_id = f"IMG{len(rows) + 1:04d}"
        desc = ""
        if model is not None:
            dets = detect_objects(model, src, yolo_conf)
            desc = "；".join(f"{n}({c:.2f})" for n, c in dets) if dets \
                else _no_detect_desc()
        rows.append({
            "id": img_id, "file": dst.name, "format": ext.upper(),
            "size": f"{rec.width}x{rec.height}" if (rec.width and rec.height) else "-",
            "page": rec.page, "shape": rec.shape_name,
            "yolo": desc, "source": rec.output_file,
        })
    # 写 images.md
    md_path = out_dir / "images.md"
    lines = [f"# {stem} 图片清单", "",
             f"> 由 `pptx-img` 提取后整理（共 {len(rows)} 张）。"
             "图片本体在 `images/` 目录；ID 按 页码→序号 递增。"
             "YOLO 说明来自本地 yolov5 推理。", "",
             "| ID | 文件 | 格式 | 尺寸 | 页码 | 形状 | YOLO 说明 |",
             "|---|---|---|---|---|---|---|"]
    for r in rows:
        size = r["size"]
        lines.append(f"| {r['id']} | `{r['file']}` | {r['format']} | "
                     f"{size} | {r['page']} | {r['shape']} | {r['yolo']} |")
    lines += ["", f"图片总数 **{len(rows)}** 张",
              f"- 栅格图：{sum(1 for r in rows if r['format'] not in {e.upper() for e in VECTOR_EXTS})} 张",
              f"- 矢量图：{sum(1 for r in rows if r['format'] in {e.upper() for e in VECTOR_EXTS})} 张"]
    md_path.write_text("\n".join(lines), encoding="utf-8")
    return len(rows), str(md_path)
