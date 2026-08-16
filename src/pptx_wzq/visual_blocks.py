#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
visual_blocks.py — PPT 页面的可视逻辑块（Visual Block）全栈解析核心库
======================================================================

把 extract_pptx_images.extract() 产出的原子对象（像素图/矢量图/Visio/图表/
原生 shape/连接符/表格）按空间邻近规则聚成 1~6 个可视逻辑块，并输出：

  - block_type：raster_image / vector_image / visio_diagram / chart /
    native_table / virtual_table / native_logic_diagram / decorated_diagram /
    formula_block / title_region / body_text_region
  - bbox / center / z_index_range：几何
  - assets：渲染图（images/）+ 矢量源（sources/）+ 内部资源
  - internal_structure：节点/边（图或树）或表格矩阵
  - semantic_description：expression_goal / expression_role /
    expression_features / vlm_caption / teaching_use（VLM 生成）
  - cross_modal_relations：块与页面文字的关系（VLM/规则生成）

设计原则（对齐 pptx-kb 六项需求改造的向后兼容约束）：
  1. 纯本地可跑：空间聚类、拓扑推断、资源归档不需要任何模型；
  2. VLM 是增强层：describe_block() 需要 client，缺失时回退规则模板；
  3. 不改变 extract_pptx_images 的既有输出（manifest/images/by_page）。

作者：吴振谦 · wuzhenqian@nbu.edu.cn
"""
from __future__ import annotations

import json
import re
import sys
import time
from dataclasses import dataclass, field, asdict
from pathlib import Path

# --------------------------------------------------------------------------
# 常量与配置
# --------------------------------------------------------------------------
BLOCK_TYPE_ALL = {
    "raster_image", "vector_image", "visio_diagram", "chart",
    "native_table", "virtual_table", "native_logic_diagram",
    "decorated_diagram", "formula_block", "title_region", "body_text_region",
}

# 空间聚类参数（方案 §5.1）
CLUSTER_CONFIG = {
    "min_area": 1600,          # 40x40 px，小于此视为噪声（但保留文本对象）
    "alpha_h": 1.0,            # 水平邻近系数（原 1.5，链式误并过多）
    "alpha_v": 1.0,            # 垂直邻近系数（原 1.5）
    "max_gap_px": 40,          # 绝对最大间距（原 60，收紧防远距串并）
    "text_gap_px": 100,        # 双方都有文本（短标签）时的放宽间距
                               # （表格单元格/框图节点倾向于同属一个结构）
    "max_blocks_per_slide": 6, # 一页最多输出块数
    "page_w": 960,             # 幻灯片标准宽 px（16:9 参考，用于越界过滤）
    "page_h": 720,             # 幻灯片标准高 px
    # 文本密度判据（用户准则：可视逻辑块内文本框应为短标签）
    "max_shape_text": 10,      # 单个 shape 文本 >10 字 → 视为潜在正文文本框，
                               # 不参与聚类（逻辑图节点/标注通常字少）
    "max_block_text": 30,      # 聚类后块内文本总量 >30 字 → 判定混入了文本区，
                               # 剔除长文本成员并重组块
    # 装饰/符号过滤判据（吸收 img_filter 教学判据思路）
    "min_size": 48,            # 无文本 shape 宽或高 <48px → 装饰小图标
    "max_ratio": 10,           # 无文本 shape 宽高比 >10 → 装饰细长线
    "formula_min_area": 20000, # 小型公式面积 <20000px²（约141×141）→ 公式符号
    # 用户准则：大面积对象独立 + 凸区域分割
    "big_area_ratio": 0.30,    # raster/vector/visio 面积 > 整页 30% → 独立成块
    "split_gap_px": 40,        # 凸分割空白走廊阈值：贯穿块 bbox 的空隙 > 该值 → 切分
    "min_sub_area_ratio": 0.20,  # 凸分割切出的子块面积 < 页面 20% → 不切
                                  # （用户准则：面积过小的区域不单独成块）
    # 页眉/页脚横幅过滤（战略管理等课件常见的跨页重复装饰色带）
    "banner_ratio": 0.9,       # 宽度 ≥ 90% 页面宽
    "banner_max_h": 80,        # 且高度 ≤ 80px → 判为装饰横幅
    # 符号碎片过滤（用户准则）：无文本、无媒体、无连接的纯形状块不构成
    # 可视逻辑块。对象数 ≤3 无条件剔除；对象数 >3 但块面积 < 页面
    # min_block_area_ratio（10%≈263×263px）也剔除。含文本标签/图/表/
    # 连接的块不受此限制（逻辑图节点短标签、独立小图均保留）
    "min_block_area_ratio": 0.10,
}

# 单对象块直接映射（不调 VLM）
SINGLE_KIND_TO_TYPE = {
    "raster": "raster_image",
    "vector": "vector_image",
    "visio": "visio_diagram",
    "chart": "chart",
    "formula": "formula_block",
    "table": "native_table",
}

# VLM 提示词（合并块类型识别 + Semantic Captioning，方案 §3.4 Step4）
DESCRIBE_SYSTEM = (
    "你是高校教材建设专家与可视化信息解析专家。下面给出一张从 PPT 页面裁剪出的"
    "可视区域图片（可能是一个独立图片、一个表格、一组箭头/形状/文本框组成的逻辑图，"
    "或像素图与标注叠加的复合图）以及该页的文字上下文。请输出 JSON：\n"
    "{\n"
    "  \"block_type\": \"raster_image|vector_image|visio_diagram|chart|native_table|"
    "virtual_table|native_logic_diagram|decorated_diagram|formula_block|title_region|"
    "body_text_region\",\n"
    "  \"expression_goal\": \"该可视化要表达的教学主题/结论，一句话（≤40字）\",\n"
    "  \"expression_role\": \"它如何帮助理解该页文字（如具象化/对比/流程演示/数据呈现），"
    "约50字（≤70字）\",\n"
    "  \"expression_features\": [\"时序图\", \"流程图\", \"层次结构\", \"网格表格\" 等抽象特征，"
    "2~5个词],\n"
    "  \"vlm_caption\": \"对图片内容的客观描述，100~200字\",\n"
    "  \"teaching_use\": \"适用于什么课程/教学场景，30~60字\"\n"
    "}\n"
    "只输出 JSON，不要任何前缀或解释。"
)

RELATION_SYSTEM = (
    "你是高校教材建设专家。下面给出一页 PPT 的正文文本和一个可视逻辑块的描述。"
    "请输出该可视逻辑块与页面文字的关系，JSON 格式：\n"
    "{\n"
    "  \"relation_type\": \"title_caption|elaboration|example|summary|definition|contrast\",\n"
    "  \"text_anchor\": \"正文中最相关的一句话（原样引用，≤60字）\",\n"
    "  \"semantic_link\": \"块与文字的逻辑关系陈述，约50字（≤60字）\"\n"
    "}\n"
    "只输出 JSON，不要任何前缀或解释。"
)


# --------------------------------------------------------------------------
# 数据结构
# --------------------------------------------------------------------------
@dataclass
class AtomicObject:
    """原子对象（与 extract_pptx_images._collect_atomic_objects 的 dict 同构）。"""
    obj_id: str
    page: int
    kind: str                # raster/vector/visio/chart/formula/shape/connector/table
    shape_name: str = ""
    text: str = ""
    bbox: dict = field(default_factory=lambda: {"x": 0, "y": 0, "w": 0, "h": 0})
    z_index: int = 0
    source_media: str = ""
    output_file: str = ""
    original_format: str = ""
    children: list = field(default_factory=list)

    @classmethod
    def from_dict(cls, d: dict) -> "AtomicObject":
        return cls(
            obj_id=d.get("obj_id", ""),
            page=int(d.get("page", 0)),
            kind=d.get("kind", "shape"),
            shape_name=d.get("shape_name", ""),
            text=d.get("text", ""),
            bbox=d.get("bbox") or {"x": 0, "y": 0, "w": 0, "h": 0},
            z_index=int(d.get("z_index", 0)),
            source_media=d.get("source_media", ""),
            output_file=d.get("output_file", ""),
            original_format=d.get("original_format", ""),
            children=d.get("children") or [],
        )


@dataclass
class VisualBlock:
    block_id: str
    block_type: str = "native_logic_diagram"
    bbox: dict = field(default_factory=lambda: {"x": 0, "y": 0, "w": 0, "h": 0})
    z_index_range: list = field(default_factory=list)
    assets: dict = field(default_factory=dict)
    internal_structure: dict = field(default_factory=dict)
    semantic_description: dict = field(default_factory=dict)
    member_obj_ids: list = field(default_factory=list)
    page: int = 0
    is_single: bool = False
    text_density: float = 0.0   # 文字空间密度：块内文本字符数 / 块面积(px²)
                                # 高密度=文本区，低密度=逻辑图/图（判据）

    @property
    def center(self) -> dict:
        b = self.bbox
        return {"x": b["x"] + b["w"] / 2, "y": b["y"] + b["h"] / 2}

    def to_dict(self) -> dict:
        d = asdict(self)
        d.pop("member_obj_ids", None)
        d.pop("is_single", None)
        d["center"] = self.center
        return d


# --------------------------------------------------------------------------
# 空间聚类（并查集）
# --------------------------------------------------------------------------
def _union_find_cluster(objects: list[AtomicObject],
                        config: dict = None) -> list[list[AtomicObject]]:
    """按 bbox 邻近规则做并查集聚类，返回簇列表（每簇为 AtomicObject 列表）。"""
    cfg = {**CLUSTER_CONFIG, **(config or {})}
    n = len(objects)
    parent = list(range(n))

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[ra] = rb

    def bbox_of(o: AtomicObject) -> dict:
        return o.bbox or {"x": 0, "y": 0, "w": 0, "h": 0}

    for i in range(n):
        bi = bbox_of(objects[i])
        ci = (bi["x"] + bi["w"] / 2, bi["y"] + bi["h"] / 2)
        for j in range(i + 1, n):
            bj = bbox_of(objects[j])
            cj = (bj["x"] + bj["w"] / 2, bj["y"] + bj["h"] / 2)
            # 边缘间隙（≥0）：中心距减去两对象尺寸的一半
            gap_x = max(0.0, abs(ci[0] - cj[0]) - (bi["w"] + bj["w"]) / 2)
            gap_y = max(0.0, abs(ci[1] - cj[1]) - (bi["h"] + bj["h"]) / 2)
            # 恒定间隙阈值（不受对象高度放大，防大对象链式串并）
            gap_max = cfg["max_gap_px"]
            h_close = gap_x <= gap_max      # 水平间隙小
            v_close = gap_y <= gap_max      # 垂直间隙小
            h_ov = gap_x == 0               # 水平有重叠（含包含）
            v_ov = gap_y == 0               # 垂直有重叠（含包含）
            # 合并条件：双向重叠 或 一个方向重叠+另一方向间隙小
            if (h_ov and v_ov) or (h_ov and v_close) or (v_ov and h_close):
                union(i, j)

    groups = {}
    for i in range(n):
        groups.setdefault(find(i), []).append(objects[i])
    return list(groups.values())


def _split_cluster_convex(cluster: list, gap_px: int = 40,
                          page_w: int = 960, page_h: int = 720,
                          min_sub_area: float = None) -> list:
    """凸区域递归分割（用户准则）：删除多字文本框后，剩余连续区域若是凸的
    则为一个可视逻辑块；若是凹的，按最少凸分割拆成若干块。

    切割判定（修复：网格结构不再被误切）：
    1) 规则网格检测：成员中心点行列对齐（≥2列×≥2行且填充率≥60%）→
       判为 virtual_table（虚拟表格），跳过凸分割——表格的行列间距在
       投影上就是空隙，若按空隙切会把表格切成碎片；
    2) 有 connector 的簇不切（显式关联的逻辑图）；
    3) 真正的"贯穿走廊"：空隙带两侧的成员必须"横跨"空隙（即空隙
       不是单列/单行的内部间距），且空隙宽度 > gap_px；
    4) 用户规则：切出的任一子块 bbox 面积 < min_sub_area（默认
       页面 20%）→ 该切割线无效（不切），防止切出过小碎片块。
    """
    if len(cluster) < 2:
        return [cluster]
    if min_sub_area is None:
        min_sub_area = page_w * page_h * 0.20   # 用户准则：<20% 页面的块不切
    # 规则网格 → 虚拟表格，不切
    if _is_grid_table(cluster):
        return [cluster]
    # 有 connector 关联的簇视为显式逻辑图，不切
    if any(o.kind == "connector" for o in cluster):
        return [cluster]

    min_x = min(o.bbox["x"] for o in cluster)
    min_y = min(o.bbox["y"] for o in cluster)
    max_x = max(o.bbox["x"] + o.bbox["w"] for o in cluster)
    max_y = max(o.bbox["y"] + o.bbox["h"] for o in cluster)
    bw, bh = max_x - min_x, max_y - min_y

    # --- 水平切割：找贯穿的垂直空白带（x 空隙把成员分左右）---
    for cut in _find_gaps(
            [(o.bbox["x"], o.bbox["x"] + o.bbox["w"]) for o in cluster],
            bw, gap_px):
        left = [o for o in cluster if o.bbox["x"] + o.bbox["w"] <= cut]
        right = [o for o in cluster if o.bbox["x"] >= cut]
        if left and right and _cut_ok(left, right, cut, min_sub_area):
            return (_split_cluster_convex(left, gap_px, page_w, page_h,
                                          min_sub_area) +
                    _split_cluster_convex(right, gap_px, page_w, page_h,
                                          min_sub_area))
    # --- 垂直切割：找贯穿的水平空白带（y 空隙把成员分上下）---
    for cut in _find_gaps(
            [(o.bbox["y"], o.bbox["y"] + o.bbox["h"]) for o in cluster],
            bh, gap_px):
        top = [o for o in cluster if o.bbox["y"] + o.bbox["h"] <= cut]
        bottom = [o for o in cluster if o.bbox["y"] >= cut]
        if top and bottom and _cut_ok(top, bottom, cut, min_sub_area):
            return (_split_cluster_convex(top, gap_px, page_w, page_h,
                                          min_sub_area) +
                    _split_cluster_convex(bottom, gap_px, page_w, page_h,
                                          min_sub_area))
    return [cluster]


def _cut_ok(grp_a: list, grp_b: list, cut: float,
            min_sub_area: float) -> bool:
    """切割合法性：两侧子块 bbox 面积都 ≥ min_sub_area（防切碎片）。"""
    for grp in (grp_a, grp_b):
        gx0 = min(o.bbox["x"] for o in grp)
        gy0 = min(o.bbox["y"] for o in grp)
        gx1 = max(o.bbox["x"] + o.bbox["w"] for o in grp)
        gy1 = max(o.bbox["y"] + o.bbox["h"] for o in grp)
        if (gx1 - gx0) * (gy1 - gy0) < min_sub_area:
            return False
    return True


def _is_grid_table(cluster: list) -> bool:
    """规则网格检测：成员中心点在行列方向都规律对齐 → 虚拟表格。
    条件：成员 ≥4；按中心点一维聚类出 ≥2 列 且 ≥2 行；
    成员数 ≥ 行列组合的 60%（允许少量空格）。"""
    if len(cluster) < 4:
        return False
    xs = sorted(round(o.bbox["x"] + o.bbox["w"] / 2) for o in cluster)
    ys = sorted(round(o.bbox["y"] + o.bbox["h"] / 2) for o in cluster)
    cols = _cluster_1d(xs, tol=60)
    rows = _cluster_1d(ys, tol=50)
    if len(cols) >= 2 and len(rows) >= 2:
        if len(cluster) >= len(cols) * len(rows) * 0.6:
            return True
    return False


def _cluster_1d(vals: list, tol: float) -> list:
    """一维聚类：相邻值差 ≤ tol 合并为同一组，返回组中心列表。"""
    groups = []
    for v in vals:
        if groups and v - groups[-1][-1] <= tol:
            groups[-1].append(v)
        else:
            groups.append([v])
    return [sum(g) / len(g) for g in groups]


def _find_gaps(intervals: list, total: float, gap_px: float) -> list:
    """在 [0,total] 区间内找"贯穿空隙"的切割位置。
    intervals: [(start, end)] 成员的投影区间；返回可切割的坐标列表
    （空隙中点），要求空隙宽度 > gap_px 且两端都有成员覆盖。"""
    if not intervals:
        return []
    # 把成员区间按 start 排序，检查相邻成员之间的空隙是否"贯穿"（无人覆盖）
    ivs = sorted((max(0.0, a), min(total, b)) for a, b in intervals)
    cuts = []
    prev_end = None
    for a, b in ivs:
        if prev_end is not None and a - prev_end > gap_px:
            cuts.append((prev_end + a) / 2)
        prev_end = max(prev_end or 0, b)
    return cuts


def _split_big_objects(objs: list, cfg: dict) -> tuple:
    """用户准则：raster/vector/visio 面积 > 整页 30% → 独立成块（不合并）。
    返回 (big_objs 列表, 其余对象列表)。"""
    pw, ph = cfg.get("page_w", 960), cfg.get("page_h", 720)
    ratio = cfg.get("big_area_ratio", 0.30)
    page_area = pw * ph * ratio
    big, rest = [], []
    for o in objs:
        a = (o.bbox or {}).get("w", 0) * (o.bbox or {}).get("h", 0)
        if o.kind in ("raster", "vector", "visio") and a >= page_area:
            big.append(o)
        else:
            rest.append(o)
    return big, rest


def _is_adjacent(a: AtomicObject, b: AtomicObject,
                 gap_max: float, text_gap: float = None) -> bool:
    """四向邻接判定：a、b 是否可生长连接。
    规则（方向区分）：
    - 左右相邻：水平间隙 ≤ 阈值 且 垂直方向有重叠/接近；
    - 上下相邻：垂直间隙 ≤ 阈值 且 水平方向有重叠/接近；
    阈值：双方都有文本（短标签，如表格单元格/框图节点）时放宽为
    text_gap（默认 gap_max），因为这类对象倾向于同属一个结构
    （表格/流程图），间距可能大于纯图形的 40px。"""
    if text_gap is None:
        text_gap = gap_max
    thr = text_gap if (a.text.strip() and b.text.strip()) else gap_max
    ax0, ay0 = a.bbox["x"], a.bbox["y"]
    ax1, ay1 = ax0 + a.bbox["w"], ay0 + a.bbox["h"]
    bx0, by0 = b.bbox["x"], b.bbox["y"]
    bx1, by1 = bx0 + b.bbox["w"], by0 + b.bbox["h"]
    gap_x = max(ax0, bx0) - min(ax1, bx1)      # >0 分离，≤0 重叠
    gap_y = max(ay0, by0) - min(ay1, by1)
    # 左右相邻：水平间隙小 且 垂直有重叠或接近
    if gap_x <= thr and gap_y <= gap_max:
        return True
    # 上下相邻：垂直间隙小 且 水平有重叠或接近
    if gap_y <= thr and gap_x <= gap_max:
        return True
    return False


def _region_grow(objects: list, cfg: dict) -> list:
    """四向种子扩展区域生长（用户准则：从种子对象出发，沿四个方向扩展，
    遇到"字多文本对象（墙）"或页面边界即停止）。

    与"全局聚类+凸分割"的本质区别：
    - 加法生长：从每个未访问种子开始 BFS 四向扩展，把紧邻对象并入同一
      区域；表格/框图单元格四向紧邻 → 自然长成一片，不会被切碎；
    - 文本墙：>max_shape_text 字的对象是"墙"，既不作为种子也不被并入
      （生长碰到它即停止），凹区域天然被墙/间隙阻断成多个区域，
      无需凸分割；
    - 每页对象多时按空间顺序取种子，避免重复遍历。
    """
    max_shape_text = cfg.get("max_shape_text", 10)
    gap_max = cfg.get("max_gap_px", 40)
    text_gap = cfg.get("text_gap_px", gap_max * 2.5)
    walls = {o.obj_id for o in objects
             if len((o.text or "").strip()) > max_shape_text}
    members = [o for o in objects if o.obj_id not in walls]
    # 种子按 z_index 排序（从底层对象开始生长，稳定输出顺序）
    members.sort(key=lambda o: o.z_index)

    regions = []
    visited = set()
    for seed in members:
        if seed.obj_id in visited:
            continue
        region = []
        stack = [seed]
        visited.add(seed.obj_id)
        while stack:
            cur = stack.pop()
            region.append(cur)
            for other in members:
                if other.obj_id in visited or other.obj_id in walls:
                    continue
                if _is_adjacent(cur, other, gap_max, text_gap):
                    visited.add(other.obj_id)
                    stack.append(other)
        regions.append(region)
    return regions


def _filter_noise(objects: list[AtomicObject],
                  config: dict = None) -> list[AtomicObject]:
    """过滤噪声（吸收 img_filter 教学判据思路）：

    1) 面积过小且无文本、无输出文件的装饰对象 → 丢弃；
    2) 小型公式（formula 面积 < formula_min_area）→ 丢弃
       （单个公式符号/上下标，内容已由公式提取步骤保留）；
    3) 无文本 shape：宽或高 < min_size（装饰小图标）或
       宽高比 > max_ratio（装饰细长线）→ 丢弃；
    4) 整页宽横幅（raster/vector 宽度 > 页面宽 banner_ratio，且
       高度 < banner_max_h，如页眉/页脚装饰色带）→ 丢弃；
    有文本的 shape/connector（逻辑图节点标签，即使小）与
    chart/table（主内容）始终保留。
    """
    cfg = {**CLUSTER_CONFIG, **(config or {})}
    min_size = cfg.get("min_size", 48)
    max_ratio = cfg.get("max_ratio", 10)
    formula_min_area = cfg.get("formula_min_area", 20000)
    pw = cfg.get("page_w", 960)
    banner_ratio = cfg.get("banner_ratio", 0.9)
    banner_max_h = cfg.get("banner_max_h", 80)
    out = []
    for o in objects:
        b = o.bbox or {}
        w, h = b.get("w", 0), b.get("h", 0)
        area = w * h
        text = (o.text or "").strip()
        # 方案A：公式（OLE/OMML）一律不参与可视逻辑块——公式属于文本流
        # 内容，已由公式提取步骤（formulas.md）完整保留；且公式对象普遍
        # 带 output_file（提取落盘产物），若不在此前置剔除会绕过下方
        # 面积过滤被当作"主内容"保留，导致公式区域被圈进可视逻辑块。
        if o.kind == "formula":
            continue
        # 有文本 → 保留（逻辑图节点/箭头标注，即使小）
        if text:
            out.append(o)
            continue
        # 整页宽横幅：宽度 ≥ 90% 页面宽 且 高度 ≤ 80px（页眉/页脚色带）
        if o.kind in ("raster", "vector") and w >= pw * banner_ratio \
                and 0 < h <= banner_max_h:
            continue
        # 有媒体文件（raster/vector/visio/chart）→ 主内容，保留
        if o.output_file:
            out.append(o)
            continue
        # 表格 → 保留
        if o.kind == "table":
            out.append(o)
            continue
        # 小型公式（无文本且面积不足）→ 丢弃（符号/上下标）
        if o.kind == "formula":
            if area >= formula_min_area:
                out.append(o)
            continue
        # 无文本 shape/connector：尺寸过小 / 细长装饰 → 丢弃
        if area < cfg["min_area"] or \
                min(w, h) < min_size or \
                (min(w, h) > 0 and max(w, h) / min(w, h) > max_ratio):
            continue
        out.append(o)
    return out


# --------------------------------------------------------------------------
# 拓扑推断
# --------------------------------------------------------------------------
def _infer_topology(members: list[AtomicObject]) -> dict:
    """生成块内部图/树结构或表格矩阵。

    - 含表格对象 → 直接取 rows 作为 table_matrix；
    - 含 connector 且能解析 start/end 或几何最近邻 → nodes + edges；
    - 否则仅 nodes（无边的孤立节点集合）。
    """
    nodes, edges = [], []
    table_matrix = None
    formula_list = []
    id2obj = {o.obj_id: o for o in members}

    # 1) 节点
    for o in members:
        b = o.bbox or {"x": 0, "y": 0, "w": 0, "h": 0}
        if o.kind == "connector":
            continue
        if o.kind == "formula":
            formula_list.append(o.text or o.output_file or o.obj_id)
            continue
        bbox_norm = None
        if b["w"] > 0 and b["h"] > 0:
            bbox_norm = [round(b["x"] / max(1, b["w"]), 3),
                         round(b["y"] / max(1, b["h"]), 3),
                         round((b["x"] + b["w"]) / max(1, b["w"]), 3),
                         round((b["y"] + b["h"]) / max(1, b["h"]), 3)]
        nodes.append({
            "id": o.obj_id,
            "type": o.kind,
            "text": (o.text or "")[:200],
            "bbox_norm": bbox_norm,
            "label": o.shape_name or o.kind,
        })
    node_ids = {n["id"] for n in nodes}

    # 2) 边：connector 显式连接优先
    for o in members:
        if o.kind != "connector":
            continue
        s_id = getattr(o, "children", None)
        start_id = getattr(o, "start_shape_id", "")
        end_id = getattr(o, "end_shape_id", "")
        # connector 的 stCxn/endCxn id 是 shape id（cNvPr id），不一定等于 obj_id；
        # 兜底：用几何最近邻找起点/终点
        if start_id and end_id:
            src = _find_node_by_cxn_id(start_id, nodes, id2obj, members)
            dst = _find_node_by_cxn_id(end_id, nodes, id2obj, members)
        else:
            src, dst = _nearest_two_nodes(o, nodes, members)
        if src and dst and src != dst:
            edges.append({
                "source": src, "target": dst,
                "label": (o.text or "")[:100] or "",
                "line_style": "solid_arrow",
            })
    # 3) 无 connector 时：空间最近邻启发式（右/下优先）
    if not edges and len(nodes) >= 2:
        edges = _heuristic_edges(nodes, members)

    # 4) 表格矩阵
    for o in members:
        if o.kind == "table" and o.children:
            table_matrix = [[str(c) for c in row] for row in o.children]
            break

    # 5) topology_type
    if table_matrix is not None:
        topo = "grid_matrix"
    elif edges and _is_tree(nodes, edges):
        topo = "tree"
    elif edges:
        topo = "directed_graph"
    else:
        topo = "isolated_nodes"

    return {
        "topology_type": topo,
        "nodes": nodes,
        "edges": edges,
        "table_matrix": table_matrix,
        "formula_list": formula_list,
    }


def _find_node_by_cxn_id(cxn_id, nodes, id2obj, members):
    """cxn 的 stCxn/endCxn id 是 <a:cNvPr id>，与 obj_id 不同。
    尝试在 node label/shape_name 或 bbox 中心匹配；最坏回退最近邻。"""
    for n in nodes:
        if n["id"] == cxn_id:
            return n["id"]
    # cNvPr id 与 z_index 无关；此处回退最近邻
    return None


def _nearest_two_nodes(connector: AtomicObject, nodes, members):
    """对 connector，找其 bbox 中心两侧最近的节点。"""
    c = (connector.bbox["x"] + connector.bbox["w"] / 2,
         connector.bbox["y"] + connector.bbox["h"] / 2)
    dists = []
    for n in nodes:
        m = next((m for m in members if m.obj_id == n["id"]), None)
        if m is None:
            continue
        b = m.bbox
        center = (b["x"] + b["w"] / 2, b["y"] + b["h"] / 2)
        d = (c[0] - center[0]) ** 2 + (c[1] - center[1]) ** 2
        dists.append((d, n["id"]))
    dists.sort()
    if len(dists) >= 2:
        return dists[0][1], dists[1][1]
    return None, None


def _heuristic_edges(nodes, members):
    """无 connector 时的空间最近邻：每个节点连接其右侧/下方最近节点（阈值内）。"""
    edges = []
    centers = {}
    for n in nodes:
        m = next((m for m in members if m.obj_id == n["id"]), None)
        if m is None:
            continue
        b = m.bbox
        centers[n["id"]] = (b["x"] + b["w"] / 2, b["y"] + b["h"] / 2)
    for n in nodes:
        cx, cy = centers.get(n["id"], (0, 0))
        best, best_d = None, None
        for oid, (ox, oy) in centers.items():
            if oid == n["id"]:
                continue
            dx, dy = ox - cx, oy - cy
            # 右侧或下方（允许轻微上偏/左偏）
            if dx > -10 and dy > -10 and (dx > 0 or dy > 0):
                d = dx * dx + dy * dy
                if best_d is None or d < best_d:
                    best, best_d = oid, d
        if best and best_d is not None and best_d < 400 * 400:
            edges.append({"source": n["id"], "target": best,
                          "label": "", "line_style": "solid_arrow"})
    return edges


def _is_tree(nodes, edges):
    """判断图是否为树：边数 = 节点数-1 且无环（并查集）。"""
    if len(edges) != len(nodes) - 1:
        return False
    parent = {n["id"]: n["id"] for n in nodes}

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    for e in edges:
        a, b = find(e["source"]), find(e["target"])
        if a == b:
            return False
        parent[a] = b
    return True


# --------------------------------------------------------------------------
# 块类型判定（规则 + 可选 VLM）
# --------------------------------------------------------------------------
def _guess_block_type(members: list[AtomicObject]) -> str:
    """无 VLM 时的规则判定。"""
    kinds = [m.kind for m in members]
    # 单对象
    if len(members) == 1:
        return SINGLE_KIND_TO_TYPE.get(kinds[0], "raster_image")
    # 表格
    if "table" in kinds:
        return "native_table"
    # 全 shape/文本框/connector（无媒体）→ 原生逻辑图
    if all(k in ("shape", "connector", "textbox") for k in kinds):
        return "native_logic_diagram"
    # 有图片/矢量 + 有 shape/connector → 复合装饰图
    has_media = any(k in ("raster", "vector", "visio") for k in kinds)
    has_shape = any(k in ("shape", "connector", "textbox") for k in kinds)
    if has_media and has_shape:
        return "decorated_diagram"
    if has_media:
        return "raster_image"
    return "native_logic_diagram"


# --------------------------------------------------------------------------
# 块级渲染与资源
# --------------------------------------------------------------------------
def render_block_to_png(slide_page: int, block: VisualBlock,
                        pptx_path: str, out_png: Path, dpi: int = 150) -> bool:
    """用整页渲染 + bbox 裁剪渲染单块为 PNG（复用 extract_pptx_images 的
    render_pptx_pages / crop_page_png）。返回是否成功。
    自动读取真实页面尺寸（EMU）传入裁剪，避免 16:9 页面左侧被切。"""
    try:
        from pptx_wzq import extract_pptx_images as E
        cache = Path(pptx_path).parent / ".render_cache"
        pages = E.render_pptx_pages(str(pptx_path), cache, dpi=dpi)
        if not pages or (slide_page - 1) >= len(pages):
            return False
        sld_cx, sld_cy = E.read_sld_size(str(pptx_path))
        xfrm = (int(block.bbox["x"] / 96 * 914400),
                int(block.bbox["y"] / 96 * 914400),
                int(block.bbox["w"] / 96 * 914400),
                int(block.bbox["h"] / 96 * 914400))
        return E.crop_page_png(pages[slide_page - 1], xfrm, dpi, out_png,
                               sld_cx=sld_cx, sld_cy=sld_cy)
    except Exception as e:
        print(f"[渲染] 块 {block.block_id} 渲染失败：{e}", file=sys.stderr)
        return False


def export_block_vector(block: VisualBlock, by_page: Path, sources_dir: Path,
                        prefix: str) -> str | None:
    """把块内矢量/Visio 源文件复制到 sources/（统一资源目录）。
    返回复制后的相对路径或 None。"""
    vec_exts = {"vsdx", "vsd", "svg", "wmf", "emf"}
    for m in block.member_obj_ids:
        # member 的 output_file 需由调用方注入；此处从 block.assets 取
        pass
    # assets 里可能有 vector_svg 或 internal_resources；由 cli_blocks 组装
    return None


# --------------------------------------------------------------------------
# VLM：块描述 + 跨模态关系
# --------------------------------------------------------------------------
def _safe_json(text: str) -> dict:
    m = re.search(r"\{.*\}", text, re.S)
    if not m:
        return {}
    try:
        return json.loads(m.group(0))
    except Exception:
        return {}


def describe_block(client, model: str, block: VisualBlock,
                   page_text: str, image_path: Path = None) -> dict:
    """调用 VLM 生成 block_type + semantic_description；无 client 或失败时
    回退规则模板。image_path 存在时以多模态方式传入。"""
    if client is None:
        return _fallback_description(block)
    try:
        content = [{"type": "text",
                    "text": f"【页面文字上下文】\n{page_text[:1200]}\n\n"
                            f"请分析这张可视区域图片。"}]
        if image_path and image_path.is_file():
            import base64
            b64 = base64.b64encode(image_path.read_bytes()).decode()
            ext = image_path.suffix.lower().lstrip(".") or "png"
            mime = "image/png" if ext == "png" else f"image/{ext}"
            content.insert(0, {"type": "image_url",
                               "image_url": {"url": f"data:{mime};base64,{b64}"}})
        resp = client.chat.completions.create(
            model=model,
            messages=[{"role": "system", "content": DESCRIBE_SYSTEM},
                      {"role": "user", "content": content}],
            stream=False,
            reasoning_effort="high",
            extra_body={"thinking": {"type": "enabled"}},
        )
        data = _safe_json(resp.choices[0].message.content or "")
        btype = data.get("block_type", "")
        if btype not in BLOCK_TYPE_ALL:
            btype = _guess_block_type(
                [AtomicObject.from_dict({"obj_id": i, "page": block.page,
                                         "kind": "shape", "bbox": block.bbox})
                 for i in block.member_obj_ids])
        return {
            "block_type": btype,
            "expression_goal": str(data.get("expression_goal", "") or "")[:80],
            "expression_role": str(data.get("expression_role", "") or "")[:140],
            "expression_features": [str(x)[:20] for x in
                                    (data.get("expression_features") or [])][:5],
            "vlm_caption": str(data.get("vlm_caption", "") or "")[:500],
            "teaching_use": str(data.get("teaching_use", "") or "")[:120],
        }
    except Exception as e:
        print(f"[VLM] 块描述失败：{e}，回退规则模板", file=sys.stderr)
        return _fallback_description(block)


def _fallback_description(block: VisualBlock) -> dict:
    """无 VLM 时的规则模板（保证 JSON 结构完整）。"""
    members_text = []
    st = block.internal_structure or {}
    for n in (st.get("nodes") or [])[:3]:
        if n.get("text"):
            members_text.append(n["text"])
    brief = "；".join(members_text[:3])[:80]
    return {
        "block_type": _guess_block_type(
            [AtomicObject.from_dict({"obj_id": i, "page": block.page,
                                     "kind": "shape", "bbox": block.bbox})
             for i in block.member_obj_ids]),
        "expression_goal": f"展示本页与「{brief}」相关的可视化内容" if brief
        else "展示本页的可视化内容",
        "expression_role": "通过可视化方式辅助理解本页知识点（规则回退）",
        "expression_features": ["可视化"],
        "vlm_caption": f"本块包含 {len(block.member_obj_ids)} 个对象"
                       f"{'：' + brief if brief else ''}。",
        "teaching_use": "教学辅助图示",
    }


# DeepSeek 语义增强提示词（用户要求：图文绑定步骤用 LLM 生成/精炼
# semantic_description，原料=块结构化信息 + 该页文本/公式 + 已有描述）
SEMANTIC_SYSTEM = (
    "你是高校教材建设专家。下面给出一个 PPT 页面中的一个\"可视逻辑块\""
    "（由像素图/矢量图/形状/文本框/箭头组成的、表达一个主题逻辑的可视化"
    "对象集合）的结构化信息：块类型、内部节点文本、占据页面比例、该页文字、"
    "公式，以及该块已有的描述（如有）。请从教材角度输出该块的完整语义描述 "
    "JSON：\n"
    "{\n"
    "  \"expression_goal\": \"该块要表达的教学主题/结论，一句话（≤40字）\",\n"
    "  \"expression_role\": \"它如何帮助理解该页文字（如具象化/对比/流程演示/"
    "数据呈现/电路示意），约50字（≤70字）\",\n"
    "  \"expression_features\": [\"时序图\",\"流程图\",\"层次结构\",\"网格表格\","
    "\"电路图\" 等抽象表达特征，2~5个词],\n"
    "  \"vlm_caption\": \"对该块内容的教学性客观描述，100~200字\",\n"
    "  \"teaching_use\": \"适用于什么课程/教学场景，30~60字\"\n"
    "}\n"
    "只输出 JSON，不要任何前缀或解释。"
)


def enrich_semantics(client, model: str, slides: list, page_texts: dict,
                     page_formulas: dict,
                     on_progress=None) -> int:
    """用 DeepSeek（文本模型）为每个可视逻辑块的 semantic_description 生成
    真实语义内容（expression_goal / expression_role / expression_features /
    vlm_caption / teaching_use），覆盖规则回退模板。

    输入原料（不依赖视觉模型）：
      - 块类型、内部节点文本摘要、bbox 面积占比；
      - 该页文字（textual_content.raw_text）与该页公式；
      - 已有 semantic_description（若是 qwen 视觉生成的可直接精炼）。

    slides：_assemble_slides 后的结构 [{slide_info, textual_content,
    visual_blocks: [dict], ...}]——直接就地更新块的 semantic_description。
    无 client / 调用失败 / 模型返回非法时保留原内容。返回增强块数。
    """
    if client is None or not model:
        return 0
    enriched = 0
    total = sum(len(s.get("visual_blocks") or []) for s in slides)
    done = 0
    for s in slides:
        page = s.get("slide_info", {}).get("slide_index") or s.get("page", 0)
        raw = (s.get("textual_content") or {}).get("raw_text", "") or \
            page_texts.get(page, "")
        fm = page_formulas.get(page, "")
        for blk in s.get("visual_blocks") or []:
            done += 1
            if on_progress is not None:
                try:
                    on_progress(done, total, {"kind": "semantics"})
                except Exception:
                    pass
            try:
                st = blk.get("internal_structure") or {}
                node_texts = [n.get("text", "") for n in st.get("nodes", [])][:8]
                node_texts = [t for t in node_texts if t.strip()]
                bb = blk.get("bbox") or {}
                old_sd = blk.get("semantic_description") or {}
                prompt = (
                    f"【块类型】{blk.get('block_type', '')}\n"
                    f"【节点文本】{' / '.join(node_texts[:8]) or '（无文本标签）'}\n"
                    f"【占据页面比例】块 {bb.get('w', 0):.0f}x{bb.get('h', 0):.0f}"
                    f" px（约 {bb.get('w', 0)*bb.get('h', 0)/691200*100:.0f}% 页）\n"
                    f"【该页文字】\n{raw[:1200]}\n"
                    + (f"【该页公式】\n{fm[:500]}\n" if fm else "")
                    + (f"【已有描述】\n{json.dumps(old_sd, ensure_ascii=False)[:400]}\n"
                       if old_sd else "")
                    + "请输出该块的语义描述 JSON。"
                )
                resp = client.chat.completions.create(
                    model=model,
                    messages=[{"role": "system", "content": SEMANTIC_SYSTEM},
                              {"role": "user", "content": prompt}],
                    stream=False,
                    reasoning_effort="high",
                    extra_body={"thinking": {"type": "enabled"}},
                )
                data = _safe_json(resp.choices[0].message.content or "")
                if not data:
                    continue
                goal = str(data.get("expression_goal", "") or "").strip()
                role = str(data.get("expression_role", "") or "").strip()
                feats = data.get("expression_features") or []
                cap = str(data.get("vlm_caption", "") or "").strip()
                use = str(data.get("teaching_use", "") or "").strip()
                if not (goal or role or cap):
                    continue
                blk["semantic_description"] = {
                    "block_type": blk.get("block_type", ""),
                    "expression_goal": goal or old_sd.get("expression_goal", ""),
                    "expression_role": role or old_sd.get("expression_role", ""),
                    "expression_features": [
                        str(f)[:30] for f in feats if str(f).strip()][:5]
                        or old_sd.get("expression_features", ["可视化"]),
                    "vlm_caption": cap or old_sd.get("vlm_caption", ""),
                    "teaching_use": use or old_sd.get("teaching_use",
                                                      "教学辅助图示"),
                }
                enriched += 1
            except Exception as e:
                print(f"[语义] 块 {blk.get('block_id')} 增强失败：{e}",
                      file=sys.stderr)
            if done % 10 == 0:
                time.sleep(0.2)
    return enriched


def build_cross_modal_relations(blocks: list[VisualBlock],
                                page_text: str,
                                client=None, model: str = "",
                                page_no: int = 0) -> list[dict]:
    """生成文本锚点 ↔ 块 的跨模态关系。VLM 可用时用 VLM 判定，
    否则用规则：找与块中心水平/垂直最接近的文本行。"""
    relations = []
    if client is not None and model:
        for i, blk in enumerate(blocks, start=1):
            try:
                resp = client.chat.completions.create(
                    model=model,
                    messages=[{"role": "system", "content": RELATION_SYSTEM},
                              {"role": "user", "content":
                               f"【页面正文】\n{page_text[:1500]}\n\n"
                               f"【可视逻辑块】\n"
                               f"类型：{blk.block_type}；"
                               f"目标：{(blk.semantic_description or {}).get('expression_goal', '')}"
                               f"；描述：{(blk.semantic_description or {}).get('vlm_caption', '')[:200]}"}],
                    stream=False,
                    reasoning_effort="high",
                    extra_body={"thinking": {"type": "enabled"}},
                )
                data = _safe_json(resp.choices[0].message.content or "")
                anchor = str(data.get("text_anchor", "") or "")[:60]
                relations.append({
                    "relation_id": f"rel_{page_no:02d}_{i:02d}",
                    "text_anchor": anchor,
                    "target_block_id": blk.block_id,
                    "relation_type": str(data.get("relation_type", "elaboration")),
                    "semantic_link": str(data.get("semantic_link", "") or "")[:120],
                })
            except Exception:
                pass
        return relations
    # 规则回退：无锚点则取页面第一行文本
    first_line = ""
    for ln in page_text.splitlines():
        ln = ln.strip()
        if ln:
            first_line = ln[:60]
            break
    for i, blk in enumerate(blocks, start=1):
        relations.append({
            "relation_id": f"rel_{page_no:02d}_{i:02d}",
            "text_anchor": first_line,
            "target_block_id": blk.block_id,
            "relation_type": "elaboration",
            "semantic_link": f"本块与该页文字共同服务于页面主题（规则回退）",
        })
    return relations


# --------------------------------------------------------------------------
# 主入口
# --------------------------------------------------------------------------
def extract_blocks(pptx_path: str, out_dir: str,
                   atomic_objects: list | None = None,
                   page_texts: dict | None = None,
                   client=None, model: str = "",
                   config: dict | None = None,
                   on_progress=None) -> list[dict]:
    """输入原子对象 + 每页文本，输出每页可视逻辑块列表。

    返回 [{page, blocks: [VisualBlock.to_dict()], relations: [...]}, ...]。
    纯本地可用（client=None 时规则回退）。
    """
    if atomic_objects is None:
        ao_path = Path(out_dir) / "atomic_objects.json"
        if ao_path.is_file():
            atomic_objects = json.loads(ao_path.read_text(encoding="utf-8"))
        else:
            return []
    atomic = [AtomicObject.from_dict(d) for d in atomic_objects]
    page_texts = page_texts or {}
    cfg = {**CLUSTER_CONFIG, **(config or {})}
    # 硬编码修复：用真实页面尺寸（sldSz）覆盖 960×720 参考值——
    # 16:9 页面（1280×720）若按 960×720 算，越界判定会把中心 x>960 的
    # 页面右侧有效对象误剔，且横幅/大图/碎片等比例类阈值偏差 33%
    try:
        from pptx_wzq import extract_pptx_images as _E
        _cx, _cy = _E.read_sld_size(str(pptx_path))
        if _cx and _cy:
            cfg["page_w"] = _cx / 914400 * 96
            cfg["page_h"] = _cy / 914400 * 96
    except Exception:  # pragma: no cover
        pass
    out_dir = Path(out_dir)

    by_page_objs = {}
    for o in atomic:
        by_page_objs.setdefault(o.page, []).append(o)

    slides_out = []
    for page_no in sorted(by_page_objs):
        objs = _filter_noise(by_page_objs[page_no], cfg)
        # 剔除页面文本区（标题/正文/副标题等占位符 shape）：它们不属于
        # 可视逻辑块（内容已由文本提取步骤保留），否则大文本框会与
        # 页内图片/形状互相重叠导致整页合并成 1 块
        objs = [o for o in objs if o.kind != "text_region"]
        # 剔除无几何对象（bbox 宽或高 ≤0，如内联 OMML 公式无独立区域）：
        # 它们无法参与空间聚类、无法渲染成块图；内容由公式提取步骤保留
        objs = [o for o in objs
                if (o.bbox or {}).get("w", 0) > 0 and
                (o.bbox or {}).get("h", 0) > 0]
        # 剔除越界/装饰对象：中心落在页面外（如 x=-18 的页外 shape）
        pw, ph = cfg.get("page_w", 960), cfg.get("page_h", 720)
        objs = [o for o in objs
                if 0 <= o.bbox["x"] + o.bbox["w"] / 2 <= pw and
                0 <= o.bbox["y"] + o.bbox["h"] / 2 <= ph]
        # 文本密度判据（用户准则）：可视逻辑块内文本框应为短标签。
        # 单个 shape 文本 > max_shape_text 字 → 视为潜在正文文本框，
        # 不参与聚类（逻辑图节点/箭头标注通常字少）
        max_shape_text = cfg.get("max_shape_text", 10)
        objs = [o for o in objs
                if not (o.kind == "shape" and
                        len((o.text or "").strip()) > max_shape_text)]
        if not objs:
            slides_out.append({"page": page_no, "blocks": [], "relations": []})
            continue
        # 用户准则①：raster/vector/visio 面积 > 整页 30% → 独立成块
        big_objs, objs = _split_big_objects(objs, cfg)
        # 用户准则②：四向种子扩展区域生长（从种子对象出发四向扩展，
        # 遇到字多文本对象（墙）或边界即停；加法生长天然不切碎表格/
        # 框图，凹区域被墙/间隙阻断成多个区域，无需凸分割）
        sub_clusters = _region_grow(objs, cfg)
        # 大面积对象各自成簇
        for o in big_objs:
            sub_clusters.append([o])
        # 限制每页块数：按面积降序取前 max_blocks_per_slide
        sub_clusters.sort(key=lambda c: -sum(
            (o.bbox.get("w", 0) * o.bbox.get("h", 0)) for o in c))
        if len(sub_clusters) > cfg["max_blocks_per_slide"]:
            sub_clusters = sub_clusters[:cfg["max_blocks_per_slide"]]

        blocks = []
        for ci, cl in enumerate(sub_clusters, start=1):
            blk_id = f"blk_{ci:02d}"
            # 块级文本密度校验：块内文本总量 > max_block_text 字 →
            # 判定混入了正文文本区，剔除长文本成员后重组块（剩余成员
            # 若仍为 1 个有效可视化对象则保留为单对象块）
            max_block_text = cfg.get("max_block_text", 30)
            total_text = sum(len((o.text or "").strip())
                             for o in cl if o.kind in ("shape", "connector"))
            if total_text > max_block_text:
                cl = [o for o in cl
                      if not (o.kind in ("shape", "connector") and
                              len((o.text or "").strip()) > max_shape_text)]
            if not cl:
                continue
            min_x = min(o.bbox["x"] for o in cl)
            min_y = min(o.bbox["y"] for o in cl)
            max_x = max(o.bbox["x"] + o.bbox["w"] for o in cl)
            max_y = max(o.bbox["y"] + o.bbox["h"] for o in cl)
            bbox = {"x": round(min_x, 1), "y": round(min_y, 1),
                    "w": round(max_x - min_x, 1), "h": round(max_y - min_y, 1)}
            # 符号碎片过滤（用户准则）：可视逻辑块应是有表达目的的可视化
            # 对象集合。无文本、无媒体（图/表/visio）、无连接（connector）
            # 的纯形状块（如电路元件符号：2 个小形状拼成）不构成可视逻辑块：
            #   对象数 ≤3 → 无条件碎片（单个/成对符号）；
            #   对象数 >3 但块面积 < 页面 min_block_area_ratio → 碎片。
            # 含文本标签 / 图 / 表 / 连接的块不受此限制（逻辑图节点短标签、
            # 独立小图均保留）。
            has_text = any((o.text or "").strip() for o in cl)
            has_media = any(o.kind in ("raster", "vector", "visio",
                                       "table", "chart") for o in cl)
            has_conn = any(o.kind == "connector" for o in cl)
            if not has_text and not has_media and not has_conn:
                _page_area = cfg.get("page_w", 960) * cfg.get("page_h", 720)
                if len(cl) <= 3 or \
                        bbox["w"] * bbox["h"] < _page_area * \
                        cfg.get("min_block_area_ratio", 0.10):
                    continue  # 符号碎片：不输出
            zs = [o.z_index for o in cl]
            block = VisualBlock(
                block_id=blk_id, page=page_no, bbox=bbox,
                z_index_range=[min(zs), max(zs)],
                member_obj_ids=[o.obj_id for o in cl],
                is_single=len(cl) == 1,
            )
            block.block_type = _guess_block_type(cl)
            block.internal_structure = _infer_topology(cl)
            block.semantic_description = _fallback_description(block)
            # 文字空间密度：块内文本字符数 / 块面积（px²），放大 1e6 便于阅读
            _area = max(1.0, bbox["w"] * bbox["h"])
            _chars = sum(len((o.text or "").strip())
                         for o in cl if o.kind in ("shape", "connector"))
            block.text_density = round(_chars / _area * 1_000_000, 3)
            blocks.append(block)

        # VLM 增强（可选）
        if client is not None:
            page_text = page_texts.get(page_no, "")
            for blk in blocks:
                # 渲染块 PNG（仅复合块或需要时）
                img_path = out_dir / "by_page" / f"{blk.block_id}.png"
                desc = describe_block(client, model, blk, page_text,
                                      img_path if img_path.is_file() else None)
                if desc.get("block_type"):
                    blk.block_type = desc["block_type"]
                blk.semantic_description = desc

        relations = build_cross_modal_relations(
            blocks, page_texts.get(page_no, ""), client, model, page_no)
        slides_out.append({
            "page": page_no,
            "blocks": [b.to_dict() for b in blocks],
            "relations": relations,
        })
        if on_progress is not None:
            try:
                on_progress(page_no, len(by_page_objs),
                            {"kind": "blocks", "n_blocks": len(blocks)})
            except Exception:
                pass
    return slides_out


def load_atomic_objects(out_dir: Path) -> list[dict]:
    """从结果目录读 atomic_objects.json（递归搜索 _proc 也可）。"""
    for p in (out_dir / "atomic_objects.json",
              out_dir / "过程文件" / "img" / "atomic_objects.json",
              out_dir / "_proc" / "img" / "atomic_objects.json"):
        if p.is_file():
            try:
                return json.loads(p.read_text(encoding="utf-8"))
            except Exception:
                return []
    return []
