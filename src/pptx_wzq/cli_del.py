# -*- coding: utf-8 -*-
"""pptx-del —— 图块删除后处理命令（v3.2.0）

pptx-paser 全流程完成后的用户后处理：删除指定图块（images/ 渲染图
或 sources/ 源文件）及其**全部关联输出**，等价于该组合在源 PPT 中不存在。
不影响其他图块/文本/公式资源；默认备份 + 一致性校验，不修改任何
pptx-paser 现有功能。

用法示例：
    pptx-del .\\images\\slide_29_blk_01.png -all            # 单块全量删除
    pptx-del .\\images\\slide_29_blk_01.png -all --dry-run   # 预演
    pptx-del a.png b.png -all                                # 批量

RC：0=成功；1=校验失败/删除异常（可从备份恢复）；2=定位/用法错误。
"""
import argparse
import json
import re
import shutil
import sys
import time
from collections import Counter
from pathlib import Path

# ---------------------------------------------------------------------------
# 基础定位
# ---------------------------------------------------------------------------
_SLIDE_PAT = re.compile(r"slide_(\d+)_(blk_\d+)")
_GRP_XML_PAT = re.compile(r"slide_(\d+)_(blk_\d+)_grp\.xml$")


def parse_target(path: Path) -> tuple[int, str] | None:
    """从文件名解析 (page, block_id)：slide_29_blk_01.png / _grp.xml 均可。"""
    m = _SLIDE_PAT.search(path.name)
    if m:
        return int(m.group(1)), m.group(2)
    return None


def locate_result_dir(out_dir: Path) -> Path:
    """校验结果目录：必须存在且含唯一 *_visual_blocks.json。"""
    if not out_dir.is_dir():
        raise SystemExit(f"[错误] 结果目录不存在：{out_dir}")
    vbs = list(out_dir.glob("*_visual_blocks.json"))
    if len(vbs) != 1:
        raise SystemExit(
            f"[错误] 结果目录需含唯一 *_visual_blocks.json（当前 {len(vbs)} 个）："
            f"{out_dir}\n提示：请用 -o 指定正确的输出目录")
    return out_dir


# ---------------------------------------------------------------------------
# 删除计划（内存模型）
# ---------------------------------------------------------------------------
class Plan:
    def __init__(self, out_dir: Path, stem: str):
        self.out = out_dir
        self.stem = stem
        self.vb_path = out_dir / f"{stem}_visual_blocks.json"
        self.bd_path = out_dir / f"{stem}_visualBlock_text_binding.json"
        self.cap_path = out_dir / f"{stem}_captions.md"
        self.images_dir = out_dir / "images"
        self.sources_dir = out_dir / "sources"
        self.targets: list[tuple[int, str]] = []   # (page, block_id)
        self.blocks: list[dict] = []               # 找到的块（含所在 slide 引用）
        self.files: list[Path] = []                # 待移出的物理文件
        self.n_rel = 0
        self.n_bind = 0
        self.n_cap = 0
        self.warnings: list[str] = []


def build_plan(out_dir: Path, raw_targets: list[Path]) -> Plan:
    """解析目标 → 定位块 → 汇总待删清单。目标不在 JSON 中则报错（防悬空）。"""
    out_dir = locate_result_dir(out_dir)
    stem = next(out_dir.glob("*_visual_blocks.json")).name[:-len("_visual_blocks.json")]
    plan = Plan(out_dir, stem)

    vb = json.loads(plan.vb_path.read_text(encoding="utf-8"))
    by_key = {}
    for s in vb.get("slides", []):
        pg = (s.get("slide_info") or {}).get("slide_index", s.get("page", 0))
        for b in s.get("visual_blocks", []):
            by_key[(pg, b.get("block_id"))] = (s, b)

    # 去重（同目标多次给出）
    seen = set()
    for rt in raw_targets:
        key = parse_target(rt)
        if key is None:
            plan.warnings.append(f"[警告] 无法解析目标文件名（跳过）：{rt.name}")
            continue
        if key in seen:
            continue
        seen.add(key)
        hit = by_key.get(key)
        if hit is None:
            raise SystemExit(
                f"[错误] 目标块不存在于 visual_blocks.json：{rt.name}（页 {key[0]} {key[1]}）。"
                f"\n可能已删除过（幂等）或文件名与结果不匹配。RC=2 不删除任何文件。")
        s, b = hit
        plan.targets.append(key)
        plan.blocks.append(b)
        # 物理文件：rendered_image / xml_source / raster_png / vector_resources
        a = b.get("assets") or {}
        for k in ("rendered_image", "xml_source", "raster_png"):
            v = a.get(k)
            if v:
                p = (out_dir / v).resolve()
                if p.is_file():
                    plan.files.append(p)
        for v in (a.get("vector_resources") or []):
            p = (out_dir / v).resolve()
            if p.is_file():
                plan.files.append(p)
        # rldimg：独占校验（其他块 XML 引用则保留）
        for v in (a.get("internal_resources") or []):
            p = (out_dir / v).resolve()
            if p.is_file() and _is_rldimg_exclusive(out_dir, p.name):
                plan.files.append(p)
            elif p.is_file():
                plan.warnings.append(f"[警告] rldimg 被其他块引用（保留）：{v}")

    # 统计文档内命中
    for s in vb.get("slides", []):
        pg = (s.get("slide_info") or {}).get("slide_index", s.get("page", 0))
        plan.n_rel += sum(
            1 for r in s.get("cross_modal_relations", [])
            if (pg, r.get("target_block_id")) in set(plan.targets))
    bd = _read_json(plan.bd_path)
    if bd:
        tset = set(plan.targets)
        plan.n_bind = sum(1 for b in bd.get("bindings", [])
                          if (b.get("page"), b.get("target_block_id")) in tset)
    if plan.cap_path.is_file():
        cap = plan.cap_path.read_text(encoding="utf-8")
        tset = set(plan.targets)
        plan.n_cap = sum(
            1 for m in re.finditer(r"### IMG\d+ — `(slide_(\d+)_(blk_\d+)_grp\.xml)`", cap)
            if (int(m.group(2)), m.group(3)) in tset)
    return plan


def _is_rldimg_exclusive(out_dir: Path, fname: str) -> bool:
    """该 rldimg 文件是否只被目标块引用（扫描所有 XML 段内文件名出现次数）。"""
    n_ref = 0
    for x in out_dir.glob("sources/*.xml"):
        try:
            txt = x.read_text(encoding="utf-8", errors="ignore")
        except OSError:
            continue
        if fname in txt:
            n_ref += 1
    return n_ref <= 1   # 仅自身 1 处引用（或 0）视为独占


def _read_json(p: Path) -> dict | None:
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return None


# ---------------------------------------------------------------------------
# 备份
# ---------------------------------------------------------------------------
def backup(out_dir: Path, plan: Plan) -> Path | None:
    """复制受影响文件到 _del_backup_<ts>/（JSON 为删除前原始版，可一键回滚）。"""
    ts = time.strftime("%Y%m%d_%H%M%S")
    bak = out_dir / f"_del_backup_{ts}"
    bak.mkdir(exist_ok=True)
    n = 0
    for p in (plan.vb_path, plan.bd_path, plan.cap_path):
        if p.is_file():
            shutil.copy2(p, bak / p.name)
            n += 1
    for f in set(plan.files):
        if f.is_file():
            shutil.copy2(f, bak / f.name)
            n += 1
    return bak if n else None


# ---------------------------------------------------------------------------
# 文档修改
# ---------------------------------------------------------------------------
def apply_vb_json(plan: Plan) -> None:
    """visual_blocks.json：删块 + 同页目标关系 + summary 重建 + relation_id 重排。"""
    vb = json.loads(plan.vb_path.read_text(encoding="utf-8"))
    tset = set(plan.targets)
    for s in vb.get("slides", []):
        pg = (s.get("slide_info") or {}).get("slide_index", s.get("page", 0))
        s["visual_blocks"] = [b for b in (s.get("visual_blocks") or [])
                              if (pg, b.get("block_id")) not in tset]
        s["cross_modal_relations"] = [
            r for r in (s.get("cross_modal_relations") or [])
            if (pg, r.get("target_block_id")) not in tset]
        for i, r in enumerate(s["cross_modal_relations"], start=1):
            r["relation_id"] = f"rel_{pg:02d}_{i:02d}"
    bt = Counter(b.get("block_type", "")
                 for s in vb.get("slides", [])
                 for b in (s.get("visual_blocks") or []))
    vb["summary"]["blocks_total"] = sum(
        len(s.get("visual_blocks") or []) for s in vb.get("slides", []))
    vb["summary"]["slides_with_blocks"] = sum(
        1 for s in vb.get("slides", []) if s.get("visual_blocks"))
    vb["summary"]["block_types"] = dict(bt)
    plan.vb_path.write_text(json.dumps(vb, ensure_ascii=False, indent=1),
                            encoding="utf-8")


def apply_binding(plan: Plan) -> None:
    """binding：删 (page, block_id) 命中条目 + relation_id 重排 + summary 更新。"""
    bd = json.loads(plan.bd_path.read_text(encoding="utf-8"))
    tset = set(plan.targets)
    bd["bindings"] = [b for b in bd.get("bindings", [])
                      if (b.get("page"), b.get("target_block_id")) not in tset]
    for i, b in enumerate(bd["bindings"], start=1):
        b["relation_id"] = f"rel_{b.get('page', 0):02d}_{i:02d}"
    bd["summary"]["blocks_total"] = len(bd["bindings"])
    bd["summary"]["bindings_total"] = len(bd["bindings"])
    plan.bd_path.write_text(json.dumps(bd, ensure_ascii=False, indent=1),
                            encoding="utf-8")


def apply_captions(plan: Plan) -> None:
    """captions.md：删目标 IMG 条目（含条目后续内容行，到下一个 ### / ## / 文件尾）。"""
    if not plan.cap_path.is_file():
        return
    lines = plan.cap_path.read_text(encoding="utf-8").split("\n")
    tset = set(plan.targets)
    out, skip = [], False
    for ln in lines:
        m = re.match(r"^### IMG\d+ — `(slide_(\d+)_(blk_\d+)_grp\.xml)`", ln)
        if m:
            skip = (int(m.group(2)), m.group(3)) in tset
            if not skip:
                out.append(ln)
            continue
        if not skip:
            out.append(ln)
    plan.cap_path.write_text("\n".join(out), encoding="utf-8")


def move_files(plan: Plan, bak: Path | None) -> None:
    """images/sources 目标文件 mv 到备份目录（等价删除 + 天然备份）。"""
    if bak is None:
        for f in set(plan.files):
            if f.is_file():
                f.unlink(missing_ok=True)
        return
    for f in set(plan.files):
        if f.is_file():
            shutil.move(str(f), bak / f.name)


def fix_binding_summary(plan: Plan) -> None:
    """binding summary 物理修正（必须在文件移动后）：sources_total/xml/rldimg。"""
    bd = json.loads(plan.bd_path.read_text(encoding="utf-8"))
    src = plan.sources_dir
    if not src.is_dir():
        return
    top = [f for f in src.iterdir() if f.is_file()]
    n_xml = len([f for f in top if f.suffix == ".xml"])
    rld = src / "rldimg"
    n_rld = len(list(rld.iterdir())) if rld.is_dir() else 0
    bd["summary"]["sources_total"] = len(top) + n_rld
    bd["summary"]["xml_sources_total"] = n_xml
    bd["summary"]["rldimg_total"] = n_rld
    plan.bd_path.write_text(json.dumps(bd, ensure_ascii=False, indent=1),
                            encoding="utf-8")


# ---------------------------------------------------------------------------
# 一致性校验
# ---------------------------------------------------------------------------
def verify(plan: Plan, targets: set) -> list[str]:
    """删除后一致性校验：JSON 可解析 / 目标残留 0 / captions 与物理 XML 数一致。"""
    errs = []
    vb = _read_json(plan.vb_path)
    bd = _read_json(plan.bd_path)
    if vb is None:
        errs.append("visual_blocks.json 解析失败")
    else:
        for s in vb.get("slides", []):
            pg = (s.get("slide_info") or {}).get("slide_index", s.get("page", 0))
            if any((pg, b.get("block_id")) in targets
                   for b in (s.get("visual_blocks") or [])):
                errs.append(f"visual_blocks 残留目标块：p{pg}")
            if any((pg, r.get("target_block_id")) in targets
                   for r in (s.get("cross_modal_relations") or [])):
                errs.append(f"visual_blocks 残留目标关系：p{pg}")
    if bd is None:
        errs.append("binding 解析失败")
    else:
        if any((b.get("page"), b.get("target_block_id")) in targets
               for b in bd.get("bindings", [])):
            errs.append("binding 残留目标条目")
    if plan.cap_path.is_file():
        cap = plan.cap_path.read_text(encoding="utf-8")
        if any((int(m.group(2)), m.group(3)) in targets
               for m in re.finditer(
                   r"### IMG\d+ — `(slide_(\d+)_(blk_\d+)_grp\.xml)`", cap)):
            errs.append("captions 残留目标条目")
        n_cap = len(set(re.findall(r"slide_\d+_blk_\d+_grp\.xml", cap)))
        n_xml = len([f for f in (plan.sources_dir.glob("*.xml")
                                 if plan.sources_dir.is_dir() else [])])
        if n_cap != n_xml:
            errs.append(f"captions 条目({n_cap}) ≠ 物理 XML({n_xml})")
    for f in set(plan.files):
        if f.is_file():
            errs.append(f"物理文件残留：{f.name}")
    return errs


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def _build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(
        prog="pptx-del",
        description="图块删除后处理：删除指定块及其全部关联输出（images/sources/JSON/binding/captions），"
                    "等价于该组合不存在；默认备份 + 一致性校验，不改动 pptx-paser 现有功能。",
        epilog="示例：\n  pptx-del .\\images\\slide_29_blk_01.png -all\n"
               "  pptx-del .\\images\\a.png .\\images\\b.png -all --dry-run")
    ap.add_argument("targets", nargs="+", type=Path,
                    help="images/ 下块渲染图（slide_XX_blk_NN.png）或 sources/ 下源文件路径，≥1 个")
    ap.add_argument("-o", "--out-dir", type=Path, default=None,
                    help="结果目录（默认由 target 路径自动推断：images 的父目录）")
    ap.add_argument("-a", "-all", "--all", dest="all", action="store_true",
                    help="全量删除（images+sources+JSON+binding+captions+summary 重建）")
    ap.add_argument("--dry-run", action="store_true",
                    help="预演：只输出将删清单与影响范围，不落盘")
    ap.add_argument("--no-backup", action="store_true",
                    help="跳过备份（默认备份到 _del_backup_<ts>/）")
    ap.add_argument("--verify", action="store_true",
                    help="删除后强制详细校验（默认也做基础校验）")
    ap.add_argument("--textbook", action="store_true",
                    help="textbook.md 出现块级引用时强制删除（默认仅警告保留）")
    ap.add_argument("--json", action="store_true", dest="as_json",
                    help="输出机器可读 JSON 报告")
    return ap


def _infer_out_dir(targets: list[Path]) -> Path | None:
    """images/slide_X.png → out_dir（父目录的父目录）；多 target 时取共同父目录。"""
    parents = []
    for t in targets:
        p = t.resolve()
        if p.parent.name == "images" or p.parent.name == "sources":
            parents.append(p.parent.parent)
        else:
            parents.append(p.parent)
    if len(set(map(str, parents))) == 1:
        return parents[0]
    return None


def main(argv=None) -> int:
    ap = _build_parser()
    args = ap.parse_args(argv)
    out_dir = args.out_dir or _infer_out_dir(args.targets)
    if out_dir is None:
        print("[错误] 无法推断结果目录（多个 target 不在同一输出下），请用 -o 指定",
              file=sys.stderr)
        return 2

    try:
        plan = build_plan(out_dir, args.targets)
    except SystemExit as e:
        print(e, file=sys.stderr)
        return 2

    # ---- 清单输出 ----
    lines = [f"[定位] 目标 {len(plan.targets)} 个块：" +
             "、".join(f"p{p} {b}" for p, b in plan.targets)]
    lines.append(f"[清单] 将删除：文件 {len(set(plan.files))} 个、"
                 f"JSON 块 {len(plan.targets)}、关系 {plan.n_rel}、"
                 f"binding {plan.n_bind} 条、captions {plan.n_cap} 条")
    for w in plan.warnings:
        lines.append(w)
    if args.as_json:
        import json as _j
        print(_j.dumps({
            "targets": [{"page": p, "block_id": b} for p, b in plan.targets],
            "files": [str(f) for f in sorted(set(plan.files))],
            "relations": plan.n_rel, "bindings": plan.n_bind,
            "captions": plan.n_cap, "warnings": plan.warnings,
            "dry_run": args.dry_run}, ensure_ascii=False, indent=1))
    else:
        print("\n".join(lines))

    if args.dry_run:
        print("[预演] 未执行任何删除")
        return 0

    # ---- 执行 ----
    bak = None if args.no_backup else backup(out_dir, plan)
    if bak:
        print(f"[备份] {len(list(bak.iterdir())) if bak.is_dir() else 0} 个文件 → {bak}")
    apply_vb_json(plan)
    apply_binding(plan)
    apply_captions(plan)
    move_files(plan, bak)
    fix_binding_summary(plan)

    # ---- 校验 ----
    tset = set(plan.targets)
    errs = verify(plan, tset)
    if errs:
        print("[校验失败]", file=sys.stderr)
        for e in errs:
            print(f"  ✗ {e}", file=sys.stderr)
        print(f"  备份位置：{bak}（复制回原位即恢复）", file=sys.stderr)
        return 1
    if not args.as_json:
        print("[校验] 通过：目标残留 0、JSON 可解析、captions 与物理一致")
    if args.verify and not args.as_json:
        vb = _read_json(plan.vb_path)
        print(f"[汇总] 块 {vb['summary']['blocks_total']}（删除 {len(plan.targets)}）、"
              f"绑定 {plan.n_bind} 条已删；其他资源未改动")
    return 0


if __name__ == "__main__":
    sys.exit(main())
