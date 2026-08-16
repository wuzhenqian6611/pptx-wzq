"""cli_bind.py — pptx-bind：把教材文案与对应页图片关系绑定为 JSON。

输入：textbook.md（每页文案，## 第 N 页）、text_entries.json（文本 ID+坐标）、
     manifest.json（图片坐标/类型/来源）、images 目录（文件名 slide_NN_ 编码页码）、
     captions.md（图片 AI 解读，### IMGxxxx — slide_NN_）。
输出：<名>_binding.json，按页组织（v3 结构）：

    {
      "stem": "xxx",
      "pages": [
        {"page": 4,
         "text": "第 4 页文案全文…",
         "images": [
           {"file": "slide_04_pic_05.png",
            "caption": "…qwen 诠释…",
            "source": "ppt/media/image1.png",
            "kind": "raster|vector|visio",
            "image_id": "IMG0001",
            "page": 4,
            "paragraph": 2, "text_id": "TXT004-02",
            "w": 448, "h": 300, "x": 96, "y": 220,
            "position": "位于第 4 页（第 2 段），该图对该页文字表达起『…』的作用",
            "relation": "…≤60字 图文逻辑关系…"}]}
      ],
      "summary": {"pages": 10, "images_total": 12, "pages_with_image": 8,
                  "relations": 12, "positions": 12}
    }

v3.1 变更（与 word-wzq 格式统一，方案见跨项目分析报告）：
- position：确定性位置前缀 + DeepSeek position_role（≤40 字，功能角色维度）
  拼接；relation 为 ≤50 字逻辑关联（逻辑维度）——两项在一次调用中判断；
- kind 统一为 original_format 直出（vsdx/svg/wmf/emf/png…）；
- w/h 统一为图片物理像素（manifest width/height，0 时回退显示尺寸 shape_w/h）；
- summary 含 relations/positions 计数；解析兼容「第 N 页/节」（可互读 word 产物）。

用法：
    pptx-bind <产物目录> [-o binding.json]
              [--textbook a.md] [--text-entries e.json] [--manifest m.json]
              [--images-dir dir] [--captions c.md]
              [--model deepseek-v4-flash] [--base-url …] [--api-key-env …]
              [--no-relation] [--json] [--version]

退出码：0 成功 / 1 处理异常 / 2 参数或环境错误。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from pathlib import Path

from pptx_wzq.cli_common import (EXIT_ERR, EXIT_OK, EXIT_USAGE, print_json,
                        quiet_stdout)

VERSION = "pptx-bind 3.1.0 (图文ID+坐标+位置作用+逻辑关系·统一格式)"
DEFAULT_MODEL = "deepseek-v4-flash"
DEFAULT_BASE_URL = "https://api.deepseek.com"
DEFAULT_KEY_ENV = "DEEPSEEK_API_KEY"

RASTER_EXTS = {"png", "jpg", "jpeg", "bmp", "gif", "webp", "tif", "tiff"}

LINK_SYSTEM = (
    "你是高校教材建设专家、PPT 文档图文分析专家。下面给出 PPT 文档某一页"
    "的文本、其中一张图片的 AI 诠释。请输出 JSON，包含两个字段：\n"
    '{"position_role": "该图对该页文字表达所起的作用，不超过40字——描述图在'
    '本页表达中扮演的功能角色（结构/功能维度），如：直观呈现本页所述概念 / '
    '以实例补充本页未展开的细节 / 以数据例证本页观点 / 总结本页流程与要点 / '
    '引入本页话题 / 对比本页所述方案",\n'
    ' "relation": "该图与本页文字的逻辑关联陈述，不超过50字——描述图与文字的'
    '逻辑关系（逻辑维度），如：该图以示意图说明…，与本页呈说明关系 / '
    '该图以数据证明…，与本页呈证明关系"}\n'
    "注意区分两个角度：position_role 讲「图在本页表达中起什么作用」"
    "（功能/角色），relation 讲「图与文字是什么逻辑关系」（关联类型），两者不同。\n"
    "只输出 JSON，不要输出任何其他内容。"
)


def _compose_position(page: int, paragraph: int, role: str = "") -> str:
    """拼接 position：确定性位置前缀（代码生成）+ LLM 生成的功能角色。"""
    pos = f"位于第 {page} 页（第 {paragraph} 段）" if paragraph else \
        f"位于第 {page} 页"
    role = (role or "").strip().rstrip("。；，")
    if not role:
        role = "辅助理解本页知识点"
    return f"{pos}，该图对该页文字表达起『{role}』的作用"


def _split_pages(content: str) -> dict:
    """按 '## 第 N 页/节' 分块（页/节双兼容，可互读 word 产物）。"""
    pages = {}
    cur = None
    for line in content.splitlines():
        m = re.match(r"^##\s*第\s*(\d+)\s*(?:页|节)\s*$", line.strip())
        if m:
            cur = int(m.group(1))
            pages.setdefault(cur, [])
        elif cur is not None:
            pages[cur].append(line)
    return pages


def parse_textbook(path: Path) -> dict:
    """textbook.md → {page: 文案全文}。"""
    out = {}
    for page, lines in _split_pages(path.read_text(encoding="utf-8")).items():
        text = "\n".join(x for x in lines if x.strip()).strip()
        if text:
            out[page] = text
    return out


def parse_text_entries(path: Path) -> dict:
    """text_entries.json → {page: [{text_id, text, x, y, w, h}]}。"""
    out = {}
    try:
        entries = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return out
    for e in entries:
        if not e.get("id"):
            continue
        pg = e.get("page")
        if pg is None:
            continue
        out.setdefault(pg, []).append({
            "text_id": e["id"],
            "text": e.get("text", ""),
            "x": e.get("x", 0), "y": e.get("y", 0),
            "w": e.get("w", 0), "h": e.get("h", 0),
        })
    return out


def parse_manifest(path: Path) -> dict:
    """manifest.json → {page: {file: rec}}，rec 含 x/y/shape_w/shape_h/kind。"""
    out = {}
    try:
        recs = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return out
    for r in recs:
        if not r.get("output_file"):
            continue
        pg = r.get("page")
        if pg is None:
            continue
        out.setdefault(pg, {})[r["output_file"]] = r
    return out


def _page_of(file_name: str):
    """从文件名 slide_NN_... 提取页码；失败返回 None。"""
    m = re.match(r"slide_(\d+)_", file_name)
    return int(m.group(1)) if m else None


def index_images(images_dir: Path) -> dict:
    """images 目录 → {page: [文件名]}（按 slide_NN 分页）。"""
    out = {}
    if not images_dir.is_dir():
        return out
    for p in sorted(images_dir.iterdir()):
        pg = _page_of(p.name)
        if pg is not None:
            out.setdefault(pg, []).append(p.name)
    return out


def parse_captions(path: Path) -> dict:
    """captions.md → {file_name: {"id": img_id, "text": 解读}}。"""
    out = {}
    cur_file, cur_id, buf = None, None, []
    for ln in path.read_text(encoding="utf-8").splitlines():
        m = re.match(r"^###\s+(IMG\d+)\s*—\s*`?([^`\s]+\.\w+)\s*`?", ln)
        if m:
            if cur_file is not None and buf:
                out[cur_file] = {"id": cur_id,
                                 "text": " ".join(x.strip() for x in buf
                                                  if x.strip())}
            cur_id, cur_file = m.group(1), m.group(2)
            buf = []
        elif cur_file is not None:
            buf.append(ln)
    if cur_file is not None and buf:
        out[cur_file] = {"id": cur_id,
                         "text": " ".join(x.strip() for x in buf
                                          if x.strip())}
    return out


def _judge_pos_rel(client, model: str, caption: str, page_text: str,
                   retries: int = 1) -> dict:
    """DeepSeek 判断「位置作用 + 逻辑关系」→ {position_role, relation}。

    - position_role：图对该页文字表达所起的作用（≤40 字，功能/结构维度）；
    - relation      ：图与文字的逻辑关联（≤50 字，逻辑维度）。
    失败返回 {}（调用方用确定性位置兜底，不中断绑定）。
    """
    out = {"position_role": "", "relation": ""}
    if not caption:
        return out
    try:
        messages = [
            {"role": "system", "content": LINK_SYSTEM},
            {"role": "user", "content":
             f"【所在页文本】\n{page_text[:1500]}\n\n"
             f"【图片 AI 诠释】\n{caption[:800]}\n\n"
             "请输出 JSON。"},
        ]
        last_err = None
        for attempt in range(retries + 1):
            try:
                resp = client.chat.completions.create(
                    model=model, messages=messages, stream=False,
                    reasoning_effort="high",
                    extra_body={"thinking": {"type": "enabled"}},
                )
                s = (resp.choices[0].message.content or "").strip()
                m = re.search(r"\{.*\}", s, re.S)
                if m:
                    data = json.loads(m.group(0))
                    role = str(data.get("position_role", "") or "").strip()
                    rel = str(data.get("relation", "") or "").strip()
                    out["position_role"] = re.sub(
                        r'^["\'“”\s]+|["\'“”\s]+$', "", role)[:60]
                    out["relation"] = re.sub(
                        r'^["\'“”\s]+|["\'“”\s]+$', "", rel)[:80]
                    return out
                return out
            except Exception as e:
                last_err = e
                if attempt < retries:
                    time.sleep(2 * (attempt + 1))
        print(f"[警告] 图文位置/关系判断失败({last_err})，用确定性位置兜底",
              file=sys.stderr)
    except Exception:
        pass
    return out


def _fallback_relation(caption: str, n: int = 50) -> str:
    """失败回退：取 caption 前 n 字。"""
    return caption[:n].rstrip("；，。 ")


def _fallback_position(page: int, paragraph: int) -> str:
    """position 失败回退：前缀 + 默认作用。"""
    return _compose_position(page, paragraph, "辅助理解本页知识点")


def _kind_of(rec: dict) -> str:
    """kind 统一为 original_format 直出（vsdx/svg/wmf/emf/png…）。"""
    return (rec.get("original_format") or "").lower() or "unknown"


def _pick_paragraph(entries: list) -> tuple:
    """从该页文本条目选代表段：优先「内容」型首条，否则首条。
    返回 (paragraph 序号, text_id)。"""
    if not entries:
        return 0, ""
    # 优先内容型（type=="内容"）且非过短
    for i, e in enumerate(entries, start=1):
        if e.get("type") == "内容" and len(e.get("text", "")) >= 4:
            return i, e.get("text_id", "")
    return 1, entries[0].get("text_id", "")


def build_binding(stem: str, textbook: dict, entries_by_page: dict,
                  manifest_by_page: dict, img_by_page: dict,
                  cap_by_file: dict,
                  relations: dict | None = None,
                  positions: dict | None = None) -> dict:
    """按页绑定文案与图片（v3.1：position=确定性位置+LLM 角色，w/h=物理像素）。

    relations: {file: relation 文本}；positions: {file: position_role 文本}；
    None 或缺失时回退 caption 首句 / 模板。
    """
    pages = sorted(set(textbook) | set(img_by_page))
    out_pages = []
    used_ids = set()
    for p in pages:
        entries = entries_by_page.get(p, [])
        paragraph, text_id = _pick_paragraph(entries)
        imgs = []
        for fname in sorted(img_by_page.get(p, [])):
            cap = cap_by_file.get(fname, {})
            img_id = cap.get("id") if cap.get("id") else ""
            if not img_id:
                # 无 caption 条目的图按页序补稳定 ID
                base = f"IMG{len(used_ids) + 1:04d}"
                while base in used_ids:
                    base = f"IMG{len(used_ids) + 2:04d}"
                img_id = base
            used_ids.add(img_id)
            caption = cap.get("text", "")
            rel = (relations or {}).get(fname, "") or _fallback_relation(caption)
            role = (positions or {}).get(fname, "")
            pos = _compose_position(p, paragraph, role) if role else \
                _fallback_position(p, paragraph)
            rec = (manifest_by_page.get(p) or {}).get(fname, {})
            # w/h：优先图片物理像素（width/height），0 时回退幻灯片显示尺寸
            w = rec.get("width") or rec.get("shape_w", 0)
            h = rec.get("height") or rec.get("shape_h", 0)
            imgs.append({
                "file": fname,
                "caption": caption,
                "source": rec.get("source_media", ""),
                "kind": _kind_of(rec),
                "image_id": img_id,
                "page": p,
                "paragraph": paragraph,
                "text_id": text_id,
                "w": w, "h": h,
                "x": rec.get("x", 0), "y": rec.get("y", 0),
                "position": pos,
                "relation": rel,
            })
        out_pages.append({"page": p,
                          "text": textbook.get(p, ""),
                          "images": imgs,
                          "has_image": bool(imgs)})
    n_img = sum(len(p["images"]) for p in out_pages)
    n_bound = sum(1 for p in out_pages if p["has_image"])
    n_link = sum(1 for p in out_pages for im in p["images"]
                 if im.get("relation"))
    n_pos = sum(1 for p in out_pages for im in p["images"]
                if im.get("position"))
    return {
        "stem": stem,
        "pages": out_pages,
        "summary": {"pages": len(out_pages),
                    "images_total": n_img,
                    "pages_with_image": n_bound,
                    "relations": n_link,
                    "positions": n_pos},
    }


def _locate(out_dir: Path, args):
    """定位 textbook / text_entries / manifest / images / captions。
    text_entries/manifest 可能位于子目录（_proc/text、_proc/img 等），
    根目录找不到时递归搜索（rglob）。"""
    tb = args.textbook
    if tb is None:
        cands = sorted(out_dir.glob("*_textbook.md"))
        tb = cands[0] if cands else None
    te = args.text_entries
    if te is None:
        cands = sorted(out_dir.glob("*_text_entries.json")) or \
            sorted(out_dir.rglob("*_text_entries.json"))
        te = cands[0] if cands else None
    mf = args.manifest
    if mf is None:
        cands = sorted(out_dir.glob("manifest.json")) or \
            sorted(out_dir.rglob("manifest.json"))
        mf = cands[0] if cands else None
    img_dir = args.images_dir or (out_dir / "images")
    if not img_dir.is_dir():
        img_dir = out_dir / "过程文件" / "img" / "images"
    cap = args.captions
    if cap is None:
        cands = sorted(out_dir.glob("*_captions.md")) + \
            sorted(out_dir.glob("images_captions.md"))
        cap = cands[0] if cands else None
    return tb, te, mf, img_dir, cap


def _main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        prog="pptx-bind",
        description="把教材文案与对应页图片关系绑定为 JSON"
                    "（含图片ID/文本ID/坐标/图文逻辑关系）")
    ap.add_argument("dir", nargs="?", default=".",
                    help="产物目录（自动找 textbook/text_entries/manifest/"
                         "images/captions）")
    ap.add_argument("-o", "--output", default=None,
                    help="输出 json 路径（默认 <目录>/<名>_binding.json）")
    ap.add_argument("--textbook", type=Path, default=None, help="教材文案 md")
    ap.add_argument("--text-entries", type=Path, default=None,
                    help="文本条目 json（含 text_id/坐标，默认自动查找）")
    ap.add_argument("--manifest", type=Path, default=None,
                    help="manifest.json（图片坐标/类型，默认自动查找）")
    ap.add_argument("--images-dir", type=Path, default=None, help="images 目录")
    ap.add_argument("--captions", type=Path, default=None, help="图片解读 md")
    ap.add_argument("--model", default=DEFAULT_MODEL,
                    help=f"逻辑关系模型（默认 {DEFAULT_MODEL}）")
    ap.add_argument("--base-url", default=DEFAULT_BASE_URL,
                    help=f"OpenAI 兼容端点（默认 {DEFAULT_BASE_URL}）")
    ap.add_argument("--api-key-env", default=DEFAULT_KEY_ENV,
                    help=f"API Key 环境变量名（默认 {DEFAULT_KEY_ENV}）")
    ap.add_argument("--no-relation", action="store_true",
                    help="不调用模型，relation 回退为 caption 首句（0 Token）")
    ap.add_argument("--resume", action="store_true",
                    help="断点续跑：已有 binding.json 的页保留原样，"
                         "只处理未完成页并合并写出")
    ap.add_argument("--json", action="store_true",
                    help="把统计输出到 stdout")
    ap.add_argument("--version", action="version", version=VERSION)
    args = ap.parse_args(argv)

    out_dir = Path(args.dir)
    tb, te, mf, img_dir, cap = _locate(out_dir, args)
    missing = [n for n, p in (("textbook", tb), ("captions", cap))
               if p is None or not p.exists()]
    if missing:
        print(f"[错误] 缺少输入：{', '.join(missing)}"
              f"（请传目录或 --textbook/--captions）", file=sys.stderr)
        return EXIT_USAGE
    if not img_dir.is_dir():
        print(f"[错误] images 目录不存在：{img_dir}", file=sys.stderr)
        return EXIT_USAGE

    try:
        stem = tb.stem[: -len("_textbook")] if tb.stem.endswith("_textbook") \
            else tb.stem
        textbook = parse_textbook(tb)
        entries_by_page = parse_text_entries(te) if te and te.exists() else {}
        manifest_by_page = parse_manifest(mf) if mf and mf.exists() else {}
        cap_by_file = parse_captions(cap)
        img_by_page = index_images(img_dir)

        # 断点续跑：解析已有 binding.json 的已完成页（有图片且关系+作用已生成）
        old_pages = {}
        out_path = Path(args.output) if args.output else \
            out_dir / f"{stem}_binding.json"
        if args.resume and out_path.exists():
            try:
                old = json.loads(out_path.read_text(encoding="utf-8"))
                for p in old.get("pages", []):
                    if p.get("images") and all(
                            im.get("relation") and im.get("position")
                            for im in p["images"]):
                        old_pages[p["page"]] = p
                if old_pages:
                    print(f"[续跑] 已有 {len(old_pages)} 页绑定完成，"
                          f"仅处理其余页", file=sys.stderr)
            except Exception:
                old_pages = {}

        # 图文位置作用 + 逻辑关系：DeepSeek 一次调用判断两项（role ≤40 字、
        # relation ≤50 字，功能角色与逻辑关系两个维度分离）
        relations, positions = {}, {}
        if not args.no_relation:
            api_key = os.environ.get(args.api_key_env, "")
            if not api_key:
                print(f"[警告] 未设置 {args.api_key_env}，relation/position "
                      f"回退默认值（--no-relation 可显式关闭）",
                      file=sys.stderr)
            else:
                from openai import OpenAI
                client = OpenAI(api_key=api_key, base_url=args.base_url)
                todo = [f for files in img_by_page.values() for f in files
                        if _page_of(f) not in old_pages]
                print(f"[关系] 生成 {len(todo)} 张图的位置作用与逻辑关系"
                      f"（模型 {args.model}）…", file=sys.stderr)
                for i, fname in enumerate(todo, start=1):
                    caption = cap_by_file.get(fname, {}).get("text", "")
                    if not caption:
                        continue
                    pg = _page_of(fname)
                    page_text = textbook.get(pg, "")[:1500] or \
                        " ".join(e["text"] for e in
                                 entries_by_page.get(pg, []))[:1500]
                    r = _judge_pos_rel(client, args.model, caption, page_text)
                    relations[fname] = r["relation"] or \
                        _fallback_relation(caption)
                    positions[fname] = r["position_role"] or ""
                    print(f"[关系 {i}/{len(todo)}] {fname}: "
                          f"relation={relations[fname][:26]}… | "
                          f"role={positions[fname][:26]}…",
                          file=sys.stderr)
                    time.sleep(0.2)

        binding = build_binding(stem, textbook, entries_by_page,
                                manifest_by_page, img_by_page,
                                cap_by_file, relations, positions)
        if old_pages:
            binding["pages"] = [
                old_pages.get(p["page"], p) for p in binding["pages"]]
            n_img = sum(len(p["images"]) for p in binding["pages"])
            n_bound = sum(1 for p in binding["pages"] if p["has_image"])
            n_link = sum(1 for p in binding["pages"] for im in p["images"]
                         if im.get("relation"))
            n_pos = sum(1 for p in binding["pages"] for im in p["images"]
                        if im.get("position"))
            binding["summary"] = {"pages": len(binding["pages"]),
                                  "images_total": n_img,
                                  "pages_with_image": n_bound,
                                  "relations": n_link, "positions": n_pos}
        out_path.write_text(
            json.dumps(binding, ensure_ascii=False, indent=1),
            encoding="utf-8")
        s = binding["summary"]
        n_rel = sum(1 for p in binding["pages"] for im in p["images"]
                    if im.get("relation"))
        n_pos = sum(1 for p in binding["pages"] for im in p["images"]
                    if im.get("position"))
        print(f"[OK] 图文绑定已写出：{out_path}")
        print(f"     {s['pages']} 页 / 图片 {s['images_total']} 张 / "
              f"含图页 {s['pages_with_image']} / 关系 {n_rel} 条 / "
              f"作用 {n_pos} 条")
        if args.json:
            print_json({"pages": s["pages"], "images": s["images_total"],
                        "bound_pages": s["pages_with_image"],
                        "relations": n_rel, "positions": n_pos,
                        "output": str(out_path)})
        return EXIT_OK
    except Exception as e:
        print(f"[错误] 图文绑定失败：{e}", file=sys.stderr)
        return EXIT_ERR


if __name__ == "__main__":
    sys.exit(_main())
