"""cli_related.py — pptx-related：图文相关性过滤（需求5）。

把 images/ 每张图的 qwen 诠释（captions.md）与该图所在页的正文文本
（texts.md）交给 DeepSeek 判断相关性；判定为无关的图（品牌 logo、作者/
教师姓名、单位全称与校徽、课题/项目编号、每页重复出现的页眉页脚装饰、
纯装饰分隔线、二维码、联系方式等）→ 从 images/、by_page/、captions.md
中删除该图及其解释，并写 <名>_related_filter.json 审计。

用法：
    pptx-related <产物目录> [-o captions.md]
                 [--images-dir dir] [--texts a.md]
                 [--model deepseek-v4-flash] [--base-url …] [--api-key-env …]
                 [--keep-all] [--json] [--version]

说明：
    - captions.md 被重写为仅保留相关条目（格式不变：### IMGxxxx — `file`）；
    - 被删记录写入 <目录>/<名>_related_filter.json（页号/文件/原因，可审计）；
    - 定位到 manifest.json 时同步标记 related="drop"，否则跳过（不报错）；
    - --keep-all 跳过过滤（调试用）。

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
                        quiet_stdout, banner, banner_end)

VERSION = "pptx-related 1.0.0 (图文相关性过滤)"
DEFAULT_MODEL = "deepseek-v4-flash"
DEFAULT_BASE_URL = "https://api.deepseek.com"
DEFAULT_KEY_ENV = "DEEPSEEK_API_KEY"

JUDGE_SYSTEM = (
    "你是高校教材内容审核专家。下面给出一张 PPT 图片的 AI 诠释与该图所在"
    "页的正文文本，请判断这张图片在**教材/课程知识**意义上是否与正文相关。\n"
    "判定为「不相关、应删除」的典型信号：品牌 logo、作者/教师姓名、单位"
    "全称与校徽、课题/项目编号、每页重复出现的页眉页脚装饰、纯装饰分隔线、"
    "二维码、联系方式等与知识点无关的内容。\n"
    "优先保留讲解知识点、原理、电路、结构、数据、流程、示例、波形、图表"
    "等内容的图片。\n"
    "**重要：如果正文为空或依据不足、无法确认相关性时，请输出 keep=true "
    "（默认保留，不得因无法确认而删除）。**\n"
    "只输出 JSON：{\"keep\": true 或 false, \"reason\": \"相关 | 具体不相关"
    "原因（如：logo/作者信息/单位名称/项目类别/每页重复装饰/其他）\"}"
)


def _page_of(file_name: str):
    m = re.match(r"slide_(\d+)_", file_name)
    return int(m.group(1)) if m else None


def _md_page_text(path: Path, page: int) -> str:
    """从 texts.md 提取第 page 页正文（## 第 N 页 块，含表格行）。"""
    if not path or not path.exists():
        return ""
    cur, buf = None, []
    for ln in path.read_text(encoding="utf-8").splitlines():
        m = re.match(r"^##\s*第\s*(\d+)\s*页\s*$", ln.strip())
        if m:
            cur = int(m.group(1))
            buf = []
        elif cur == page:
            buf.append(ln)
    lines = [x.strip() for x in buf if x.strip()]
    return "\n".join(lines)[:3000]


def parse_caption_entries(path: Path) -> list:
    """captions.md → [(img_id, file, text)]（顺序保留）。"""
    out, cur_id, cur_file, buf = [], None, None, []
    for ln in path.read_text(encoding="utf-8").splitlines():
        m = re.match(r"^###\s+(IMG\d+)\s*—\s*`?([^`\s]+\.\w+)\s*`?", ln)
        if m:
            if cur_file is not None:
                out.append((cur_id, cur_file,
                            " ".join(x.strip() for x in buf if x.strip())))
            cur_id, cur_file = m.group(1), m.group(2)
            buf = []
        elif cur_file is not None:
            buf.append(ln)
    if cur_file is not None:
        out.append((cur_id, cur_file,
                    " ".join(x.strip() for x in buf if x.strip())))
    return out


def _judge_one(client, model: str, caption: str, page_text: str,
               retries: int = 1) -> tuple:
    """DeepSeek 判定相关性 → (keep: bool, reason: str)。失败默认保留。
    v3.0.1：普通生成（去 thinking），相关性二分类无需深度思考。"""
    try:
        messages = [
            {"role": "system", "content": JUDGE_SYSTEM},
            {"role": "user", "content":
             f"【图片诠释】\n{caption[:800]}\n\n"
             f"【该页正文】\n{page_text[:1500]}\n\n"
             "请判断图片与正文是否相关，输出 JSON。"},
        ]
        last_err = None
        for attempt in range(retries + 1):
            try:
                print(f"[DeepSeek] 相关性判定输入（第 {attempt + 1} 次）："
                      f"{caption[:200]}…", file=sys.stderr)
                resp = client.chat.completions.create(
                    model=model, messages=messages, stream=False,
                )
                s = (resp.choices[0].message.content or "").strip()
                print(f"[DeepSeek] 相关性判定输出：{s[:200]}…", file=sys.stderr)
                m = re.search(r"\{.*\}", s, re.S)
                if m:
                    data = json.loads(m.group(0))
                    keep = bool(data.get("keep", True))
                    reason = str(data.get("reason", "") or "")[:40]
                    return keep, reason
                return True, "无法解析判定（默认保留）"
            except Exception as e:
                last_err = e
                if attempt < retries:
                    time.sleep(2 * (attempt + 1))
        print(f"[警告] 判定调用失败({last_err})，默认保留",
              file=sys.stderr)
    except Exception:
        pass
    return True, "判定异常（默认保留）"


def _load_page_texts(path: Path) -> dict:
    """把 texts.md 一次解析为 {页号: 正文}（v3.0.1：避免每条目重复读文件）。"""
    out = {}
    if not path or not path.exists():
        return out
    cur, buf = None, []
    for ln in path.read_text(encoding="utf-8").splitlines():
        m = re.match(r"^##\s*第\s*(\d+)\s*页\s*$", ln.strip())
        if m:
            if cur is not None:
                out[cur] = "\n".join(x.strip() for x in buf if x.strip())
            cur = int(m.group(1))
            buf = []
        elif cur is not None:
            buf.append(ln)
    if cur is not None:
        out[cur] = "\n".join(x.strip() for x in buf if x.strip())
    return out


def _rewrite_captions(path: Path, entries: list, n_del: int) -> None:
    """重写 captions.md：仅保留相关条目，尾部追加过滤统计。"""
    lines = [f"# images 图片 AI 解读（已过滤无关图）", "",
             f"> 由 `pptx-caption` 生成、`pptx-related` 过滤："
             f"原 {n_del + len(entries)} 条，保留 {len(entries)} 条，"
             f"删除 {n_del} 条（明细见 *_related_filter.json）。", ""]
    for img_id, file, text in entries:
        lines.append(f"### {img_id} — `{file}`  ✅")
        lines.append("")
        lines.append(text)
        lines.append("")
    lines.append("---")
    lines.append("")
    lines.append(f"过滤统计：保留 {len(entries)} 条，删除 {n_del} 条。")
    path.write_text("\n".join(lines), encoding="utf-8")


def _mark_manifest_related(manifest_path: Path, dropped: dict) -> bool:
    """在 manifest.json 的对应记录上标记 related='drop'（溯源用）。"""
    try:
        recs = json.loads(manifest_path.read_text(encoding="utf-8"))
        changed = 0
        for r in recs:
            if r.get("output_file") in dropped:
                r["related"] = "drop"
                changed += 1
        if changed:
            manifest_path.write_text(
                json.dumps(recs, ensure_ascii=False, indent=2),
                encoding="utf-8")
        return True
    except Exception:
        return False


def related_filter(captions_path: Path, texts_path: Path,
                   images_dir: Path, by_page_dir: Path | None,
                   out_dir: Path, model: str = DEFAULT_MODEL,
                   base_url: str = DEFAULT_BASE_URL,
                   api_key_env: str = DEFAULT_KEY_ENV,
                   keep_all: bool = False, on_progress=None) -> dict:
    """执行相关性过滤。返回统计 dict。

    v3.0.1：① 并发判定（ThreadPoolExecutor，去 thinking 普通生成）；
    ② 页面正文为空 → 保守保留（不调模型，避免误删）；
    ③ texts.md 预加载为 {页: 正文}（不再每条目重读文件）。"""
    entries = parse_caption_entries(captions_path)
    n_total = len(entries)
    if n_total == 0:
        print("[提示] captions.md 无条目，无需过滤", file=sys.stderr)
        return {"total": 0, "kept": 0, "dropped": 0, "reasons": {}}

    kept, dropped = [], {}
    reasons = {}
    delete_failed = 0
    api_key = os.environ.get(api_key_env, "")
    client = None
    if keep_all or not api_key:
        if not api_key and not keep_all:
            print(f"[警告] 未设置 {api_key_env}，跳过过滤（--keep-all 等效）",
                  file=sys.stderr)
    else:
        from openai import OpenAI
        client = OpenAI(api_key=api_key, base_url=base_url)

    # 预加载正文（v3.0.1）
    page_texts = _load_page_texts(texts_path)

    # 并发判定：返回 (img_id, file, text, keep, reason)
    from concurrent.futures import ThreadPoolExecutor, as_completed

    def _work(item):
        img_id, file, text = item
        pg = _page_of(file)
        page_text = page_texts.get(pg, "") if pg else ""
        if client is None:
            return (img_id, file, text, True, "跳过（未配置模型/keep-all）")
        if not page_text.strip():
            # 正文为空：无法判断相关性 → 保守保留（不调模型，避免误删）
            print(f"[保留] {img_id} `{file}` 页面正文为空，无法判断，保守保留",
                  file=sys.stderr)
            return (img_id, file, text, True, "页面正文为空，无法判断，保守保留")
        keep, reason = _judge_one(client, model, text, page_text)
        return (img_id, file, text, keep, reason)

    done = 0
    if client is not None:
        with ThreadPoolExecutor(max_workers=8) as ex:
            futs = [ex.submit(_work, e) for e in entries]
            for fut in as_completed(futs):
                done += 1
                if on_progress is not None:
                    try:
                        on_progress(done, n_total, {"kind": "related"})
                    except Exception:
                        pass
                img_id, file, text, keep, reason = fut.result()
                if keep:
                    kept.append((img_id, file, text))
                else:
                    dropped[file] = {"page": _page_of(file), "reason": reason,
                                     "caption": text[:120]}
                    reasons[reason or "其他"] = \
                        reasons.get(reason or "其他", 0) + 1
                    # 删除 images/ 与 by_page/ 下的对应文件（失败告警）
                    for d in (images_dir, by_page_dir):
                        if d and d.is_dir():
                            f = d / file
                            try:
                                if f.is_file():
                                    f.unlink()
                            except OSError as e:
                                delete_failed += 1
                                print(f"[警告] 删除失败（可能被安全删除机制拦截）："
                                      f"{f}（{e}）", file=sys.stderr)
                    print(f"[删除] {img_id} `{file}` 与正文无关（{reason}）",
                          file=sys.stderr)
    else:
        # 无模型：全部保留（keep-all 或未配置 Key）
        for img_id, file, text in entries:
            done += 1
            if on_progress is not None:
                try:
                    on_progress(done, n_total, {"kind": "related"})
                except Exception:
                    pass
            kept.append((img_id, file, text))

    # 重写 captions.md（保留相关条目）
    _rewrite_captions(captions_path, kept, len(dropped))

    # 审计文件
    stem = captions_path.name[: -len("_captions.md")] \
        if captions_path.name.endswith("_captions.md") else captions_path.stem
    report = {"stem": stem, "model": model,
              "total": n_total, "kept": len(kept), "dropped": len(dropped),
              "delete_failed": delete_failed,
              "reasons": reasons, "dropped_items": dropped}
    report_path = out_dir / f"{stem}_related_filter.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=1),
                           encoding="utf-8")
    return {"total": n_total, "kept": len(kept), "dropped": len(dropped),
            "delete_failed": delete_failed,
            "reasons": reasons, "report": str(report_path),
            "dropped_files": dropped}


def _main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        prog="pptx-related",
        description="图文相关性过滤：DeepSeek 判定图片与正文是否相关，"
                    "无关图（logo/作者/单位/装饰等）从图片集与解读中删除")
    ap.add_argument("dir", nargs="?", default=".",
                    help="产物目录（自动找 captions.md / texts.md / images）")
    ap.add_argument("-o", "--output", type=Path, default=None,
                    help="captions.md 路径（默认自动查找并原地重写）")
    ap.add_argument("--texts", type=Path, default=None, help="文本 md")
    ap.add_argument("--images-dir", type=Path, default=None, help="images 目录")
    ap.add_argument("--model", default=DEFAULT_MODEL,
                    help=f"判定模型（默认 {DEFAULT_MODEL}）")
    ap.add_argument("--base-url", default=DEFAULT_BASE_URL,
                    help=f"OpenAI 兼容端点（默认 {DEFAULT_BASE_URL}）")
    ap.add_argument("--api-key-env", default=DEFAULT_KEY_ENV,
                    help=f"API Key 环境变量名（默认 {DEFAULT_KEY_ENV}）")
    ap.add_argument("--keep-all", action="store_true",
                    help="跳过过滤（调试用，全部保留）")
    ap.add_argument("--json", action="store_true",
                    help="结构化统计输出到 stdout")
    ap.add_argument("--version", action="version", version=VERSION)
    args = ap.parse_args(argv)

    out_dir = Path(args.dir)
    cap = args.output
    if cap is None:
        cands = sorted(out_dir.glob("*_captions.md")) + \
            sorted(out_dir.glob("images_captions.md"))
        cap = cands[0] if cands else None
    if cap is None or not cap.exists():
        print("[错误] 未找到 captions.md（请传目录或 -o 指定）",
              file=sys.stderr)
        return EXIT_USAGE
    texts = args.texts
    if texts is None:
        stem = cap.name[: -len("_captions.md")] \
            if cap.name.endswith("_captions.md") else ""
        cands = sorted(out_dir.glob(f"{stem}_texts.md")) if stem else []
        texts = cands[0] if cands else None
    images_dir = args.images_dir or (out_dir / "images")
    if not images_dir.is_dir():
        images_dir = out_dir / "过程文件" / "img" / "images"

    try:
        # by_page 定位（过程文件/…/by_page 或 out/by_page）
        by_page_dir = None
        for cand in (out_dir / "by_page",
                     out_dir / "过程文件" / "img" / "by_page"):
            if cand.is_dir():
                by_page_dir = cand
                break
        cb = None
        if args.json:
            with quiet_stdout():
                result = related_filter(cap, texts, images_dir, by_page_dir,
                                        out_dir, model=args.model,
                                        base_url=args.base_url,
                                        api_key_env=args.api_key_env,
                                        keep_all=args.keep_all,
                                        on_progress=cb)
            # 标记 manifest（尽力而为，不影响主流程）
            mf_cands = sorted(out_dir.glob("过程文件/**/manifest.json")) + \
                sorted(out_dir.glob("manifest.json"))
            if mf_cands:
                _mark_manifest_related(mf_cands[0],
                                       result.get("dropped_files", {}))
            print_json(result)
        else:
            result = related_filter(cap, texts, images_dir, by_page_dir,
                                    out_dir, model=args.model,
                                    base_url=args.base_url,
                                    api_key_env=args.api_key_env,
                                    keep_all=args.keep_all,
                                    on_progress=cb)
            print(f"[OK] 图文相关性过滤完成：{cap}")
            print(f"     共 {result['total']} 条，保留 {result['kept']}，"
                  f"删除 {result['dropped']} 条")
            if result["reasons"]:
                for r, n in result["reasons"].items():
                    print(f"       - {r}: {n}")
            print(f"     审计：{result['report']}")
        return EXIT_OK
    except Exception as e:
        print(f"[错误] 相关性过滤失败：{e}", file=sys.stderr)
        return EXIT_ERR


def main() -> int:  # console
    banner("pptx-related")
    rc = _main()
    banner_end("pptx-related")
    return rc


if __name__ == "__main__":
    sys.exit(_main())
