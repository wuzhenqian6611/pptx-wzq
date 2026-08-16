"""cli_caption.py — pptx-caption 薄壳：把 images 目录图片逐张喂入
视觉大模型做 AI 解读，生成 <名>_captions.md 文档。

用法：
    pptx-caption <images_dir> [-o captions.md] [--model <视觉模型>]
                 [--base-url <兼容端点>] [--api-key-env <变量名>]
                 [--sleep 0.5] [--json] [--version]

流程：
    遍历 images_dir 栅格图（PNG/JPG/BMP/GIF/WEBP/TIF）→ base64 →
    OpenAI 兼容接口（content 含 image_url data URL）→ 逐张调用 →
    解读写入 md。

说明：
    - 默认对接阿里云百炼（DashScope）：模型 qwen3.7-plus，端点
      https://dashscope.aliyuncs.com/compatible-mode/v1，Key 从环境变量
      DASHSCOPE_API_KEY 读取（用户自行 setx 注册）；
    - 换其他 OpenAI 兼容视觉端点用 --base-url / --api-key-env / --model，
      例如：
        智谱 GLM-4V：  --base-url https://open.bigmodel.cn/api/paas/v4
                       --model glm-4v-flash --api-key-env ZHIPU_API_KEY
        通义 omni：    --model qwen3.5-omni-flash --stream
    - 注意：DeepSeek 官方 API（api.deepseek.com）全系为纯文本模型，
      **不支持图片输入**，不可用于本指令；
    - 部分 omni 模型要求流式响应，用 --stream（聚合 delta.content）；
    - 矢量图（WMF/EMF/SVG）无法喂入多模态模型，跳过并注明；
    - 单张调用前先做 PIL 预检：损坏/无法解码的图片直接标记跳过，
      不调用 API（避免浪费额度）；
    - 每张解读结果**打印到屏幕**；并每 --flush-interval 秒（默认 60）
      **增量落盘**到 md（追加写，中途中断不丢已解读结果），
      全部完成后追加尾部统计；
    - 单张失败自动重试 2 次，仍失败记录错误后继续（--stop-on-error 可中断）。

退出码：0 成功 / 1 处理异常 / 2 参数或环境错误。

作者：吴振谦 · wuzhenqian@nbu.edu.cn · QQ：38328063"""
from __future__ import annotations

import argparse
import base64
import os
import re
import sys
import time
from pathlib import Path

from pptx_wzq.cli_common import (EXIT_ERR, EXIT_OK, EXIT_USAGE,
                        make_progress, print_json, quiet_stdout,
                        banner, banner_end)

VERSION = "pptx-caption 1.5.0 (方案B薄壳)"
DEFAULT_MODEL = "qwen3.7-plus"
DEFAULT_BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"
DEFAULT_KEY_ENV = "DASHSCOPE_API_KEY"
RASTER_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp", ".tif", ".tiff"}
VECTOR_EXTS = {".wmf", ".emf", ".svg"}

# 文档上下文模式提示词（用户模板；{subject} 由 texts 前 3-5 页让模型生成）
DOC_SYSTEM_PROMPT_TMPL = (
    "你是高校教材建设专家、高级人工智能多模态知识库工程师、"
    "和{subject}专业资深教师，现给你输两个md文档，一个是从PPT中提取的"
    "文本文档，一个是从PPT中提取的公式文档，你先理解这些文档，然后我"
    "再给一个接一个给你每页PPT中的图片，你给出这些图片从教材或者课程"
    "知识内容角度的理解，每个图大概100-200字左右。\n"
    "输出请按「图片类型；内容理解（教材/课程角度）；教学用途」组织，"
    "总字数 100-200 字。"
)

# 兼容旧用法（无 --texts/--formulas 时的通用视觉提示词）
SYSTEM_PROMPT = (
    "你是高校教材建设领域的高级 AI 视觉解读专家。请解读这张图片，输出：\n"
    "1) 图片类型（照片/框图/电路图/结构图/设计图/波形图/字符符号/其他）；\n"
    "2) 主要内容（这张图讲了什么，包含哪些关键元素/器件/信号）；\n"
    "3) 教学用途（适合放在教材或课件的哪个位置，讲解什么知识点）。\n"
    "用中文回答，总字数 80-150 字，条理清晰。"
)


def _page_of(file_name: str):
    """从文件名 slide_NN_... 提取页码；失败返回 None。"""
    m = re.match(r"slide_(\d+)_", file_name)
    return int(m.group(1)) if m else None


def _md_page_text(path: Path, page: int) -> str:
    """从 texts/formulas md 中提取第 page 页的正文文本（## 第 N 页 块）。"""
    if not path or not path.exists():
        return ""
    cur, buf = None, []
    for ln in path.read_text(encoding="utf-8").splitlines():
        m = re.match(r"^##\s*第\s*(\d+)\s*(?:页|节)\s*$", ln.strip())
        if m:
            cur = int(m.group(1))
            buf = []
        elif cur == page:
            buf.append(ln)
    lines = [x.strip() for x in buf if x.strip()]
    return "\n".join(lines)[:3000]


def _infer_subject(client, texts_path: Path, model: str,
                   base_url: str, retries: int = 2) -> str:
    """用当前模型从 texts.md 前 3-5 页判断课程/专业名称。"""
    import re as _re
    try:
        pages = {}
        cur = None
        for ln in texts_path.read_text(encoding="utf-8").splitlines():
            m = _re.match(r"^##\s*第\s*(\d+)\s*(?:页|节)\s*$", ln.strip())
            if m:
                cur = int(m.group(1))
                pages.setdefault(cur, [])
            elif cur is not None:
                pages[cur].append(ln)
        first = sorted(pages)[:5]
        snippet = "\n".join(
            f"## 第 {p} 页\n" + "\n".join(pages[p]) for p in first)[:6000]
        messages = [
            {"role": "system", "content":
             "你是高校教材建设专家。根据以下 PPT 前几页的文本内容，"
             "判断这门课程属于哪个专业/课程名称。只输出课程或专业名称"
             "（如：模拟电子技术、高等数学、计算机组成原理等），"
             "不要输出任何其他内容。"},
            {"role": "user", "content":
             f"以下是该课程前几页的文本内容：\n\n{snippet}\n\n"
             "请判断课程/专业名称，只输出名称。"},
        ]
        last_err = None
        for attempt in range(retries + 1):
            try:
                resp = client.chat.completions.create(
                    model=model, messages=messages, stream=False,
                    reasoning_effort="high",
                    extra_body={"thinking": {"type": "enabled"}},
                )
                name = (resp.choices[0].message.content or "").strip()
                for ln in name.splitlines():
                    ln = ln.strip(" -#*·、.。")
                    if ln:
                        return ln[:30]
                return name[:30]
            except Exception as e:
                last_err = e
                if attempt < retries:
                    time.sleep(2 * (attempt + 1))
        print(f"[警告] 学科判断失败({last_err})，使用默认名",
              file=sys.stderr)
    except Exception as e:
        print(f"[警告] 学科判断异常({e})，使用默认名", file=sys.stderr)
    return "电子信息"


def _check_env(args) -> bool:
    """环境检查：openai SDK + API Key 环境变量 + 目录。返回是否可继续。"""
    ok = True
    try:
        import openai  # noqa
        print("[环境] openai SDK: OK", file=sys.stderr)
    except Exception:
        print("[环境] openai SDK: 缺失（pip install openai）", file=sys.stderr)
        ok = False
    key = os.environ.get(args.api_key_env)
    if key:
        print(f"[环境] {args.api_key_env}: 已设置", file=sys.stderr)
    else:
        print(f"[环境] {args.api_key_env}: 未设置（请先执行 "
              f"setx {args.api_key_env} 你的Key，再重开命令窗口）",
              file=sys.stderr)
        ok = False
    if not args.images_dir.exists() or not args.images_dir.is_dir():
        print(f"[环境] 图片目录不存在：{args.images_dir}", file=sys.stderr)
        ok = False
    return ok


def _encode_image(path: Path) -> str | None:
    try:
        b64 = base64.b64encode(path.read_bytes()).decode("ascii")
        ext = path.suffix.lower().lstrip(".")
        mime = {"jpg": "jpeg", "jpeg": "jpeg", "tif": "tiff",
                "tiff": "tiff"}.get(ext, ext)
        return f"data:image/{mime};base64,{b64}"
    except Exception:
        return None


def _caption_one(client, path: Path, model: str,
                 api_key: str, retries: int = 2,
                 stream: bool = False,
                 system_prompt: str = None,
                 user_text: str = None) -> str:
    """单张图喂入大模型，返回解读文本；重试 retries 次。

    system_prompt / user_text：文档上下文模式（--texts/--formulas）下
    传入自定义 system 与该页文本/公式上下文；默认用通用视觉提示词。
    stream=True 用于部分 omni 模型（如 qwen3.5-omni-flash）要求流式
    响应的场景：聚合各 chunk 的 delta.content。
    """
    data_url = _encode_image(path)
    if data_url is None:
        raise RuntimeError("图片编码失败")
    if system_prompt is None:
        system_prompt = SYSTEM_PROMPT
    user_text = user_text or "请解读这张图片。"
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": [
            {"type": "text", "text": user_text},
            {"type": "image_url", "image_url": {"url": data_url}},
        ]},
    ]
    last_err = None
    for attempt in range(retries + 1):
        try:
            if stream:
                resp = client.chat.completions.create(
                    model=model, messages=messages, stream=True,
                    stream_options={"include_usage": True})
                parts = []
                for chunk in resp:
                    if chunk.choices and chunk.choices[0].delta \
                            and chunk.choices[0].delta.content:
                        parts.append(chunk.choices[0].delta.content)
                text = "".join(parts).strip()
                if text:
                    return text
                raise RuntimeError("流式响应为空")
            resp = client.chat.completions.create(
                model=model,
                messages=messages,
                stream=False,
                reasoning_effort="high",
                extra_body={"thinking": {"type": "enabled"}},
            )
            return (resp.choices[0].message.content or "").strip()
        except Exception as e:
            last_err = e
            if attempt < retries:
                time.sleep(2 * (attempt + 1))
    raise RuntimeError(f"API 调用失败: {last_err}")


def _write_header(out_path: Path, title: str, model: str,
                  n: int, n_vec: int) -> None:
    """首次创建 md：标题 + 说明头。"""
    lines = [f"# {title} 图片 AI 解读", "",
             f"> 由 `pptx-caption` 生成 · 模型 `{model}` · "
             f"共解读 {n} 张栅格图" +
             (f"（另有 {n_vec} 张矢量图无法喂入模型，跳过）"
              if n_vec else "") + "。",
             "> 增量落盘：每 60 秒把已完成条目追加写入本文件，"
             "中途中断不丢已解读结果。", ""]
    out_path.write_text("\n".join(lines), encoding="utf-8")


def _append_entries(out_path: Path, entries) -> None:
    """把一批 (img_id, name, status, text) 条目追加写入 md（UTF-8）。"""
    with open(out_path, "a", encoding="utf-8") as f:
        for img_id, name, status, text in entries:
            f.write(f"### {img_id} — `{name}`  {status}\n\n")
            f.write(f"{text}\n\n")


def _append_footer(out_path: Path, n: int, done: int, failed: int,
                   model: str, base_url: str,
                   vectors: list | None = None,
                   note: str = "") -> None:
    """末尾追加：矢量跳过清单 + 统计（仅在全部完成后写一次）。"""
    with open(out_path, "a", encoding="utf-8") as f:
        if vectors:
            f.write("## 跳过（矢量图）\n\n")
            for v in vectors:
                f.write(f"- `{v.name}`（矢量图，无法喂入多模态模型）\n")
            f.write("\n")
        f.write("---\n\n")
        f.write(f"解读统计：共 {n} 张栅格图，成功 {done} 张，失败 {failed} 张。\n")
        f.write(f"- 模型：{model}；API：{base_url}\n")
        if note:
            f.write(f"- {note}\n")


def _existing_entries(out_path: Path) -> set:
    """解析已有 captions.md 的条目文件名集合（--resume 续跑用）。"""
    done = set()
    if not out_path.exists():
        return done
    for ln in out_path.read_text(encoding="utf-8").splitlines():
        m = re.match(r"^###\s+IMG\d+\s*—\s*`?([^`\s]+\.\w+)\s*`?", ln)
        if m:
            done.add(m.group(1))
    return done


def caption_dir(images_dir: Path, out_path: Path,
                model: str = DEFAULT_MODEL,
                base_url: str = DEFAULT_BASE_URL,
                api_key_env: str = DEFAULT_KEY_ENV,
                sleep: float = 0.5, stop_on_error: bool = False,
                stream: bool = False,
                flush_interval: float = 60.0,
                texts_path: Path = None, formulas_path: Path = None,
                subject: str = None,
                resume: bool = False,
                on_progress=None):
    """逐张解读 images_dir 下栅格图，写 captions.md。

    texts_path/formulas_path：文档上下文模式——学科（subject 或由
    texts 前 3-5 页让模型生成）填入提示词括号，每张图附带该页文本/
    公式上下文（从文件名 slide_NN 取页码）。不提供时用通用视觉提示词。
    resume=True：解析已有 captions.md 条目，跳过已完成图片（断点续跑，
    不重写文件头，只追加新条目）。
    其他特性：屏幕打印、增量落盘、矢量跳过、失败重试。
    返回统计 dict。
    """
    api_key = os.environ.get(api_key_env, "")
    from openai import OpenAI
    client = OpenAI(api_key=api_key, base_url=base_url)

    out_path.parent.mkdir(parents=True, exist_ok=True)   # 输出目录可不存在

    images = sorted(
        p for p in images_dir.iterdir()
        if p.suffix.lower() in RASTER_EXTS)
    vectors = sorted(p for p in images_dir.iterdir()
                     if p.suffix.lower() in VECTOR_EXTS)
    resumed = 0
    if resume and out_path.exists():
        done_files = _existing_entries(out_path)
        todo = [p for p in images if p.name not in done_files]
        resumed = len(images) - len(todo)
        if resumed:
            print(f"[续跑] 已解读 {resumed} 张（{out_path.name}），"
                  f"继续剩余 {len(todo)} 张", file=sys.stderr)
        images = todo
    n = len(images)

    # 文档上下文模式：先判断学科（texts 前 3-5 页 → 模型），再逐图带页上下文
    doc_mode = texts_path is not None and texts_path.exists()
    if doc_mode:
        if not subject:
            subject = _infer_subject(client, texts_path, model, base_url)
        print(f"[学科] {subject}（文档上下文模式，"
              f"每图附带该页文本/公式上下文）", file=sys.stderr)
    else:
        subject = ""

    if not (resume and out_path.exists()):
        _write_header(out_path, "images",
                      model, n, len(vectors))
        print(f"[落盘] 已创建 {out_path.name}（共 {n} 张）", file=sys.stderr)
    else:
        print(f"[落盘] 续跑模式：追加写入 {out_path.name}", file=sys.stderr)

    done, failed = 0, 0
    pending = []                     # 本分钟待落盘条目
    last_flush = time.time()

    def _flush(reason: str) -> None:
        nonlocal pending, last_flush
        if pending:
            _append_entries(out_path, pending)
            print(f"[落盘] {reason}：追加 {len(pending)} 条到 "
                  f"{out_path.name}", file=sys.stderr)
            pending = []
        last_flush = time.time()

    for i, p in enumerate(images):
        if on_progress is not None:
            try:
                on_progress(i + 1, n, {"kind": "caption"})
            except Exception:
                pass
        img_id = f"IMG{i + 1:04d}"
        # 预检：图片能否解码（损坏图不调用 API，避免浪费额度）
        try:
            from PIL import Image
            with Image.open(p) as im:
                im.verify()
        except Exception:
            failed += 1
            text = "（图片损坏/无法解码，已跳过，未调用 API）"
            status = "⚠️"
            print(f"[{img_id}] {p.name}  {status} {text}",
                  file=sys.stderr)
            pending.append((img_id, p.name, status, text))
            if time.time() - last_flush >= flush_interval:
                _flush("定时")
            continue
        # 构造本条消息（文档上下文模式：该页文本/公式 + 图片）
        system_prompt, user_text = None, None
        if doc_mode:
            page = _page_of(p.name)
            seg = [f"这是第 {page} 页 PPT 的图片，该页内容如下："]
            txt = _md_page_text(texts_path, page) if texts_path else ""
            if txt:
                seg.append("【文本】\n" + txt)
            if formulas_path:
                fm = _md_page_text(formulas_path, page)
                if fm:
                    seg.append("【公式】\n" + fm)
            seg.append("请从教材/课程知识内容角度理解这张图片，"
                       "100-200 字。")
            system_prompt = DOC_SYSTEM_PROMPT_TMPL.format(subject=subject)
            user_text = "\n".join(seg)
        try:
            text = _caption_one(client, p, model, api_key, stream=stream,
                                system_prompt=system_prompt,
                                user_text=user_text)
            done += 1
            status = "✅"
        except Exception as e:
            text = f"（解读失败：{e}）"
            failed += 1
            status = "❌"
            if stop_on_error:
                pending.append((img_id, p.name, status, text))
                _flush("停止前")
                _append_footer(out_path, n, done, failed, model, base_url,
                               note=f"停止于第 {i + 1}/{n} 张（--stop-on-error）。")
                return {"total": n, "done": done, "failed": failed,
                        "model": model, "md": str(out_path)}
        # 屏幕打印本张解读结果（与进度同流，--json 模式也可见）
        print(f"[{img_id}] {p.name}  {status}", file=sys.stderr)
        print(text, file=sys.stderr)
        print("", file=sys.stderr)
        pending.append((img_id, p.name, status, text))
        if time.time() - last_flush >= flush_interval:
            _flush("定时")
        if sleep and i < n - 1:
            time.sleep(sleep)

    _flush("完成")
    _append_footer(out_path, n + resumed, done, failed, model, base_url,
                   vectors,
                   note=f"（本次续跑 {resumed} 张已完成，新处理 {n} 张）"
                        if resumed else "")
    return {"total": n, "done": done, "failed": failed,
            "resumed": resumed, "model": model, "md": str(out_path)}


def _main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        prog="pptx-caption",
        description="把 images 目录图片逐张喂入视觉大模型做 AI 解读")
    ap.add_argument("images_dir", help="图片目录（如 out/images）")
    ap.add_argument("-o", "--output", default=None,
                    help="输出 md 路径（默认 <目录>_captions.md）")
    ap.add_argument("--model", default=DEFAULT_MODEL,
                    help=f"视觉模型名（默认 {DEFAULT_MODEL}）")
    ap.add_argument("--base-url", default=DEFAULT_BASE_URL,
                    help=f"OpenAI 兼容端点（默认 {DEFAULT_BASE_URL}；"
                         f"换端点示例见文件头注释）")
    ap.add_argument("--api-key-env", default=DEFAULT_KEY_ENV,
                    help=f"API Key 所在环境变量名（默认 {DEFAULT_KEY_ENV}）")
    ap.add_argument("--stream", action="store_true",
                    help="用流式响应调用（部分 omni 模型如 qwen3.5-omni-flash"
                         "要求 stream=True，默认关）")
    ap.add_argument("--sleep", type=float, default=0.5,
                    help="每张调用间隔秒数（默认 0.5，避免触发限速）")
    ap.add_argument("--flush-interval", type=float, default=60.0,
                    help="增量落盘间隔秒数：每隔该时长把当前完成的条目追加"
                         "写入 md（默认 60；中途中断不丢已解读结果）")
    ap.add_argument("--stop-on-error", action="store_true",
                    help="单张失败即中断（默认跳过继续）")
    ap.add_argument("--resume", action="store_true",
                    help="断点续跑：解析已有 captions.md 条目，跳过已完成"
                         "图片，只解读剩余（不重写文件头）")
    ap.add_argument("--texts", type=Path, default=None,
                    help="文本文档 md（提供后启用「先理解文档」模式："
                         "学科由 texts 前 3-5 页让模型生成，每图附带该页"
                         "文本/公式上下文）")
    ap.add_argument("--formulas", type=Path, default=None,
                    help="公式文档 md（配合 --texts 使用）")
    ap.add_argument("--subject", default=None,
                    help="课程/专业名称（默认由 --texts 前 3-5 页让模型生成）")
    ap.add_argument("--json", action="store_true",
                    help="结构化统计输出到 stdout")
    ap.add_argument("--version", action="version", version=VERSION)
    args = ap.parse_args(argv)

    images_dir = Path(args.images_dir)
    args.images_dir = images_dir
    if not _check_env(args):
        return EXIT_USAGE
    out_path = Path(args.output) if args.output else \
        images_dir.parent / f"{images_dir.name}_captions.md"

    try:
        cb = make_progress("图片解读")
        if args.json:
            with quiet_stdout():
                result = caption_dir(images_dir, out_path,
                                     model=args.model, base_url=args.base_url,
                                     api_key_env=args.api_key_env,
                                     sleep=args.sleep,
                                     stop_on_error=args.stop_on_error,
                                     stream=args.stream,
                                     flush_interval=args.flush_interval,
                                     texts_path=args.texts,
                                     formulas_path=args.formulas,
                                     subject=args.subject,
                                     resume=args.resume,
                                     on_progress=cb)
            print_json(result)
        else:
            result = caption_dir(images_dir, out_path,
                                 model=args.model, base_url=args.base_url,
                                 api_key_env=args.api_key_env,
                                 sleep=args.sleep,
                                 stop_on_error=args.stop_on_error,
                                 stream=args.stream,
                                 flush_interval=args.flush_interval,
                                 texts_path=args.texts,
                                 formulas_path=args.formulas,
                                 subject=args.subject,
                                 resume=args.resume,
                                 on_progress=cb)
            print(f"[OK] AI 解读完成：{out_path}")
            print(f"     成功 {result['done']} / 共 {result['total']} 张"
                  f"{'（失败 ' + str(result['failed']) + '）' if result['failed'] else ''}"
                  + (f"（续跑 {result['resumed']} 张已完成）"
                     if result.get("resumed") else ""))
        return EXIT_OK
    except Exception as e:
        print(f"[错误] 解读失败：{e}", file=sys.stderr)
        return EXIT_ERR


def main() -> int:  # console
    banner("pptx-caption")
    rc = _main()
    banner_end("pptx-caption")
    return rc


if __name__ == "__main__":
    sys.exit(_main())
