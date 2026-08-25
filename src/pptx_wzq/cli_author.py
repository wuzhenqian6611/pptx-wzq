"""cli_author.py — pptx-author 薄壳：把 PPT 提取的文本/公式/图片解读
三个 md 文档一次性输入 DeepSeek 大模型，系统性生成整份教材文案 md。

用法：
    pptx-author <产物目录> [-o out.md] [--subject 学科]
                [--texts a.md --formulas b.md --captions c.md]
                [--model deepseek-v4-flash] [--api-key-env DEEPSEEK_API_KEY]
                [--base-url https://api.deepseek.com]
                [--pages "1-5,8"]
                [--max-input-chars 48000] [--max-pages-per-batch 20]
                [--max-output-chars 12000]
                [--json] [--version]

流程：
    1) 自动在目录内找 <名>_texts.md / <名>_formulas.md /
       <名>_captions.md（或 --texts/--formulas/--captions 显式指定）；
    2) 按页聚合三个来源（texts/formulas 自带页码，captions 从图片
       文件名 slide_NN_pic 提取页码）；
    3) 学科：默认用 DeepSeek 判断（取 texts.md 前三页文本），
       失败回退关键词表；--subject 可手动覆盖；
    4) 三个文档**全部内容输入模型**一次生成整份教材文案 md
       （系统性更好、不遗漏公式/图片解读）；
    5) **自适应分批**：输入/估算输出超过模型限制（--max-input-chars /
       --max-pages-per-batch / --max-output-chars）时，按页自动拆批，
       每批一次调用，各批结果顺序合并到同一个 <名>_textbook.md。

说明：
    - 本指令为纯文本任务，默认走 DeepSeek（api.deepseek.com，
      deepseek-v4-flash，thinking 开启）；Key 从 DEEPSEEK_API_KEY 读取；
    - 换 OpenAI 兼容文本模型用 --base-url/--api-key-env/--model；
    - --pages 用于小批量测试：只把目标页内容传给模型并只生成这些页。

退出码：0 成功 / 1 处理异常 / 2 参数或环境错误。

作者：吴振谦 · wuzhenqian@nbu.edu.cn · QQ：38328063"""
from __future__ import annotations

import argparse
import os
import re
import sys
import time
from pathlib import Path

from pptx_wzq.cli_common import (EXIT_ERR, EXIT_OK, EXIT_USAGE,
                        make_progress, print_json, quiet_stdout,
                        banner, banner_end)

VERSION = "pptx-author 1.2.0 (方案B薄壳)"
DEFAULT_MODEL = "deepseek-v4-flash"
DEFAULT_BASE_URL = "https://api.deepseek.com"
DEFAULT_KEY_ENV = "DEEPSEEK_API_KEY"
MIN_CHARS = 500  # 每页教材文案目标字数下限（v2.1：每页至少扩写到 500 字）

# 学科关键词表（自动推断用；--subject 可覆盖）
SUBJECT_KEYWORDS = {
    "电子信息工程/电子技术": ["放大电路", "电阻", "电容", "晶体管", "运放",
                          "放大器", "信号", "频率", "半导体", "波形", "电流",
                          "电压", "反馈", "振荡", "滤波", "二极管", "场效应",
                          "共射", "静态工作点", "偏置"],
    "数学": ["函数", "导数", "积分", "方程", "矩阵", "概率", "极限",
           "微分", "级数", "向量", "几何"],
    "物理学": ["力学", "电场", "磁场", "量子", "热学", "光学", "电磁感应",
            "牛顿", "能量守恒", "波动"],
    "计算机科学与技术": ["算法", "数据结构", "操作系统", "编程", "数据库",
                    "网络协议", "编译器", "进程", "线程", "机器学习"],
    "自动控制": ["控制系统", "传递函数", "稳定性", "闭环", "开环", "PID",
              "状态空间", "根轨迹", "频域"],
    "经济学": ["需求", "供给", "市场", "价格", "成本", "利润", "GDP",
            "货币", "投资", "汇率"],
}
_SUBJECT_CACHE = {}


def _infer_subject(text: str) -> str:
    """依据文本关键词计数推断学科；无法判定返回通用学科名。"""
    if text in _SUBJECT_CACHE:
        return _SUBJECT_CACHE[text]
    best, best_n = "相关专业", 0
    for subj, kws in SUBJECT_KEYWORDS.items():
        n = sum(1 for k in kws if k in text)
        if n > best_n:
            best, best_n = subj, n
    _SUBJECT_CACHE[text] = best
    return best


# --------------------------------------------------------------------------
# 三个 md 的解析（统一输出 {page: [条目文本]}）
# --------------------------------------------------------------------------
def _split_pages(content: str):
    """按 '## 第 N 页/节' 分块（页/节双兼容，可互读 word 产物），返回 {page: 块内容}。"""
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


def parse_texts(path: Path) -> dict:
    """texts.md：'## 第 N 页' 表格行 → {page: [文本条目]}。"""
    out = {}
    for page, lines in _split_pages(path.read_text(encoding="utf-8")).items():
        items = []
        for ln in lines:
            m = re.match(r"^\|\s*(TXT[\d\-]+)\s*\|\s*([^|]+?)\s*\|\s*(.*?)\s*\|$", ln)
            if m:
                items.append(f"[{m.group(2)}] {m.group(3)}")
        if items:
            out[page] = items
    return out


def parse_formulas(path: Path) -> dict:
    """formulas.md：'## 第 N 页' 内 $$...$$ 公式 → {page: [latex]}。"""
    out = {}
    for page, lines in _split_pages(path.read_text(encoding="utf-8")).items():
        items, buf, in_f = [], [], False
        for ln in lines:
            if ln.strip().startswith("$$"):
                if in_f:
                    items.append(" ".join(buf).strip())
                    buf, in_f = [], False
                else:
                    in_f = True
                    buf = []
            elif in_f:
                buf.append(ln.strip())
        if items:
            out[page] = items
    return out


def parse_captions(path: Path) -> dict:
    """captions.md：'### IMGxxxx — slide_NN_...' 块 → {page: [解读]}。
    页码从图片文件名 slide_NN 提取。"""
    out = {}
    cur_page, buf = None, []
    for ln in path.read_text(encoding="utf-8").splitlines():
        m = re.match(r"^###\s+IMG\d+\s*—\s*`?slide_(\d+)_", ln)
        if m:
            if cur_page is not None and buf:
                out.setdefault(cur_page, []).append(" ".join(x.strip()
                                                             for x in buf
                                                             if x.strip()))
            cur_page = int(m.group(1))
            buf = []
        elif cur_page is not None:
            buf.append(ln)
    if cur_page is not None and buf:
        out.setdefault(cur_page, []).append(" ".join(x.strip()
                                                     for x in buf
                                                     if x.strip()))
    return out


def auto_locate(dir_: Path):
    """目录内自动定位三份 md；返回 (texts, formulas, captions) 或 None。"""
    files = sorted(dir_.glob("*_texts.md"))
    if not files:
        return None
    stem = files[0].name[: -len("_texts.md")]
    texts = files[0]
    formulas = dir_ / f"{stem}_formulas.md"
    # captions 可能叫 <stem>_captions.md 或 images_captions.md
    cap = dir_ / f"{stem}_captions.md"
    if not cap.exists():
        cap = dir_ / "images_captions.md"
    return (texts, formulas, cap)


def aggregate(texts_path: Path, formulas_path: Path,
              captions_path: Path) -> tuple[dict, list]:
    """聚合三份 md 为 {page: {"texts","formulas","captions"}}。
    返回 (by_page, all_pages 排序列表)。"""
    t, f, c = (parse_texts(texts_path) if texts_path.exists() else {},
               parse_formulas(formulas_path) if formulas_path.exists() else {},
               parse_captions(captions_path) if captions_path.exists() else {})
    pages = sorted(set(t) | set(f) | set(c))
    by_page = {}
    for p in pages:
        by_page[p] = {"texts": t.get(p, []), "formulas": f.get(p, []),
                      "captions": c.get(p, [])}
    return by_page, pages


# --------------------------------------------------------------------------
# DeepSeek 调用
# --------------------------------------------------------------------------
SYSTEM_PROMPT_TMPL = (
    "你是高校教材建设专家、高级人工智能多模态知识库工程师、"
    "和{subject}专业资深教师，现给你输入三个md文档，一个是从PPT中提取的"
    "文本文档，一个是从PPT中提取的公式文档，一个是从PPT中提取的图片的"
    "解释文档，每个文档都标有页码。请根据这些文档，把整个PPT视为一部"
    "**完整的教材（一篇）**，系统性撰写每一页PPT相关知识点的教材描述文案。\n"
    "**篇章结构要求（v2.3）**：\n"
    "1. 根据全部页面内容**自主划分若干章**（章标题用 `# 第X章 章名`，"
    "如 `# 第1章 战略管理概述`），每章再**自主划分为若干节**"
    "（节标题用 `## 第X节 节名`，如 `## 第1节 战略的概念与本质`）；\n"
    "2. **一节可以包含一页或多页PPT**（同节的多页归入同一节）；\n"
    "3. **每一页**的内容（`## 第 N 页` 小节下）**第一行必须标注该页所属的"
    "「章名 · 节名」**，格式：`> 所属章节：第X章 章名 · 第X节 节名`；\n"
    "4. 每页文案**不少于500字**，用教材口吻描述，讲究知识描述的逻辑性和"
    "阅读的流畅性；跨批（分批生成）时章节命名保持前后一致、延续。\n"
    "将结果形成完整的md文档。"
)


def _infer_subject_llm(client, texts_path: Path, model: str,
                       retries: int = 2) -> str:
    """用 DeepSeek 判断学科：取 texts.md 前三页内容交给模型。

    失败时回退关键词表 _infer_subject。
    """
    try:
        pages = _split_pages(texts_path.read_text(encoding="utf-8"))
        first3 = sorted(pages)[:3]
        seg = []
        for p in first3:
            seg.append(f"## 第 {p} 页")
            seg.extend(pages[p])
        snippet = "\n".join(seg)[:6000]
        messages = [
            {"role": "system", "content":
             "你是高校教材建设专家。根据从 PPT 课程中提取的文本内容，"
             "判断这门课程属于哪个专业学科。只输出学科名称"
             "（如：电子信息工程、数学与应用数学、物理学、计算机科学与"
             "技术、自动控制、经济学等），不要输出任何其他内容。"},
            {"role": "user", "content":
             f"以下是该课程前三页的文本内容：\n\n{snippet}\n\n"
             "请判断这门课程的专业学科，只输出学科名称。"},
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
                # 清洗：去标点/多余说明，取第一个非空行
                for ln in name.splitlines():
                    ln = ln.strip(" -#*·、.。")
                    if ln:
                        return ln[:30]
                return name[:30]
            except Exception as e:
                last_err = e
                if attempt < retries:
                    time.sleep(2 * (attempt + 1))
        print(f"[警告] LLM 学科判断失败({last_err})，回退关键词表",
              file=sys.stderr)
    except Exception as e:
        print(f"[警告] 学科判断异常({e})，回退关键词表", file=sys.stderr)
    all_text = "\n".join(
        it for p, lines in _split_pages(
            texts_path.read_text(encoding="utf-8")).items() for it in lines)
    return _infer_subject(all_text)


def _build_full_prompt(by_page: dict, target: list, subject: str,
                       batch_ctx: str = "") -> str:
    """把目标页的三源内容按页拼接为整份 user 内容（含页码结构）。

    batch_ctx：批次上下文说明（如"本批为全篇第 X~Y 页，章节命名须与
    前批保持一致并延续"），v2.3 整体成篇时保证章节连贯。"""
    lines = [f"学科：{subject}", ""]
    if batch_ctx:
        lines.append(batch_ctx)
        lines.append("")
    for p in target:
        data = by_page[p]
        lines.append(f"## 第 {p} 页")
        if data["texts"]:
            lines.append("【文本】")
            for it in data["texts"]:
                lines.append(f"- {it}")
            lines.append("")
        if data["formulas"]:
            lines.append("【公式】")
            for f in data["formulas"]:
                lines.append(f"- $$ {f} $$")
            lines.append("")
        if data["captions"]:
            lines.append("【图片解读】")
            for c in data["captions"]:
                lines.append(f"- {c}")
            lines.append("")
    return "\n".join(lines)


def _gen_textbook(client, full_prompt: str, subject: str,
                  model: str, retries: int = 2) -> str:
    """一次调用模型生成整份教材文案 md。"""
    messages = [
        {"role": "system", "content":
         SYSTEM_PROMPT_TMPL.format(subject=subject)
         + " 请直接输出完整的教材文案 markdown 文档："
           "整篇分章（# 第X章 章名）→ 分节（## 第X节 节名，一节可含多页）"
           "→ 每页小节（## 第 N 页），每页内容首行标注 `> 所属章节："
           "第X章 章名 · 第X节 节名`，每页文案不少于 500 字；"
           "所有公式以 LaTeX 原样保留，图片解读内容自然融入对应页文案；"
           "不得遗漏任何一页、任何公式，不得输出文档以外的多余说明。"},
        {"role": "user", "content":
         full_prompt
         + "\n\n请根据以上三个 md 文档（文本/公式/图片解读，均标有页码），"
           "把整个 PPT 视为一部完整教材，自主划分为若干章、每章若干节"
           "（自主命名），一节可包含一页或多页；逐页撰写教材描述文案"
           "（每页不少于 500 字），每页内容首行标注所属章节"
           "（`> 所属章节：第X章 章名 · 第X节 节名`），教材口吻、逻辑连贯。"
           "直接输出完整的 markdown 文档。"},
    ]
    last_err = None
    for attempt in range(retries + 1):
        try:
            print(f"[DeepSeek] 教材文案输入（第 {attempt + 1} 次）："
                  f"{full_prompt[:300]}…", file=sys.stderr)
            resp = client.chat.completions.create(
                model=model, messages=messages, stream=False,
                reasoning_effort="high",
                extra_body={"thinking": {"type": "enabled"}},
            )
            out = (resp.choices[0].message.content or "").strip()
            print(f"[DeepSeek] 教材文案输出（前 300 字）：{out[:300]}…",
                  file=sys.stderr)
            return out
        except Exception as e:
            last_err = e
            if attempt < retries:
                time.sleep(2 * (attempt + 1))
    raise RuntimeError(f"API 调用失败: {last_err}")


# 自适应分批参数（可按模型能力调）
AVG_OUT_CHARS_PER_PAGE = 600        # 每页文案平均输出字符（估算输出上限用）
DEFAULT_MAX_INPUT_CHARS = 48000     # 单批输入字符上限（中文约 48K 字符）
DEFAULT_MAX_PAGES_PER_BATCH = 20    # 单批页数上限
DEFAULT_MAX_OUTPUT_CHARS = 12000    # 单批输出字符上限（估算）


def _make_batches(target: list, by_page: dict, subject: str,
                  max_input_chars: int, max_pages_per_batch: int,
                  max_output_chars: int) -> list:
    """贪心分批：每批满足 输入字符 ≤ max_input 且 页数 ≤ 每批页数
    且 估算输出（页数×600）≤ max_output。单页超过上限时单页自成一批。"""
    batches, cur, cur_chars = [], [], 0
    max_pages_by_out = max(1, max_output_chars // AVG_OUT_CHARS_PER_PAGE)
    cap = min(max_pages_per_batch, max_pages_by_out)
    for p in target:
        plen = len(_build_full_prompt(by_page, [p], subject))
        if cur and (cur_chars + plen > max_input_chars or len(cur) >= cap):
            batches.append(cur)
            cur, cur_chars = [], 0
        cur.append(p)
        cur_chars += plen
    if cur:
        batches.append(cur)
    return batches


def _tidy_direct(texts: list) -> str:
    """直出页轻度整理：按原文顺序拼接，去纯空白行、合并相邻重复行。
    尽量不改变原文意思、尽量不增加字数（v2.2 直出整理规则）。"""
    out = []
    for t in texts:
        line = t.strip()
        if not line:
            continue
        if out and out[-1] == line:
            continue  # 相邻重复行（PPT 版式常见）只保留一次
        out.append(line)
    return "\n".join(out)


def author_textbook(by_page: dict, pages: list, out_path: Path,
                    subject: str, model: str = DEFAULT_MODEL,
                    base_url: str = DEFAULT_BASE_URL,
                    api_key_env: str = DEFAULT_KEY_ENV,
                    pages_filter: list | None = None,
                    max_input_chars: int = DEFAULT_MAX_INPUT_CHARS,
                    max_pages_per_batch: int = DEFAULT_MAX_PAGES_PER_BATCH,
                    max_output_chars: int = DEFAULT_MAX_OUTPUT_CHARS,
                    no_expand_threshold: int = 500,
                    on_progress=None):
    """生成整份教材文案 md 并落盘；自适应分批，结果合并到一个文件。

    v2.2 撰写规则（**500 字为限**）：
    - 该页原始文本（texts 条目拼接、去空白）**≤ 500 字** → 模型扩写到
      **不少于 500 字**；
    - 该页原始文本 **> 500 字** → **直出整理**（_tidy_direct：轻度整理，
      尽量不改变原文意思、尽量不增加字数），不调用模型，省 Token；
    - 输入/输出体量超限 → 按页自动分批
      （每批受输入字符/页数/估算输出三重约束）；
    - 各页结果按页序统一合并写出（直出页与模型页顺序一致）。
    - pages_filter 给定（如 [1,4]）时，只处理目标页（小批量测试）。
    """
    api_key = os.environ.get(api_key_env, "")

    out_path.parent.mkdir(parents=True, exist_ok=True)   # 输出目录可不存在

    target = pages if pages_filter is None else \
        [p for p in pages if p in pages_filter]

    # 500 字为限：原文超阈值页 → 直出整理，不调用模型
    direct = {}
    if no_expand_threshold:
        for p in target:
            raw = "".join(by_page[p]["texts"]).replace(" ", "").replace("\n", "")
            if len(raw) > no_expand_threshold:
                direct[p] = _tidy_direct(by_page[p]["texts"])
    model_target = [p for p in target if p not in direct]
    if direct:
        print(f"[直出] {len(direct)} 页原文超过 {no_expand_threshold} 字，"
              f"直出整理（不改原意、不增字数）："
              f"页 {'、'.join(map(str, sorted(direct)))}",
              file=sys.stderr)

    results = {}          # page -> 页节内容（不含标题行）
    for p in sorted(direct):
        results[p] = direct[p].strip()

    if model_target:
        from openai import OpenAI
        client = OpenAI(api_key=api_key, base_url=base_url)
        total_input = len(_build_full_prompt(by_page, model_target, subject))
        batches = _make_batches(model_target, by_page, subject,
                                max_input_chars, max_pages_per_batch,
                                max_output_chars)
        n_batches = len(batches)
        if n_batches > 1:
            print(f"[分批] 总输入约 {total_input} 字符 > 单批上限 "
                  f"{max_input_chars}（或估算输出超限），自动拆为 "
                  f"{n_batches} 批："
                  + "、".join(f"{b[0]}~{b[-1]}({len(b)}页)"
                              for b in batches[:6])
                  + ("…" if n_batches > 6 else ""), file=sys.stderr)
        for bi, batch in enumerate(batches, start=1):
            if on_progress is not None:
                try:
                    on_progress(bi, n_batches, {"kind": "author"})
                except Exception:
                    pass
            # v2.3 整体成篇：批次上下文——告知这是全篇的第 X~Y 页，
            # 章节命名须与前批保持一致并延续（一节可跨批含多页）
            batch_ctx = (f"（本批为整篇教材的第 {batch[0]}~{batch[-1]} 页"
                         f"（共 {len(target)} 页中的一批）；"
                         f"章节划分须与之前批次保持一致并延续，"
                         f"不要重新命名已有章节。）")
            full_prompt = _build_full_prompt(by_page, batch, subject,
                                             batch_ctx=batch_ctx)
            print(f"[批次 {bi}/{n_batches}] 页 {batch[0]}~{batch[-1]}"
                  f"（{len(batch)} 页，输入约 {len(full_prompt)} 字符）",
                  file=sys.stderr)
            t0 = time.time()
            text = _gen_textbook(client, full_prompt, subject, model)
            dt = time.time() - t0
            n_out = len(re.findall(r"^##\s*第\s*\d+\s*页\s*$", text,
                                   re.MULTILINE))
            print(f"[批次 {bi}/{n_batches}] 完成，耗时 {dt:.0f}s，"
                  f"输出 {len(text)} 字符，识别到 {n_out} 页节",
                  file=sys.stderr)
            if n_out < len(batch):
                print(f"[警告] 本批输出页节数({n_out})少于输入页数"
                      f"({len(batch)})，可能被截断，可调小 --max-pages-per-batch",
                      file=sys.stderr)
            print(f"[输出预览] {text[:300]}", file=sys.stderr)
            print("…", file=sys.stderr)
            parsed = _split_pages(text)
            for pg, lines in parsed.items():
                if pg in batch:
                    results[pg] = "\n".join(x for x in lines if x.strip()).strip()
    else:
        n_batches = 0

    # 统一按页序写盘（直出页与模型页混合时顺序一致）
    header = (f"# {out_path.stem} 教材文案\n\n"
              f"> 由 `pptx-author` 生成 · 模型 `{model}` · 学科：{subject}\n"
              f"> 依据：文本/公式/图片解读三份 PPT 提取文档，共 {len(target)} 页\n"
              f"> 结构：整篇教材由模型**自主划分章（# 第X章 章名）/ 节"
              f"（## 第X节 节名）**，一节可含多页；每页内容首行标注所属章节\n"
              + (f"> 其中 {len(direct)} 页原文超 500 字直出整理"
                 f"（未参与自主分章）\n" if direct else "")
              + (f"> 分 {n_batches} 批生成（跨批章节命名延续）\n"
                 if n_batches > 1 else "")
              + "\n")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(header)
        for p in target:
            content = results.get(p)
            if content is None:
                continue
            f.write(f"## 第 {p} 页\n\n")
            if p in direct:
                f.write(f"> 直出整理（原文超过 {no_expand_threshold} 字，"
                        f"保留原意、尽量不增字数）。\n\n")
            f.write(f"{content}\n\n")
        f.write("---\n\n"
                f"共 {len(target)} 页（{len(direct)} 页直出 + {len(model_target)} 页模型，"
                f"{n_batches} 批），学科：{subject}。\n")
    total_chars = sum(len(v) for v in results.values())
    return {"pages_in": len(target), "pages_out": len(results),
            "direct_pages": len(direct), "model_pages": len(model_target),
            "batches": n_batches, "chars": total_chars,
            "subject": subject, "model": model, "md": str(out_path)}


def _check_env(args) -> bool:
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
    for name, p in (("texts", args.texts), ("formulas", args.formulas),
                    ("captions", args.captions)):
        if p is None:
            print(f"[环境] {name}.md: 未提供（目录自动查找失败）", file=sys.stderr)
            ok = False
        elif not p.exists():
            print(f"[环境] {name}.md 不存在：{p}", file=sys.stderr)
            ok = False
        else:
            print(f"[环境] {name}.md: {p.name}", file=sys.stderr)
    return ok


def _main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        prog="pptx-author",
        description="把文本/公式/图片解读三份 md 输入 DeepSeek，"
                    "按页撰写教材文案")
    ap.add_argument("dir", nargs="?", default=None,
                    help="产物目录（自动定位 *_texts.md 等；"
                         "也可用 --texts 等显式指定）")
    ap.add_argument("-o", "--output", default=None,
                    help="输出 md 路径（默认 <目录>/<名>_textbook.md）")
    ap.add_argument("--subject", default=None,
                    help="学科（如 模拟电子技术；默认按文本关键词自动推断）")
    ap.add_argument("--texts", type=Path, default=None, help="文本 md")
    ap.add_argument("--formulas", type=Path, default=None, help="公式 md")
    ap.add_argument("--captions", type=Path, default=None, help="图片解读 md")
    ap.add_argument("--model", default=DEFAULT_MODEL,
                    help=f"模型名（默认 {DEFAULT_MODEL}）")
    ap.add_argument("--base-url", default=DEFAULT_BASE_URL,
                    help=f"OpenAI 兼容端点（默认 {DEFAULT_BASE_URL}）")
    ap.add_argument("--api-key-env", default=DEFAULT_KEY_ENV,
                    help=f"API Key 环境变量名（默认 {DEFAULT_KEY_ENV}）")
    ap.add_argument("--pages", default=None,
                    help="只处理指定页，如 '1-5,8'（默认全部；测试用）")
    ap.add_argument("--max-input-chars", type=int,
                    default=DEFAULT_MAX_INPUT_CHARS,
                    help=f"单批输入字符上限（默认 {DEFAULT_MAX_INPUT_CHARS}；"
                         f"超限自动分批）")
    ap.add_argument("--max-pages-per-batch", type=int,
                    default=DEFAULT_MAX_PAGES_PER_BATCH,
                    help=f"单批页数上限（默认 {DEFAULT_MAX_PAGES_PER_BATCH}；"
                         f"控制输出体量防截断）")
    ap.add_argument("--max-output-chars", type=int,
                    default=DEFAULT_MAX_OUTPUT_CHARS,
                    help=f"单批估算输出字符上限（默认 {DEFAULT_MAX_OUTPUT_CHARS}；"
                         f"按每页600字折算页数上限）")
    ap.add_argument("--no-expand-threshold", type=int, default=500,
                    help="原文直出阈值：某页原始文本（去空白）超过该字数时"
                         "**直出整理**（轻度整理、不改原意、尽量不增字数），"
                         "不再扩写；否则由模型扩写到不少于 500 字"
                         "（默认 500：不足 500 扩写、超 500 直出；"
                         "0 关闭直出全部扩写）")
    ap.add_argument("--json", action="store_true",
                    help="结构化统计输出到 stdout")
    ap.add_argument("--version", action="version", version=VERSION)
    args = ap.parse_args(argv)

    # 定位三份 md
    if args.dir:
        loc = auto_locate(Path(args.dir))
        if loc:
            args.texts = args.texts or loc[0]
            args.formulas = args.formulas or loc[1]
            args.captions = args.captions or loc[2]
    if not args.texts:
        print("[错误] 未找到 *_texts.md（请传目录或 --texts 指定）",
              file=sys.stderr)
        return EXIT_USAGE
    if not args.formulas:
        print("[错误] 未找到 *_formulas.md（请传目录或 --formulas 指定）",
              file=sys.stderr)
        return EXIT_USAGE
    if not args.captions:
        print("[错误] 未找到 *_captions.md / images_captions.md"
              "（请传目录或 --captions 指定）", file=sys.stderr)
        return EXIT_USAGE

    if not _check_env(args):
        return EXIT_USAGE

    try:
        by_page, pages = aggregate(args.texts, args.formulas, args.captions)

        # 学科：默认用 DeepSeek 判断（取 texts.md 前三页），失败回退关键词表
        subject = args.subject
        if not subject:
            from openai import OpenAI
            client0 = OpenAI(
                api_key=os.environ.get(args.api_key_env, ""),
                base_url=args.base_url)
            subject = _infer_subject_llm(client0, args.texts, args.model)
        print(f"[学科] {subject}（--subject 可覆盖）", file=sys.stderr)

        pages_filter = None
        if args.pages:
            pages_filter = set()
            for seg in args.pages.split(","):
                seg = seg.strip()
                if "-" in seg:
                    a, b = seg.split("-")
                    pages_filter |= set(range(int(a), int(b) + 1))
                else:
                    pages_filter.add(int(seg))

        out_path = Path(args.output) if args.output else \
            Path(args.dir or ".") / f"{args.texts.stem[: -len('_texts')]}_textbook.md"

        cb = make_progress("教材文案")
        if args.json:
            with quiet_stdout():
                result = author_textbook(by_page, pages, out_path,
                                         subject, model=args.model,
                                         base_url=args.base_url,
                                         api_key_env=args.api_key_env,
                                         pages_filter=pages_filter,
                                         max_input_chars=args.max_input_chars,
                                         max_pages_per_batch=args.max_pages_per_batch,
                                         max_output_chars=args.max_output_chars,
                                         no_expand_threshold=args.no_expand_threshold,
                                         on_progress=cb)
            print_json(result)
        else:
            result = author_textbook(by_page, pages, out_path,
                                     subject, model=args.model,
                                     base_url=args.base_url,
                                     api_key_env=args.api_key_env,
                                     pages_filter=pages_filter,
                                     max_input_chars=args.max_input_chars,
                                     max_pages_per_batch=args.max_pages_per_batch,
                                     max_output_chars=args.max_output_chars,
                                     no_expand_threshold=args.no_expand_threshold,
                                     on_progress=cb)
            print(f"[OK] 教材文案已写出：{out_path}")
            print(f"     输入 {result['pages_in']} 页，分 {result['batches']} 批，"
                  f"输出 {result['pages_out']} 页节，{result['chars']} 字符"
                  + (f"（直出 {result['direct_pages']} 页）"
                     if result.get("direct_pages") else ""))
        return EXIT_OK
    except Exception as e:
        print(f"[错误] 教材文案生成失败：{e}", file=sys.stderr)
        return EXIT_ERR


def main() -> int:  # console
    banner("pptx-author")
    rc = _main()
    banner_end("pptx-author")
    return rc


if __name__ == "__main__":
    sys.exit(_main())
