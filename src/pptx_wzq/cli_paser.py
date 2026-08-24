"""cli_paser.py — pptx-wzq 整体编排指令。

把一份 PPT 依次跑完：文本提取 → 公式提取 → 可视逻辑块解析（合并图片 AI 解读）→
相关性过滤 → 教材文案生成 → 可视逻辑块 JSON 组装，并整理产物。

用法：
    pptx-paser XXX.pptx -o 生成结果目录
    pptx-paser XXX.pptx -o 结果目录 [--author-pages "1,4"] [--skip blocks]
      --skip 可逗号分隔跳过步骤：text,formula,blocks,related,author,blocks_json

流程（含交互确认）：
    1) 环境检查：逐项组件检查并显示结果；
    2) KEY 检查：两个 API Key 缺失时打印注册引导（网页/用途/资费），
       交互输入 Key 并用 setx 注册，提示重启后重跑；
    3) 执行计划：分步说明每步做什么 + Token 消耗估算 → 询问用户是否继续；
    4) 依次执行六步（text/formula 提前为 blocks 的 Semantic Captioning
       提供上下文；blocks 合并原 caption 职责）；
    5) 产物归位 + 执行结果汇总 + 结果文档使用说明。

    产物组织：
        生成结果目录/
          ├─ images/                    ← 可视逻辑块渲染图（统一图片集）
          ├─ sources/                   ← 可视逻辑区的矢量/Visio 资源（vsdx/svg/wmf/emf）
          ├─ <名>_visualBlock_text_binding.json ← 可视逻辑区↔文本 图文关联
          ├─ <名>_captions.md           ← 可视逻辑块级 AI 解读（Semantic Captioning）
          ├─ <名>_textbook.md           ← 教材文案
          ├─ <名>_visual_blocks.json    ← 可视逻辑块全栈解析（替换原 binding.json）
          过程文件/                     ← 其余全部（by_page/manifest/
                                         texts.md/formulas.md/atomic_objects/…）

执行前环境检查：
    1) 各子指令依赖：openai / Pillow / ultralytics+yolov5su.pt /
       olefile / ElementTree / LibreOffice(可选)；
    2) 两个模型的 API Key 环境变量：
       DEEPSEEK_API_KEY（DeepSeek：学科判断+教材文案+关系判定）
       DASHSCOPE_API_KEY（阿里云百炼：可视逻辑块 Semantic Captioning）

退出码：0 成功（含用户取消）/ 1 处理异常 / 2 参数或环境错误。

作者：吴振谦 · wuzhenqian@nbu.edu.cn · QQ：38328063"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path

from pptx_wzq.cli_common import banner, banner_end
from pptx_wzq import __version__ as _PKG_VERSION

VERSION = f"pptx-wzq {_PKG_VERSION} (可视逻辑块全栈解析 + Semantic Captioning 合并)"
PROC_DIR = "过程文件"

# 步骤顺序（v2.0 流程①-⑧：blocks 先行解构 → text/formula 提供上下文 →
# caption 独立解读（sources/ 顺序，DeepSeek-XML / qwen 兜底）→ related →
# author 扩写 → blocks_json 组装输出）
STEPS = ["blocks", "text", "formula", "caption", "related", "author",
         "blocks_json"]


# --------------------------------------------------------------------------
# 断点续传：state.json（机器可读，原子写）+ pipeline.log（人类可读，追加）
# --------------------------------------------------------------------------
def _load_state(out: Path):
    """读 out/state.json；不存在/损坏返回 None。"""
    p = out / "state.json"
    if p.is_file():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return None
    return None


def _save_state(out: Path, state) -> None:
    """state.json 原子写（临时文件 + rename）。"""
    p = out / "state.json"
    tmp = p.with_name(p.name + ".tmp")
    tmp.write_text(json.dumps(state, ensure_ascii=False, indent=1),
                   encoding="utf-8")
    tmp.replace(p)


def _log(out: Path, msg: str) -> None:
    """pipeline.log 追加一行（时间戳 + 消息）。"""
    try:
        with open(out / "pipeline.log", "a", encoding="utf-8") as f:
            f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n")
    except OSError:
        pass


def _init_state(pptx: Path, out: Path) -> dict:
    """新建状态机：全部步骤 pending。含 doc_md5 与 tool_version（对齐
    word .wzq_checkpoint.json §3.7：换源/工具版本变化可提示）。"""
    import hashlib
    try:
        md5 = hashlib.md5(pptx.read_bytes()).hexdigest()
    except OSError:
        md5 = ""
    state = {"pptx": str(pptx), "stem": pptx.stem,
             "doc_md5": md5,
             "tool_version": VERSION,
             "started_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
             "steps": {s: {"status": "pending", "note": ""} for s in STEPS}}
    _save_state(out, state)
    return state


def _set_step(out: Path, state: dict, step: str, status: str,
              note: str = "") -> None:
    state["steps"][step] = {"status": status, "note": note,
                            "updated_at": time.strftime(
                                "%Y-%m-%dT%H:%M:%S")}
    _save_state(out, state)


def _artifact_ok(out: Path, stem: str, step: str) -> bool:
    """该步骤的产物是否已存在（out 根 / 过程文件 / _proc 三处探测）。"""
    proc = out / PROC_DIR
    work = out / "_proc"
    cands = []
    if step == "formula":
        cands = [proc / "formula", work / "formula"]
    elif step == "text":
        cands = [proc / "text", work / "text"]
    elif step == "blocks":
        cands = [out / f"{stem}_visual_blocks.json",
                 proc / "blocks" / f"{stem}_visual_blocks.json",
                 work / "blocks" / f"{stem}_visual_blocks.json"]
    elif step == "caption":
        cands = [out / f"{stem}_captions.md",
                 proc / "blocks" / "captions.md",
                 work / "blocks" / "captions.md"]
    elif step == "related":
        cands = [out / f"{stem}_related_filter.json",
                 proc / "cap" / f"{stem}_related_filter.json",
                 work / "cap" / f"{stem}_related_filter.json"]
    elif step == "author":
        cands = [out / f"{stem}_textbook.md",
                 proc / f"{stem}_textbook.md",
                 work / f"{stem}_textbook.md"]
    elif step == "blocks_json":
        cands = [out / f"{stem}_visual_blocks.json"]
    return any(c.is_dir() if c.suffix == "" else c.is_file() for c in cands)


def _build_plan(out: Path, stem: str, state: dict, skip: set) -> dict:
    """每步 → ("skip"|"run"|"resume", 说明)。规则：
    - done 且产物存在 → skip；done 但产物缺失 → run（重跑）；
    - partial → resume；failed/无记录 → run。"""
    plan = {}
    for s in STEPS:
        if s in skip:
            plan[s] = ("skip", "用户 --skip 指定")
            continue
        st = (state or {}).get("steps", {}).get(s, {})
        status = st.get("status")
        if status == "done":
            if _artifact_ok(out, stem, s):
                plan[s] = ("skip", "已完成（state=done 且产物存在）")
            else:
                plan[s] = ("run", "状态为完成但产物缺失，重跑")
        elif status == "partial":
            plan[s] = ("resume", "上次中断，断点续跑")
        else:
            plan[s] = ("run", st.get("note") or "")
    return plan


def _reset_out(out: Path, stem: str) -> None:
    """--reset：清空本管线生成的全部旧产物与状态（保留目录本身）。"""
    for p in (out / "images", out / "sources", out / PROC_DIR, out / "_proc",
              out / "state.json", out / "pipeline.log",
              out / f"{stem}_captions.md", out / f"{stem}_textbook.md",
              out / f"{stem}_visual_blocks.json",
              out / f"{stem}_related_filter.json",
              out / "images_meta.json"):
        if p.is_dir():
            shutil.rmtree(p, ignore_errors=True)
        elif p.is_file():
            try:
                p.unlink()
            except OSError:
                pass
    _log(out, "--reset：已清空旧产物与状态")


def _print_plan(plan: dict, est: dict, skip: set) -> None:
    """打印执行/续跑计划（含每步动作与原因）。"""
    print("\n========== 执行计划（含断点续跑判断） ==========", file=sys.stderr)
    for s in STEPS:
        action, note = plan[s]
        mark = {"skip": "跳过", "run": "执行", "resume": "续跑"}[action]
        print(f"  {s:8s} [{mark}] {note}", file=sys.stderr)
    print(f"  规模：约 {est['slides']} 页 / 媒体 {est['media']} 个 / "
          f"公式对象 {est['embeds']} 个", file=sys.stderr)
    print("", file=sys.stderr)

# KEY 注册引导（网页 / 用途 / 资费）
KEY_GUIDE = {
    "DEEPSEEK_API_KEY": (
        "① DeepSeek（模型 deepseek-v4-flash，文本）\n"
        "   用途：学科自动判断 + 教材文案生成\n"
        "   注册：打开 https://platform.deepseek.com → 注册/登录 → "
        "左侧「API Keys」→ 创建 API Key → 复制保存\n"
        "   资费：按 Token 计费（输入+输出），单价以官网 "
        "https://api-docs.deepseek.com 公布为准"),
    "DASHSCOPE_API_KEY": (
        "② 阿里云百炼 DashScope（模型 qwen3.7-plus，视觉）\n"
        "   用途：图片 AI 解读（多模态）\n"
        "   注册：打开 https://bailian.console.aliyun.com → 登录阿里云 → "
        "开通「百炼」服务 → 右上角「API-KEY」→ 创建 API Key → 复制保存\n"
        "   资费：按 Token 计费，qwen 系列新用户一般有免费额度，"
        "以官网 https://help.aliyun.com/zh/model-studio 公布为准"),
}


# pip 依赖表：(import 模块名, pip 包名, 用途, 是否必需)
PIP_DEPS = [
    ("PIL", "Pillow", "图片提取/过滤（像素判据）", True),
    ("openai", "openai", "图片AI解读/教材文案/公式识别", True),
    ("win32com", "pywin32", "PowerPoint 渲染 WMF（公式识别/矢量转PNG）", True),
    ("olefile", "olefile", "公式 OLE 解包（路径2）", True),
    ("ultralytics", "ultralytics", "YOLO 图片过滤", True),
    ("omml2latex", "omml2latex", "OMML 原生公式→LaTeX（路径1）", True),
    ("pix2tex", "pix2tex[latex]", "公式 OCR 兜底（路径3/公式识别）", False),
]


def _ensure_deps(no_install: bool = False) -> list:
    """检查并自动安装 pip 依赖（用当前 Python）。返回 [(desc, ok, note)]。"""
    import importlib.util as u
    results = []
    for mod, pkg, desc, required in PIP_DEPS:
        if u.find_spec(mod):
            results.append((desc, True, "已安装"))
            continue
        if no_install:
            results.append((desc, False,
                            f"缺失（{pkg}；--no-install 已跳过自动安装）"))
            continue
        print(f"[依赖] 正在安装 {pkg}（{desc}）…", file=sys.stderr)
        r = subprocess.run(
            [sys.executable, "-m", "pip", "install", "--quiet", pkg],
            capture_output=True, text=True)
        if r.returncode == 0 and u.find_spec(mod):
            results.append((desc, True, "已自动安装"))
        else:
            note = (f"安装失败：{r.stderr.strip()[-120:]}" if r.stderr
                    else "安装失败")
            note += "（可选，跳过）" if not required else \
                f"（必需，请手动 pip install {pkg}）"
            results.append((desc, False, note))
    return results


def _check_env_all(no_install: bool = False) -> bool:
    """环境检查：pip 依赖自动安装 + 权重/LO 探测。返回是否全部必需项就绪。"""
    print("[依赖] 检查并自动安装组件库：", file=sys.stderr)
    deps = _ensure_deps(no_install=no_install)
    all_ok = True
    for desc, ok, note in deps:
        print(f"[依赖] {desc}: {'OK' if ok else note}", file=sys.stderr)
        if not ok and "必需" in note:
            all_ok = False

    # 非 pip 依赖：YOLO 权重（缺失则 ultralytics 首次调用自动下载）
    yolo_pt = any(Path(p).is_file()
                  for p in ("yolov5su.pt", "yolov5s.pt", "yolov5n.pt"))
    print(f"[依赖] YOLO 权重 yolov5su.pt: "
          f"{'OK' if yolo_pt else '缺失（首次 YOLO 调用将自动下载，约 17MB）'}",
          file=sys.stderr)
    # LibreOffice（可选：公式路径3 渲染/降级提示）
    lo = any(Path(p).is_file() for p in (
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"))
    print(f"[依赖] LibreOffice（公式路径3 渲染）: "
          f"{'OK' if lo else '未安装（可选，路径3 降级）'}",
          file=sys.stderr)
    return all_ok


def _ensure_keys() -> bool:
    """检查两个 KEY；缺失则打印注册引导 + 交互输入 + setx 注册。

    返回 False 表示需要重启后重跑（本进程读不到刚 setx 的变量）。
    """
    env_keys = {"DEEPSEEK_API_KEY": "DeepSeek（教材文案/学科判断）",
                "DASHSCOPE_API_KEY": "阿里云百炼 DashScope（图片AI解读）"}
    missing = {k: v for k, v in env_keys.items()
               if not os.environ.get(k)}
    if not missing:
        print("[KEY] 两个 API Key 均已设置 ✓", file=sys.stderr)
        return True
    print("\n[KEY] 缺少以下 API Key，将引导你注册并自动写入环境变量：\n",
          file=sys.stderr)
    for k in missing:
        print(KEY_GUIDE[k], file=sys.stderr)
        print("", file=sys.stderr)
    print("[KEY] 请把 Key 粘贴到下面（粘贴后回车；不显示为正常现象）：",
          file=sys.stderr)
    for k in list(missing):
        try:
            v = input(f"请输入 {k} 的 Key（直接回车跳过）：").strip()
        except EOFError:
            v = ""
        if v:
            r = subprocess.run(["setx", k, v],
                               capture_output=True, text=True)
            if r.returncode == 0:
                print(f"[KEY] {k} 已写入系统环境变量（setx）", file=sys.stderr)
                missing.pop(k)
            else:
                print(f"[KEY] {k} 注册失败：{r.stderr.strip()}",
                      file=sys.stderr)
    if missing:
        print(f"\n[KEY] 仍有缺失：{', '.join(missing)}", file=sys.stderr)
        print("[KEY] 请按上述引导注册，然后重新运行本指令。",
              file=sys.stderr)
        return False
    print("\n[KEY] 两个 Key 已全部注册完成。", file=sys.stderr)
    print("[重要] 环境变量要**重启系统（或至少重开命令窗口）**后才生效；", file=sys.stderr)
    print("[重要] 请重启后重新运行：pptx-paser <ppt> -o <结果目录>",
          file=sys.stderr)
    return False


def _dir_summary(path: Path) -> str:
    """目录统计：文件数 + 总大小。"""
    if not path.is_dir():
        return "（未生成）"
    files = [f for f in path.rglob("*") if f.is_file()]
    total = sum(f.stat().st_size for f in files)
    unit = "KB" if total < 1048576 else "MB"
    size = total / 1024 if unit == "KB" else total / 1048576
    return f"{len(files)} 个文件（{size:.1f} {unit}）"


def _run_step(name: str, script: str, args_list: list, cwd: Path,
              desc: str = "", resource: str = "", outs: list | None = None,
              idx: int = 0, total: int = 6) -> None:
    """用当前 Python 运行子指令，并在执行前后打印：

    工作内容（desc）/ 消耗资源（resource）/ 执行耗时 / 产生结果（outs 产物统计）。
    """
    print(f"\n========== 步骤 {idx}/{total}：{name} ==========", file=sys.stderr)
    if desc:
        print(f"[工作内容] {desc}", file=sys.stderr)
    if resource:
        print(f"[消耗资源] {resource}", file=sys.stderr)
    t0 = time.monotonic()
    if script.endswith(".py"):
        cmd = [sys.executable, script] + [str(a) for a in args_list]
    else:
        cmd = [sys.executable, "-m", f"pptx_wzq.{script}"] + \
            [str(a) for a in args_list]
    r = subprocess.run(cmd, cwd=str(cwd))
    cost = time.monotonic() - t0
    if r.returncode != 0:
        raise RuntimeError(f"步骤「{name}」失败（exit={r.returncode}），"
                           f"命令：{' '.join(map(str, cmd))}")
    print(f"========== 步骤完成：{name} ==========", file=sys.stderr)
    print(f"[结果] 耗时 {cost:.1f} 秒", file=sys.stderr)
    for o in (outs or []):
        print(f"       {o} → {_dir_summary(Path(o))}", file=sys.stderr)
    print("", file=sys.stderr)


def _missing_author_pages(work: Path, stem: str) -> str | None:
    """author 续跑：texts.md 全部页 - textbook.md 已有页 → 缺失页串。"""
    texts = work / "text" / f"{stem}_texts.md"
    tb = work / f"{stem}_textbook.md"
    if not tb.is_file() or not texts.is_file():
        return None

    def pages_of(path: Path) -> set:
        out = set()
        for ln in path.read_text(encoding="utf-8").splitlines():
            m = re.match(r"^##\s*第\s*(\d+)\s*(?:页|节)\s*$", ln.strip())
            if m:
                out.add(int(m.group(1)))
        return out
    todo = sorted(pages_of(texts) - pages_of(tb))
    return ",".join(map(str, todo)) if todo else None


def _exec_step(step: str, args, work: Path, stem: str,
               action: str) -> None:
    """执行单个步骤（text/formula/blocks/related/author/blocks_json）。
    action: run | resume（resume 透传给子命令的续跑参数）。"""
    pptx = Path(args.pptx)
    total = len(STEPS)
    idx = STEPS.index(step) + 1
    if step == "text":
        _run_step(
            "文本提取（文本 ID + 坐标）", "cli_text",
            [pptx, "-o", work / "text"], ".",
            desc="逐页提取文本框 / 表格 / 标题占位符文本（排除页眉页脚），"
                 "每条分配 text_id（TXT###-##）并记录幻灯片坐标 x/y/w/h → "
                 "<名>_texts.md + <名>_text_entries.json",
            resource="本地计算（0 Token）· 仅依赖 Python 标准库 XML 解析",
            outs=[work / "text"], idx=idx, total=total)
    elif step == "formula":
        _run_step(
            "公式提取", "cli_formula",
            [pptx, "-o", work / "formula"], ".",
            desc="遍历全部公式对象，三路径级联转为 LaTeX：路径1 原生 OMML "
                 "（omml2latex）→ 路径2 Equation.3 公式编辑器（MTEF 解码）→ "
                 "路径3 渲染 + 数学 OCR（缺失时降级为占位）→ 按页输出 "
                 "<名>_formulas.md",
            resource="本地计算（0 Token）· 依赖 omml2latex / olefile / "
                     "mtef_decoder；OCR 可选",
            outs=[work / "formula"], idx=idx, total=total)
    elif step == "blocks":
        # v2.0 流程①：只做确定性拆块 + 解构（XML 段/矢量源/像素图 → sources/，
        # 渲染图 → images/）；caption 解读独立为步骤④（caption）
        blk = [work / "blocks", "--pptx", pptx,
               "--texts", work / "text" / f"{stem}_texts.md",
               "--formulas", work / "formula" / f"{stem}_formulas.md",
               "-o", work / "blocks" / f"{stem}_visual_blocks.json",
               "--captions", work / "blocks" / "captions.md",
               "--no-vlm"]   # 结构阶段不跑 VLM；解读走 caption 步骤
        if action == "resume":
            blk.append("--resume")
        _run_step(
            "可视逻辑块提取 + 解构（v2.0 确定性拆块）",
            "cli_blocks", blk, ".",
            desc="读原子对象 → 单阶段确定性拆块（grpSp→group / visio / "
                 "像素图≥20% / SVG-WMF 矢量；无自由聚类）→ 块渲染 PNG（images/）"
                 "→ grpSp XML 段每组合一个文件 + 资源图片（rldimg）→ sources/ "
                 "→ visual_blocks.json + captions.md（规则模板）",
            resource="本地 + 渲染（PowerPoint COM，仅 PowerPoint）；解读在 caption 步骤",
            outs=[work / "blocks"], idx=idx, total=total)
    elif step == "caption":
        # v2.0 流程④：按 sources/ 文件顺序解读——.xml → DeepSeek 读 XML
        # （公式转 LaTeX）；.png 像素图原图 → qwen VLM 兜底；images/ 不解读
        cap = [work / "blocks", "--pptx", pptx,
               "--texts", work / "text" / f"{stem}_texts.md",
               "-o", work / "blocks" / f"{stem}_visual_blocks.json",
               "--captions", work / "blocks" / "captions.md",
               "--caption-sources",
               "--semantic-model", "deepseek-v4-flash"]
        if action == "resume":
            cap.append("--resume")
        _run_step(
            "图块 AI 解读（sources/ 顺序路由）",
            "cli_blocks", cap, ".",
            desc="遍历 sources/ 文件（按文件名排序）：.xml（grpSp/Visio/"
                 "SVG-WMF 段）→ DeepSeek-V4-Flash 读 XML（超长本地压缩，"
                 "组内公式转 LaTeX）；.png 像素图原图 → qwen3.7-plus 读图兜底；"
                 "DeepSeek 空响应 → qwen 读渲染图/规则模板二级兜底；"
                 "images/ 仅供人阅览不解读",
            resource="消耗 Token（DeepSeek deepseek-v4-flash + qwen3.7-plus 兜底）",
            outs=[work / "blocks"], idx=idx, total=total)
    elif step == "related":
        _run_step(
            "可视逻辑块相关性过滤（剔除 logo/作者/单位等无关块）", "cli_related",
            [work / "blocks", "-o", work / "blocks" / "captions.md",
             "--texts", work / "text" / f"{stem}_texts.md",
             "--images-dir", work / "blocks" / "images"], ".",
            desc="把每个块的 caption 与该页正文交给 DeepSeek 判断相关性，"
                 "无关块（品牌 logo / 作者信息 / 单位名称 / 项目类别 / "
                 "每页重复装饰 / 二维码等）从 images/、captions.md 中删除 → "
                 "<名>_related_filter.json 审计",
            resource="消耗 Token（文本模型 deepseek-v4-flash，DEEPSEEK_API_KEY）",
            outs=[work / "blocks"], idx=idx, total=total)
    elif step == "author":
        au = ["--texts", work / "text" / f"{stem}_texts.md",
              "--formulas", work / "formula" / f"{stem}_formulas.md",
              "--captions", work / "blocks" / "captions.md",
              "-o", work / f"{stem}_textbook.md"]
        if action == "resume":
            miss = _missing_author_pages(work, stem)
            if miss:
                au += ["--pages", miss]
                print(f"[续跑] author：只生成缺失页 {miss}", file=sys.stderr)
            else:
                au += ["--pages", "0"]   # 无缺失页：空跑保持已完成
        _run_step(
            "教材文案生成（每页扩写至不少于 500 字）", "cli_author", au, ".",
            desc="学科自动推断 → 文本/公式/可视逻辑块解读三份文档输入 DeepSeek "
                 "deepseek-v4-flash **逐页扩写教材文案，每页不少于 500 字**"
                 "（v2.1 起默认关闭原文直出）→ <名>_textbook.md；超长自动分批",
            resource="消耗 Token（文本模型 deepseek-v4-flash，DEEPSEEK_API_KEY）",
            outs=[work], idx=idx, total=total)
    elif step == "blocks_json":
        bd = [work / "blocks",
              "--atomic-objects", work / "blocks" / "atomic_objects.json",
              "--texts", work / "text" / f"{stem}_texts.md",
              "--formulas", work / "formula" / f"{stem}_formulas.md",
              "--pptx", pptx,
              "--prev-blocks", work / "blocks" / f"{stem}_visual_blocks.json",
              "-o", work.parent / f"{stem}_visual_blocks.json",
              "--captions", work.parent / f"{stem}_captions.md",
              # DeepSeek 语义增强：本步骤用文本模型生成每个可视逻辑块的
              # semantic_description（expression_goal/role/features/…）
              "--semantic-model", "deepseek-v4-flash"]
        if action == "resume":
            bd.append("--resume")
        _run_step(
            "可视逻辑块 JSON 组装（全栈解析 + DeepSeek 语义增强 + 跨模态关系）",
            "cli_blocks", bd, ".",
            desc="按 pptx_multimodal_slide_v2.0 schema 组装 <名>_visual_blocks.json"
                 "（slide_info / textual_content / visual_blocks[] / "
                 "cross_modal_relations[] / summary），调用 DeepSeek "
                 "deepseek-v4-flash 生成每个块的 semantic_description "
                 "（表达目标/作用/抽象特征/图文描述/教学用途），并把块渲染图"
                 "写入 images/、captions.md 写入结果目录",
            resource="消耗 Token（语义增强+跨模态关系：文本模型 "
                     "deepseek-v4-flash，DEEPSEEK_API_KEY）",
            outs=[work / "blocks"], idx=idx, total=total)


def _organize(out: Path, work: Path, stem: str) -> None:
    """产物归位：结果目录 = images/（块渲染图）+ sources/（块矢量资源）+
    <名>_visualBlock_text_binding.json（块↔文本关系）+ <名>_captions.md +
    <名>_textbook.md + <名>_visual_blocks.json（由 blocks_json 步骤直接写）；
    其余全部移到 过程文件/。

    旧的「单张原子图 + images_meta.json」交付物已取消：原子对象仅作为可视逻辑
    块聚类的几何地基，原子图不再单独交付；矢量/Visio 源改由「块」持有
    （sources/ 按 block_id 归档）。
    """
    proc = out / PROC_DIR
    # v2.0 生命周期：成功即清理，不留过程文件目录（不创建 过程文件/）；
    # 仅清理历史残留的 PROC_DIR 空目录（若存在）

    # 1) 可视逻辑块渲染图 → images/（统一图片集，与 captions.md / JSON 对应）
    blk_images = work / "blocks" / "images"
    if blk_images.is_dir():
        dst = out / "images"
        dst.mkdir(parents=True, exist_ok=True)
        n = 0
        for f in sorted(blk_images.iterdir()):
            if f.is_file() and not (dst / f.name).exists():
                try:
                    shutil.copy2(str(f), str(dst / f.name))
                    n += 1
                except OSError:
                    pass
        print(f"[归位] 可视逻辑块渲染图 {n} 张 → {dst}", file=sys.stderr)

    # 2) 可视逻辑块矢量/Visio/XML/资源 → sources/（按 block_id 命名）
    #    **递归**：rldimg/ 等子目录（组内资源图片）必须一并归位（规则 10）
    blk_sources = work / "blocks" / "sources"
    if blk_sources.is_dir():
        dst = out / "sources"
        dst.mkdir(parents=True, exist_ok=True)
        n = 0
        for f in sorted(blk_sources.iterdir()):
            if f.is_dir():
                # 子目录（rldimg/ 等）整目录归位，目标已存在则跳过
                sub = dst / f.name
                if not sub.exists():
                    try:
                        shutil.copytree(str(f), str(sub))
                        n += sum(1 for _ in f.rglob("*") if _.is_file())
                    except OSError:
                        pass
            elif f.is_file() and not (dst / f.name).exists():
                try:
                    shutil.copy2(str(f), str(dst / f.name))
                    n += 1
                except OSError:
                    pass
        if n:
            print(f"[归位] 块资源 {n} 个 → {dst}", file=sys.stderr)
        else:
            try:
                dst.rmdir()
            except OSError:
                pass

    # 3) 块↔文本图文关联 → <名>_visualBlock_text_binding.json
    bnd = work / "blocks" / f"{stem}_visualBlock_text_binding.json"
    if bnd.is_file():
        shutil.move(str(bnd),
                    str(out / f"{stem}_visualBlock_text_binding.json"))
        print(f"[归位] 块↔文本关联 → "
              f"{out / (stem + '_visualBlock_text_binding.json')}",
              file=sys.stderr)

    # 4) captions.md / textbook.md
    cap = work / "blocks" / "captions.md"
    cap_dst = out / f"{stem}_captions.md"
    if cap.is_file() and not cap_dst.exists():
        shutil.move(str(cap), str(cap_dst))
        print(f"[归位] 可视逻辑块解读文档 → "
              f"{out / (stem + '_captions.md')}", file=sys.stderr)
    elif cap.is_file():
        # 最终版已由 blocks_json 步骤写到结果目录，丢弃中间版
        try:
            cap.unlink()
        except OSError:
            pass
    tb = work / f"{stem}_textbook.md"
    if tb.is_file():
        shutil.move(str(tb), str(out / f"{stem}_textbook.md"))
        print(f"[归位] 教材文案 → {out / (stem + '_textbook.md')}",
              file=sys.stderr)
    # visual_blocks.json 由 blocks_json 步骤直接写到 out，这里只校验

    # 5) 文本/公式清单 → 结果目录（交付物）
    for _src, _names in ((work / "text", (f"{stem}_texts.md",
                                          f"{stem}_text_entries.json")),
                         (work / "formula", (f"{stem}_formulas.md",))):
        if _src.is_dir():
            for _n in _names:
                _f = _src / _n
                if _f.is_file() and not (out / _n).exists():
                    try:
                        shutil.move(str(_f), str(out / _n))
                        print(f"[归位] {_n} → {out}", file=sys.stderr)
                    except OSError:
                        pass

    # v2.0 生命周期：全流程成功后删除全部过程文件（不保留 过程文件/），
    # 只留交付物（images/ sources/ JSON/MD）
    removed = 0
    for item in sorted(work.iterdir()):
        try:
            if item.is_dir():
                shutil.rmtree(item)
            else:
                item.unlink()
            removed += 1
        except OSError as e:
            print(f"[警告] 删除过程文件失败 {item.name}：{e}",
                  file=sys.stderr)
    print(f"[清理] 删除过程文件 {removed} 项（v2.0 成功即清理）",
          file=sys.stderr)
    for _p in (work, out / PROC_DIR):
        try:
            if _p.exists():
                shutil.rmtree(_p)
        except OSError as e:
            print(f"[警告] 清理目录失败 {_p.name}：{e}（不影响交付物，"
                  f"可手动删除）", file=sys.stderr)


def _estimate_usage(pptx: Path) -> dict:
    """粗略估算执行体量（页数/图片数/公式数），用于 token 预估。"""
    import re as _re
    import zipfile
    try:
        with zipfile.ZipFile(pptx) as z:
            slides = [n for n in z.namelist()
                      if _re.match(r"ppt/slides/slide\d+\.xml$", n)]
            media = [n for n in z.namelist()
                     if n.startswith("ppt/media/")]
            n_embeds = sum(1 for n in z.namelist()
                           if n.startswith("ppt/embeddings/"))
        return {"slides": len(slides), "media": len(media),
                "embeds": n_embeds}
    except Exception:
        return {"slides": 0, "media": 0, "embeds": 0}


def _plan_and_confirm(est: dict, skip: set) -> bool:
    """执行前分步说明 + token 估算 + 交互确认。返回是否继续。"""
    n_s, n_img, n_f = est["slides"], est["media"], est["embeds"]
    steps = []
    if "text" not in skip:
        steps.append(("1. 文本提取",
                      "逐页提取文本框/表格文本（排除页眉页脚/母版固定文本）。"
                      "本地执行，不消耗 API。"))
    if "formula" not in skip:
        steps.append(("2. 公式提取",
                      "解析公式 OLE/OMML 为 LaTeX。本地执行，不消耗 API。"))
    if "blocks" not in skip:
        est_cap = n_s * (1200 + 400 + 250)   # 每页块+页上下文+输出 ≈ 1850
        steps.append(("3. 可视逻辑块解析 + Semantic Captioning（消耗 Token）",
                      f"每页原子对象经空间聚类拆成 1~6 个可视逻辑块，"
                      f"逐块喂入视觉模型（qwen3.7-plus）判定类型并生成语义描述，"
                      f"附带该页文本/公式上下文。"
                      f"**预计消耗约 {est_cap//1000} 万~{int(est_cap*1.3)//1000} 万 Token**，"
                      f"耗时约 {max(1, n_s*8//60)} 分钟。"))
    if "related" not in skip:
        est_rel = n_img * (300 + 200 + 100)
        steps.append(("4. 可视逻辑块相关性过滤（消耗 Token）",
                      f"约 {n_img} 个块的描述与该页正文交 DeepSeek "
                      f"(deepseek-v4-flash) 判定相关性，剔除 logo/作者/单位等"
                      f"无关块。"
                      f"**预计消耗约 {max(1, est_rel//1000)} 万 Token**，"
                      f"耗时约 {max(1, n_img*3//60)} 分钟。"))
    if "author" not in skip:
        est_au = (n_s * 500 + n_f * 80 + n_s * 400)
        steps.append(("5. 教材文案生成（消耗 Token）",
                      f"文本/公式/可视逻辑块解读三份文档输入 DeepSeek "
                      f"(deepseek-v4-flash) 生成 {n_s} 页教材文案；"
                      f"原文超 300 字的页直接提取不扩写。"
                      f"**预计消耗约 {est_au//1000} 万~{int(est_au*1.5)//1000} 万 Token**，"
                      f"耗时约 {max(1, n_s*12//60)} 分钟。"))
    if "blocks_json" not in skip:
        steps.append(("6. 可视逻辑块 JSON 组装（含跨模态关系，消耗 Token）",
                      "按 pptx_multimodal_slide_v2.0 schema 组装 "
                      "visual_blocks.json，并调用 DeepSeek 生成每块与该页"
                      "文字的关系描述。"))

    print("\n========== 执行计划 ==========", file=sys.stderr)
    for name, desc in steps:
        print(f"  {name}", file=sys.stderr)
        print(f"      {desc}", file=sys.stderr)
    print("", file=sys.stderr)
    print("[提醒] 步骤 4、5 会消耗两个模型的 API Token（费用按 Token 计，"
          "以各官网为准）；", file=sys.stderr)
    print("       以上为粗略估算，实际以模型返回为准。", file=sys.stderr)
    try:
        ans = input("是否继续执行全部步骤？输入 y 继续，其他取消：").strip().lower()
    except EOFError:
        ans = "y"
    return ans == "y" or ans == "yes" or ans == "y\n"


def _report(out: Path, stem: str) -> None:
    """执行完后：每步工作量统计 + 结果文档使用说明。"""
    print("\n========== 执行结果汇总 ==========", file=sys.stderr)
    img_dir = out / "images"
    n_img = len(list(img_dir.glob("*"))) if img_dir.is_dir() else 0
    tb = out / f"{stem}_textbook.md"
    n_pages = 0
    if tb.is_file():
        n_pages = sum(1 for ln in tb.read_text(encoding="utf-8").splitlines()
                      if ln.startswith("## 第"))
    cap = out / f"{stem}_captions.md"
    n_cap = sum(1 for ln in
                (cap.read_text(encoding="utf-8").splitlines()
                 if cap.is_file() else []) if ln.startswith("### IMG"))
    bj = out / f"{stem}_visual_blocks.json"
    import json as _json
    bj_sum = ""
    if bj.is_file():
        try:
            s = _json.loads(bj.read_text(encoding="utf-8"))["summary"]
            bj_sum = f"{s['slides']} 页 / 块 {s['blocks_total']} / "
        except Exception:
            pass
    print(f"  图片：{n_img} 张 → images/", file=sys.stderr)
    print(f"  可视逻辑块解读：{n_cap} 条 → {stem}_captions.md", file=sys.stderr)
    print(f"  教材文案：{n_pages} 页 → {stem}_textbook.md", file=sys.stderr)
    print(f"  可视逻辑块 JSON：{bj_sum}→ {stem}_visual_blocks.json", file=sys.stderr)
    print("\n========== 结果文档使用说明 ==========", file=sys.stderr)
    print(f"  1. {stem}_textbook.md —— 教材文案：每页一节，可直接作为"
          "教材/课件文字素材；", file=sys.stderr)
    print(f"  2. {stem}_captions.md —— 可视逻辑块解读：每块的语义描述"
          "（类型/表达目标/作用/图文理解/教学用途），配图说明直接引用；",
          file=sys.stderr)
    print("  3. images/ —— 可视逻辑块渲染图：与 captions.md "
          "的 IMG 编号对应；", file=sys.stderr)
    print(f"  4. {stem}_visual_blocks.json —— 可视逻辑块全栈解析：每页的块"
          "（几何/拓扑/资源/类型/语义描述）与跨模态关系，供检索/知识库/"
          "RAG 使用；", file=sys.stderr)
    print(f"  5. sources/ —— 图块源资源：grpSp XML 段 + rldimg/ 资源图片"
          "+ 矢量源（v2.0 成功即清理全部过程文件，无 过程文件/ 目录）。",
          file=sys.stderr)
    print("", file=sys.stderr)


def _main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        prog="pptx-paser",
        description="整体运行：图片/公式/文本提取→图片AI解读→教材文案→图文绑定")
    ap.add_argument("pptx", help="输入的 .pptx 文件路径")
    ap.add_argument("-o", "--output", default="output",
                    help="生成结果目录（结果文件+过程文件/ 放这里）")
    ap.add_argument("--skip", default=None,
                    help="逗号分隔跳过步骤：text,formula,blocks,"
                         "related,author,blocks_json")
    ap.add_argument("--author-pages", default=None,
                    help="透传给教材文案指令：只生成指定页，如 '1,4'"
                         "（默认全部，测试用）")
    ap.add_argument("--no-install", action="store_true",
                    help="不自动安装缺失的组件库（默认自动 pip 安装）")
    ap.add_argument("--reset", action="store_true",
                    help="强制从头重跑：清空 state.json/日志与全部旧产物")
    ap.add_argument("--dry-run", action="store_true",
                    help="仅打印续跑/执行计划，不执行任何步骤")
    ap.add_argument("--version", action="version", version=VERSION)
    args = ap.parse_args(argv)

    pptx = Path(args.pptx)
    if not pptx.is_file():
        print(f"[错误] PPT 文件不存在：{pptx}", file=sys.stderr)
        return 2
    out = Path(args.output)
    out.mkdir(parents=True, exist_ok=True)
    stem = pptx.stem
    skip = set(s.strip() for s in (args.skip or "").split(",") if s.strip())

    print(f"\n===== pptx-paser 整体运行：{pptx.name} =====", file=sys.stderr)
    print(f"输出目录：{out}\n", file=sys.stderr)

    # 1) 环境检查（自动安装缺失组件 + 权重/LO 探测并逐项显示）
    print("\n========== 环境检查 ==========", file=sys.stderr)
    _check_env_all(no_install=args.no_install)

    # 2) KEY 检查（缺失→引导注册→重启提示→退出）
    if not _ensure_keys():
        return 2

    # 3) 断点状态：读取/初始化 state.json；--reset 清空后重建
    state = _load_state(out)
    if args.reset:
        print("[重置] --reset：清空旧产物与状态，从头执行", file=sys.stderr)
        _reset_out(out, stem)
        state = None
    if state is None or state.get("pptx") != str(pptx) or \
            state.get("stem") != stem:
        state = _init_state(pptx, out)
    else:
        # 换源/工具版本变化提示（对齐 word §3.7）
        import hashlib
        try:
            md5 = hashlib.md5(pptx.read_bytes()).hexdigest()
        except OSError:
            md5 = ""
        if state.get("doc_md5") and md5 and state["doc_md5"] != md5:
            print("[提示] 检测到源文档已变更（md5 不同），建议 --reset 从头重跑",
                  file=sys.stderr)
        if state.get("tool_version") and state["tool_version"] != VERSION:
            print(f"[提示] 工具版本变化（{state['tool_version']} → {VERSION}），"
                  f"建议 --reset 重跑", file=sys.stderr)
        _log(out, f"检测到断点状态："
                  f"{sum(1 for s in state['steps'].values() if s['status']=='done')}"
                  f"/{len(STEPS)} 步已完成，重跑本命令将自动续跑")
    # 3b) 续跑计划 + token 估算 + 用户确认
    est = _estimate_usage(pptx)
    plan = _build_plan(out, stem, state, skip)
    _print_plan(plan, est, skip)
    if args.dry_run:
        print("[DRY-RUN] 以上为续跑/执行计划，未执行任何步骤。",
              file=sys.stderr)
        print("[DRY-RUN] 确认后去掉 --dry-run 重跑即可按此计划执行。",
              file=sys.stderr)
        return 0
    todo = {s for s, (a, _) in plan.items() if a != "skip"}
    if todo and not _plan_and_confirm(est, skip):
        print("[取消] 已取消执行（重新运行本指令即可按断点继续）。",
              file=sys.stderr)
        return 0

    # 4) 依次执行（skip 跳过 / run 执行 / resume 断点续跑）
    work = out / "_proc"
    try:
        for s in STEPS:
            action, note = plan[s]
            if action == "skip":
                print(f"[跳过] {s}：{note}", file=sys.stderr)
                continue
            print(f"\n[开始] 步骤 {s}（{note or '执行'}）", file=sys.stderr)
            _set_step(out, state, s, "running", note)
            _log(out, f"步骤 {s} 开始（{action}）")
            try:
                _exec_step(s, args, work, stem, action)
            except RuntimeError as e:
                _set_step(out, state, s, "failed", str(e))
                _log(out, f"步骤 {s} 失败：{e}")
                raise
            _set_step(out, state, s, "done")
            _log(out, f"步骤 {s} 完成")
        # 产物归位
        _organize(out, work, stem)
        _log(out, "全部步骤完成，产物已归位")

    except RuntimeError as e:
        print(f"\n[错误] {e}", file=sys.stderr)
        print("[提示] 重新运行本命令将依据 state.json **自动续跑**"
              "未完成的步骤（或用 --skip 手动跳过）；--dry-run 可预览计划。",
              file=sys.stderr)
        return 1
    except Exception as e:
        print(f"\n[错误] 整体运行异常：{e}", file=sys.stderr)
        _log(out, f"整体运行异常：{e}")
        return 1

    print("\n===== pptx-paser 全部完成 =====", file=sys.stderr)
    print(f"结果目录：{out}", file=sys.stderr)
    print(f"  - {out / 'images'}", file=sys.stderr)
    print(f"  - {out / 'sources'}（含 rldimg/）", file=sys.stderr)
    print(f"  - {out / (stem + '_captions.md')}", file=sys.stderr)
    print(f"  - {out / (stem + '_textbook.md')}", file=sys.stderr)
    print(f"  - {out / (stem + '_visual_blocks.json')}", file=sys.stderr)
    print("（v2.0 成功即清理，无过程文件）", file=sys.stderr)
    _report(out, stem)
    return 0


def main() -> int:  # console
    banner("pptx-paser")
    rc = _main()
    banner_end("pptx-paser")
    return rc


if __name__ == "__main__":
    sys.exit(_main())
