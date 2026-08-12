"""gen_html.py — 把 PPT-Paser 产物（binding.json + images/）生成为
图文并茂、公式正确渲染的单文件教材 HTML。

用法：
    gen_html.py <产物目录> [-o out.html] [--title 章标题]

数据来源：<名>_binding.json（每页文案+图片列表+图注）+ images/（PNG）。
公式用 MathJax v3（CDN）渲染 $...$/$$...$$；图片 base64 内嵌，
单文件可随处打开。
"""
from __future__ import annotations

import argparse
import base64
import json
import re
import sys
from pathlib import Path


# --------------------------------------------------------------------------
# Markdown 轻量 → HTML（保留 $ 公式给 MathJax）
# --------------------------------------------------------------------------
def _inline_md(s: str) -> str:
    """行内：转义 HTML、加粗/斜体/行内代码；$...$ 公式原样保留。"""
    # 公式占位保护（$..$ 与 $$..$$）
    placeholders = {}
    def _protect(m):
        key = f"@@F{len(placeholders)}@@"
        placeholders[key] = m.group(0)
        return key
    s = re.sub(r"\$\$.*?\$\$", _protect, s, flags=re.S)
    s = re.sub(r"(?<!\$)\$[^$\n]+?\$(?!\$)", _protect, s)
    s = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    s = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", s)
    s = re.sub(r"(?<!\*)\*([^*\n]+?)\*(?!\*)", r"<em>\1</em>", s)
    s = re.sub(r"`([^`\n]+?)`", r"<code>\1</code>", s)
    for k, v in placeholders.items():
        s = s.replace(k, v)
    return s


def _md_blocks(text: str) -> list:
    """md 文本 → HTML 块列表（段落/公式块/列表）。"""
    blocks = []
    lines = text.splitlines()
    i, n = 0, len(lines)
    while i < n:
        ln = lines[i]
        s = ln.strip()
        if not s:
            i += 1
            continue
        if s.startswith("$$"):
            buf = []
            if s == "$$":
                i += 1
                while i < n and lines[i].strip() != "$$":
                    buf.append(lines[i])
                    i += 1
                i += 1  # 跳过闭合 $$
            else:
                buf.append(s[2:])
                i += 1
            blocks.append('<div class="formula">$$\n'
                          + "\n".join(buf).strip() + "\n$$</div>")
            continue
        if s.startswith("- "):
            items = []
            while i < n and lines[i].strip().startswith("- "):
                items.append("<li>" + _inline_md(lines[i].strip()[2:]) + "</li>")
                i += 1
            blocks.append("<ul>" + "".join(items) + "</ul>")
            continue
        blocks.append("<p>" + _inline_md(s) + "</p>")
        i += 1
    return blocks


# --------------------------------------------------------------------------
# 页面标题提取
# --------------------------------------------------------------------------
# --------------------------------------------------------------------------
# 页面标题：优先取 PPT 原生标题（过程文件 texts.md 的 [标题] 条目）
# --------------------------------------------------------------------------
def _load_ppt_titles(out_dir: Path) -> dict:
    """从 <名>_texts.md 解析每页标题占位符 → {page: title}。"""
    cands = sorted((out_dir / "过程文件" / "text").glob("*_texts.md")) \
        if (out_dir / "过程文件" / "text").is_dir() else []
    cands += sorted(out_dir.glob("*_texts.md"))
    titles = {}
    for md in cands:
        cur = None
        for ln in md.read_text(encoding="utf-8").splitlines():
            m = re.match(r"^##\s*第\s*(\d+)\s*页\s*$", ln.strip())
            if m:
                cur = int(m.group(1))
                continue
            t = re.match(r"^\|\s*TXT[\d\-]+\s*\|\s*标题\s*\|\s*(.*?)\s*\|$", ln)
            if t and cur is not None and cur not in titles:
                title = t.group(1).strip()
                # 跳过公式/空占位标题（如 \[、$$ 等）
                if title and not re.match(r"^[\\$\[\]{}]+$", title):
                    titles[cur] = title
    return titles


def _clean_title(t: str) -> str:
    """清洗：去多余空白；过长的截断。"""
    t = re.sub(r"\s+", " ", t).strip()
    return t[:30] + "…" if len(t) > 30 else t


def _page_title(text: str, page: int, ppt_titles: dict = None) -> str:
    if ppt_titles and page in ppt_titles:
        return _clean_title(ppt_titles[page])
    for ln in text.splitlines():
        s = ln.strip().strip("#").strip()
        if 2 <= len(s) <= 22 and not re.match(r"^[\\$\[\]{}]+$", s):
            return _clean_title(s)
    first = next((ln.strip() for ln in text.splitlines()
                  if ln.strip() and not re.match(r"^[\\$\[\]{}]+$", ln.strip())),
                 "")
    return (first[:16] + "…") if len(first) > 16 else (first or f"第 {page} 节")


# --------------------------------------------------------------------------
# HTML 组装
# --------------------------------------------------------------------------
def build_html(stem: str, chapter_title: str, pages: list,
               img_dir: Path, ppt_titles: dict = None) -> str:
    toc_items, sections = [], []
    for p in pages:
        pg, text, images = p["page"], p["text"], p.get("images", [])
        if not text and not images:
            continue
        title = _page_title(text, pg, ppt_titles)
        toc_items.append(f'<a class="toc-item" href="#p{pg}">'
                         f'<span class="toc-no">{pg:02d}</span>'
                         f'<span class="toc-title">{title}</span></a>')
        body = "".join(_md_blocks(text))
        figs = []
        for im in images:
            f = im["file"]
            fp = img_dir / f
            if not fp.exists():
                continue
            b64 = base64.b64encode(fp.read_bytes()).decode("ascii")
            cap = (im.get("caption") or "").strip().replace("\n", " ")
            cap = re.sub(r"\s+", " ", cap)[:120]
            figs.append(
                f'<figure class="fig">'
                f'<img src="data:image/png;base64,{b64}" alt="{f}">'
                f'<figcaption>{_inline_md(cap)}</figcaption></figure>')
        figs_html = "".join(figs)
        sections.append(
            f'<section class="page-card" id="p{pg}">'
            f'<h2><span class="sec-no">{pg:02d}</span>{_inline_md(title)}</h2>'
            f'{body}{figs_html}'
            f'<a class="backtop" href="#toc">返回目录 ↑</a>'
            f'</section>')

    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{chapter_title}｜{stem}</title>
<script>
window.MathJax = {{
  tex: {{ inlineMath: [['$', '$'], ['\\\\(', '\\\\)']],
         displayMath: [['$$', '$$'], ['\\\\[', '\\\\]']] }},
  svg: {{ fontCache: 'global' }}
}};
</script>
<script src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-svg.js" async></script>
<style>
:root {{
  --primary: #0b4f8a; --primary-d: #083a68; --accent: #e8a33d;
  --bg: #f2f5f9; --card: #ffffff; --ink: #22303f; --ink-2: #5a6b7c;
  --line: #dfe7f0;
}}
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
html {{ scroll-behavior: smooth; }}
body {{
  font-family: "Noto Serif SC", "Songti SC", SimSun, serif;
  background: var(--bg); color: var(--ink); line-height: 1.9;
}}
/* 封面 */
.cover {{
  background: linear-gradient(135deg, #0b4f8a 0%, #0a3d6b 55%, #07294a 100%);
  color: #fff; padding: 84px 32px 60px; text-align: center;
  border-bottom: 5px solid var(--accent);
}}
.cover .course {{ font-size: 15px; letter-spacing: 6px; opacity: .85;
  text-transform: uppercase; }}
.cover h1 {{ font-size: 46px; margin: 18px 0 10px; font-weight: 700;
  letter-spacing: 3px; }}
.cover .sub {{ font-size: 18px; opacity: .9; }}
.cover .meta {{ margin-top: 26px; font-size: 13px; opacity: .75; }}
/* 目录 */
.toc-wrap {{ max-width: 980px; margin: 34px auto 8px; padding: 0 20px; }}
.toc-wrap h2 {{ color: var(--primary-d); border-bottom: 2px solid var(--accent);
  display: inline-block; padding-bottom: 4px; margin-bottom: 16px; }}
.toc {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(210px, 1fr));
  gap: 10px; }}
.toc-item {{ display: flex; align-items: center; gap: 10px;
  background: var(--card); border: 1px solid var(--line); border-radius: 10px;
  padding: 10px 14px; text-decoration: none; color: var(--ink);
  transition: all .18s; box-shadow: 0 1px 3px rgba(10,40,80,.06); }}
.toc-item:hover {{ transform: translateY(-2px); box-shadow: 0 6px 16px rgba(10,40,80,.14);
  border-color: var(--primary); }}
.toc-no {{ background: var(--primary); color: #fff; border-radius: 6px;
  font-size: 12px; padding: 2px 7px; font-family: Consolas, monospace; }}
.toc-title {{ font-size: 14px; overflow: hidden; text-overflow: ellipsis;
  white-space: nowrap; }}
/* 正文 */
main {{ max-width: 880px; margin: 0 auto; padding: 8px 20px 60px; }}
.page-card {{ background: var(--card); border: 1px solid var(--line);
  border-radius: 14px; padding: 34px 40px; margin: 26px 0;
  box-shadow: 0 3px 14px rgba(10,40,80,.07); position: relative; }}
.page-card h2 {{ color: var(--primary-d); font-size: 24px; margin-bottom: 18px;
  padding-left: 14px; border-left: 5px solid var(--accent); }}
.sec-no {{ color: var(--accent); font-family: Consolas, monospace;
  margin-right: 8px; font-size: 20px; }}
.page-card p {{ margin: 12px 0; text-align: justify; font-size: 16.5px; }}
.page-card strong {{ color: var(--primary-d); }}
.formula {{ text-align: center; margin: 18px 0; overflow-x: auto; }}
.fig {{ margin: 22px auto; max-width: 640px; text-align: center; }}
.fig img {{ max-width: 100%; border: 1px solid var(--line); border-radius: 8px;
  box-shadow: 0 4px 12px rgba(10,40,80,.1); background: #fff; }}
.fig figcaption {{ margin-top: 8px; font-size: 13px; color: var(--ink-2); }}
ul {{ margin: 10px 0 10px 26px; }}
li {{ margin: 4px 0; }}
code {{ background: #eef3f9; padding: 1px 6px; border-radius: 4px;
  font-family: Consolas, monospace; font-size: .9em; color: var(--primary-d); }}
.backtop {{ display: inline-block; margin-top: 14px; font-size: 12px;
  color: var(--ink-2); text-decoration: none; }}
.backtop:hover {{ color: var(--primary); }}
footer {{ text-align: center; color: var(--ink-2); font-size: 12.5px;
  padding: 30px 20px 50px; border-top: 1px solid var(--line); }}
@media (max-width: 700px) {{
  .cover h1 {{ font-size: 30px; }}
  .page-card {{ padding: 22px 18px; }}
}}
@media print {{
  body {{ background: #fff; }}
  .page-card {{ box-shadow: none; border: none; break-inside: avoid; }}
  .cover {{ -webkit-print-color-adjust: exact; }}
}}
</style>
</head>
<body>
<header class="cover" id="top">
  <div class="course">Electronics · 电子信息工程</div>
  <h1>{chapter_title}</h1>
  <div class="sub">{stem}</div>
  <div class="meta">由 PPT-Paser 流水线生成 · 共 {len(toc_items)} 节 ·
    图片 {len(list(img_dir.glob('*')))} 张 · 公式由 MathJax 渲染</div>
</header>
<nav class="toc-wrap" id="toc"><h2>目 录</h2><div class="toc">{''.join(toc_items)}</div></nav>
<main>{''.join(sections)}</main>
<footer>本章教材文档由 PPT 课件智能解析生成（PPT-Paser），
公式与图片均来自课件原始内容 · 仅供教学参考</footer>
</body>
</html>"""


def _main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        prog="gen-html",
        description="把 PPT-Paser 产物生成图文并茂的教材 HTML")
    ap.add_argument("dir", help="产物目录（含 <名>_binding.json 与 images/）")
    ap.add_argument("-o", "--output", default=None, help="输出 html 路径")
    ap.add_argument("--title", default=None, help="章标题（默认取自 stem）")
    args = ap.parse_args(argv)

    d = Path(args.dir)
    bjs = sorted(d.glob("*_binding.json"))
    if not bjs:
        print(f"[错误] 目录内未找到 *_binding.json：{d}", file=sys.stderr)
        return 2
    bj = json.loads(bjs[0].read_text(encoding="utf-8"))
    stem = bj["stem"]
    chapter = args.title or f"第 {stem.split('-')[0]} 章" \
        if "-" in stem else stem
    out = Path(args.output) if args.output else d / f"{stem}.html"
    ppt_titles = _load_ppt_titles(d)
    if ppt_titles:
        print(f"[标题] 已加载 {len(ppt_titles)} 页 PPT 原生标题",
              file=sys.stderr)
    html = build_html(stem, chapter, bj["pages"], d / "images", ppt_titles)
    out.write_text(html, encoding="utf-8")
    print(f"[OK] 教材 HTML 已生成：{out}")
    print(f"     章节数 {len(bj['pages'])} · 文件大小 {out.stat().st_size//1024} KB")
    return 0


def main() -> int:  # console
    return _main()


if __name__ == "__main__":
    sys.exit(_main())
