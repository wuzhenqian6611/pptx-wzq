# -*- coding: utf-8 -*-
"""build_deck.py — 用 edu-blueprint 风格把 binding 内容重构成横版教材 deck。

作者：吴振谦 · wuzhenqian@nbu.edu.cn · QQ：38328063"""
import base64
import json
import re
from pathlib import Path

def _load_inputs(d, out, base):
    """加载 binding 并设置全局路径。"""
    global D, OUT, BASE, bj, P
    D, OUT, BASE = Path(d), Path(out), Path(base)
    bjs = sorted(D.glob("*_binding.json"))
    if not bjs:
        raise SystemExit(f"[错误] 目录内未找到 *_binding.json：{D}")
    bj = json.loads(bjs[0].read_text(encoding="utf-8"))
    P = {p["page"]: p for p in bj["pages"]}
    if not BASE.exists():
        raise SystemExit(f"[错误] 模板不存在：{BASE}（请从 lieflat-html-deck 技能获取 edu-blueprint 模板）")


def img(pg, idx=0, cls="ig-img"):
    """返回内嵌图片 figure（binding 第 pg 页第 idx 张图）。"""
    imgs = P[pg].get("images", [])
    if idx >= len(imgs):
        return ""
    f = imgs[idx]["file"]
    fp = D / "images" / f
    if not fp.exists():
        return ""
    b64 = base64.b64encode(fp.read_bytes()).decode("ascii")
    cap = re.sub(r"\s+", " ", (imgs[idx].get("caption") or ""))[:90]
    return (f'<div class="{cls}"><img src="data:image/png;base64,{b64}" '
            f'alt="{f}"/><span class="igcap">{cap}</span></div>')


def _esc(t):
    """HTML 转义 + 加粗渲染；$公式$ 保留给 MathJax。"""
    t = t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    t = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", t)
    return t


def tx(pg, n=260):
    t = re.sub(r"\s+", " ", P[pg]["text"]).strip()
    return _esc(t[:n] + ("…" if len(t) > n else ""))


def full(pg):
    """完整教材文案正文（去空白、转义），用于 .body 长文块。"""
    t = re.sub(r"\s+", " ", P[pg]["text"]).strip()
    return _esc(t)


def body(*pgs, split=2):
    """多页完整文案 → 分栏正文块。"""
    cols = []
    for pg in pgs:
        t = full(pg)
        if split > 1 and len(t) > 300:
            half = len(t) // 2
            cut = t.rfind("。", 0, half) + 1
            if cut < half * 0.5:
                cut = half
            cols.extend([t[:cut].strip(), t[cut:].strip()])
        else:
            cols.append(t)
    return ("<div class='body-grid'>" + "".join(
        f"<div class='body'>{c}</div>" for c in cols) + "</div>")


def _compose(P, D, slides, A):
    """按内容组装 14 张滑片（edu-blueprint 风格）。"""
    A = slides.append

    # 1 封面
    A(f'''<article class="slide cover active">
      <div class="frame">
        <div class="accent-bar"></div>
        <h1>第九章<br/>功率放大电路</h1>
        <p class="lead">在允许一定失真的前提下，向负载提供尽可能大的信号功率 —— 从"放大信号"到"驱动负载"的工程跨越。</p>
      </div>
      <div class="metrics">
        <div class="kpi"><b>输出功率</b><span>不失真 · 尽可能大</span></div>
        <div class="kpi warm"><b>效率</b><span>η = P<sub>om</sub> / P<sub>V</sub></span></div>
        <div class="kpi"><b>不失真</b><span>幅度 · 频带 · 失真约束</span></div>
      </div>
    </article>''')

    # 2 分隔 §9.1
    A('''<article class="slide">
      <div class="divider">
        <div class="badge">SECTION · 概述</div>
        <h1>§9.1<br/>功率放大电路概述</h1>
        <div class="sub">研究对象与方法都不同于小信号放大 —— 功率、效率、非线性失真成为核心矛盾</div>
      </div>
    </article>''')

    # 3 agenda
    A('''<article class="slide">
      <div class="agenda">
        <div class="arow on"><b>01</b><h3>研究的问题</h3><span>性能指标 · 对功放的要求 · 工作方式</span></div>
        <div class="arow"><b>02</b><h3>典型电路</h3><span>变压器耦合 · OTL · OCL · BTL</span></div>
        <div class="arow"><b>03</b><h3>分析计算</h3><span>输出功率 · 效率 · 晶体管的极限参数</span></div>
        <div class="arow"><b>04</b><h3>综合讨论</h3><span>集成运放 OCL · LM386 识别 · 故障分析</span></div>
      </div>
    </article>''')

    # 4 §9.1 概述+要求（sub 框架）
    A(f'''<article class="slide sub">
      <div class="topbar"><span class="sec">§9.1 概述</span><div class="tabs"><span class="ttab active">研究的问题</span><span class="ttab ghost">对功放的要求</span><span class="ttab ghost">工作方式</span></div></div>
      <div class="rail"><div class="spine"></div><div class="tab-v active">概述</div><div class="tab-v">典型电路</div><div class="tab-v">分析计算</div><div class="tab-v">讨论</div></div>
      <div class="subrail"><div class="subtab-v active">研究的问题</div><div class="subtab-v">要求</div><div class="subtab-v">工作方式</div></div>
      <div class="content">
        <div class="kicker">核心指标 · 输出功率与效率</div>
        <h2>功率放大电路研究的问题</h2>
        <div class="info-row mt">
          <div class="info"><div class="k">输出功率</div><h3>P<sub>om</sub></h3><p>不失真输出电压有效值 × 输出电流有效值的最大值</p></div>
          <div class="info"><div class="k">效率</div><h3>η</h3><p>输出功率与电源提供功率之比，越高越省电</p></div>
          <div class="info accent"><div class="k">矛盾</div><h3>大功率 × 高效率 × 低失真</h3><p>三者相互制约，是功放设计的核心矛盾</p></div>
        </div>
        <div class="grid-3 mt">
          <div class="card"><h3>输出功率尽可能大</h3><p>{tx(4,120)}</p></div>
          <div class="card"><h3>效率尽可能高</h3><p>{tx(4, 160)[60:]}</p></div>
          <div class="card accent"><h3>非线性失真允许</h3><p>大信号摆幅下允许一定失真，与小信号放大不同</p></div>
        </div>
        {body(3, 4)}
      </div>
    </article>''')

    # 5 工作方式
    A(f'''<article class="slide">
      <div class="topbar"><span class="sec">§9.1 概述</span><div class="tabs"><span class="ttab ghost">研究的问题</span><span class="ttab ghost">对功放的要求</span><span class="ttab navy active">晶体管的工作方式</span></div></div>
      <div class="rail"><div class="spine"></div><div class="tab-v active">概述</div><div class="tab-v">典型电路</div><div class="tab-v">分析计算</div><div class="tab-v">讨论</div></div>
      <div class="content">
        <div class="kicker">导通角决定工作方式与效率</div>
        <h2>晶体管的工作方式</h2>
        <div class="grid-3 mt">
          <div class="card"><h3>甲类 · 导通角 360°</h3><p>静态工作点在放大区中央，全周期导通；失真小、效率低（≤50%）</p></div>
          <div class="card hl"><h3>乙类 · 导通角 180°</h3><p>静态工作点压到截止区边缘，半周期导通；效率高（≤78.5%）但有交越失真</p></div>
          <div class="card"><h3>甲乙类 · 180°~360°</h3><p>略偏乙类的折中，兼顾效率与交越失真改善，功放最常用</p></div>
        </div>
        <div class="grid-2 mt">{img(7)}</div>
        <div class="note">乙类推挽：T₁ 正半周导通、T₂ 负半周导通，两只对称管交替工作，向负载提供完整波形。</div>
        {body(5)}
      </div>
    </article>''')

    # 6 种类
    A(f'''<article class="slide">
      <div class="topbar"><span class="sec">§9.1 概述</span><div class="tabs"><span class="ttab ghost">研究的问题</span><span class="ttab ghost">要求</span><span class="ttab navy active">功放种类</span></div></div>
      <div class="rail"><div class="spine"></div><div class="tab-v active">概述</div><div class="tab-v">典型电路</div><div class="tab-v">分析计算</div><div class="tab-v">讨论</div></div>
      <div class="content">
        <div class="kicker">从变压器到无变压器</div>
        <h2>功率放大电路的种类</h2>
        <div class="stat-grid mt">
          <div class="sg"><b>变压器耦合</b><span>阻抗匹配 + 推挽，但笨重 / 损耗大 / 频响差</span></div>
          <div class="sg"><b>OTL</b><span>单电源 + 大电容输出耦合，免去变压器</span></div>
          <div class="sg"><b>OCL</b><span>双电源直接耦合，低频特性最好</span></div>
          <div class="sg"><b>BTL</b><span>桥式推挽，浮地负载，同电源下电压摆幅翻倍</span></div>
        </div>
        <div class="grid-2 mt">{img(6, 0)}{img(6, 1)}</div>
        <div class="note">{tx(6, 150)}</div>
        {body(6)}
      </div>
    </article>''')

    # 7 典型电路对照
    A(f'''<article class="slide">
      <div class="topbar"><span class="sec">典型电路</span><div class="tabs"><span class="ttab navy active">乙类推挽</span><span class="ttab ghost">OTL</span><span class="ttab ghost">OCL</span><span class="ttab ghost">BTL</span></div></div>
      <div class="rail"><div class="spine"></div><div class="tab-v">概述</div><div class="tab-v active">典型电路</div><div class="tab-v">分析计算</div><div class="tab-v">讨论</div></div>
      <div class="content">
        <div class="kicker">四种结构的供电与耦合方式</div>
        <h2>典型电路对照</h2>
        <div class="grid-3 mt">
          <div class="card"><h3>OTL · 单电源</h3><p>{tx(8, 110)}</p></div>
          <div class="card"><h3>OCL · 双电源</h3><p>{tx(9, 110)}</p></div>
          <div class="card"><h3>BTL · 浮地桥式</h3><p>{tx(10, 110)}</p></div>
        </div>
        <div class="grid-3 mt">{img(8)}{img(9)}{img(10)}</div>
        {body(8, 9, 10)}
      </div>
    </article>''')

    # 8 比较表
    A(f'''<article class="slide">
      <div class="topbar"><span class="sec">典型电路</span><div class="tabs"><span class="ttab active">比较</span></div></div>
      <div class="rail"><div class="spine"></div><div class="tab-v">概述</div><div class="tab-v active">典型电路</div><div class="tab-v">分析计算</div><div class="tab-v">讨论</div></div>
      <div class="content">
        <div class="kicker">选用依据 · 供电 / 耦合 / 频响 / 效率</div>
        <h2>几种电路的比较</h2>
        <div class="table mt">
          <table>
            <thead><tr><th>电路</th><th>供电</th><th>输出耦合</th><th>低频特性</th><th>特点</th></tr></thead>
            <tbody>
              <tr><td>变压器耦合乙类推挽</td><td>单电源</td><td>变压器</td><td>差</td><td>笨重 · 效率低 · 已少用</td></tr>
              <tr><td>OTL</td><td>单电源</td><td>大电容</td><td>较好</td><td>免变压器 · 需大电容</td></tr>
              <tr><td>OCL</td><td>双电源</td><td>直接</td><td>最好</td><td>低频全通 · 需对称电源</td></tr>
              <tr><td>BTL</td><td>单/双电源</td><td>直接</td><td>好</td><td>电压摆幅 ×2 · 负载浮地</td></tr>
            </tbody>
          </table>
        </div>
        <div class="note">{tx(11, 200)}</div>
        {body(11)}
      </div>
    </article>''')

    # 9 分隔 §9.2
    A('''<article class="slide">
      <div class="divider">
        <div class="badge">SECTION · 分析计算</div>
        <h1>§9.2<br/>互补输出级的分析计算</h1>
        <div class="sub">三步法：确定 U<sub>om</sub> → 求输出功率与效率 → 核算极限参数</div>
      </div>
    </article>''')

    # 10 输出功率
    A(f'''<article class="slide">
      <div class="topbar"><span class="sec">分析计算</span><div class="tabs"><span class="ttab navy active">输出功率</span><span class="ttab ghost">效率</span><span class="ttab ghost">极限参数</span></div></div>
      <div class="rail"><div class="spine"></div><div class="tab-v">概述</div><div class="tab-v">典型电路</div><div class="tab-v active">分析计算</div><div class="tab-v">讨论</div></div>
      <div class="content">
        <div class="kicker">最大不失真输出电压决定 P<sub>om</sub></div>
        <h2>输出功率</h2>
        <div class="info-row mt">
          <div class="info"><div class="k">理想</div><h3>$U_{{om}}=V_{{CC}}$</h3><p>$P_{{om}}=\\dfrac{{V_{{CC}}^{{2}}}}{{2R_L}}$</p></div>
          <div class="info"><div class="k">实际</div><h3>$U_{{om}}=V_{{CC}}-U_{{CES}}$</h3><p>饱和压降不可忽略</p></div>
          <div class="info accent"><div class="k">一般式</div><h3>$P_{{om}}=\\dfrac{{U_{{om}}^{{2}}}}{{2R_L}}$</h3><p>输出电压峰值决定输出功率</p></div>
        </div>
        <div class="grid-2 mt">{img(14)}</div>
        <div class="note">{tx(13, 220)}</div>
        {body(13, 14)}
      </div>
    </article>''')

    # 11 效率
    A(f'''<article class="slide">
      <div class="topbar"><span class="sec">分析计算</span><div class="tabs"><span class="ttab ghost">输出功率</span><span class="ttab navy active">效率</span><span class="ttab ghost">极限参数</span></div></div>
      <div class="rail"><div class="spine"></div><div class="tab-v">概述</div><div class="tab-v">典型电路</div><div class="tab-v active">分析计算</div><div class="tab-v">讨论</div></div>
      <div class="content">
        <div class="kicker">半波正弦电源电流取平均值</div>
        <h2>效率 η = P<sub>om</sub> / P<sub>V</sub></h2>
        <div class="info-row mt">
          <div class="info"><div class="k">乙类理想</div><h3>η<sub>max</sub> ≈ 78.5%</h3><p>π/4 · 峰值输出时取得</p></div>
          <div class="info"><div class="k">甲类</div><h3>η<sub>max</sub> ≤ 50%</h3><p>静态功耗大是主因</p></div>
          <div class="info accent"><div class="k">结论</div><h3>甲乙类常用</h3><p>效率接近乙类，交越失真明显改善</p></div>
        </div>
        <div class="grid-2 mt">{img(15)}</div>
        <div class="note">{tx(15, 200)}</div>
        {body(15)}
      </div>
    </article>''')

    # 12 极限参数
    A(f'''<article class="slide">
      <div class="topbar"><span class="sec">分析计算</span><div class="tabs"><span class="ttab ghost">输出功率</span><span class="ttab ghost">效率</span><span class="ttab navy active">极限参数</span></div></div>
      <div class="rail"><div class="spine"></div><div class="tab-v">概述</div><div class="tab-v">典型电路</div><div class="tab-v active">分析计算</div><div class="tab-v">讨论</div></div>
      <div class="content">
        <div class="kicker">管子安全运行的三道红线</div>
        <h2>晶体管的极限参数</h2>
        <div class="problem-n mt">
          <div class="prob"><div class="pnum">①<br/>管压降</div><div class="pbody"><h3>$U_{{CEO}} \\geq 2V_{{CC}}$</h3><p>截止管承受最大管压降约 2V<sub>CC</sub></p></div></div>
          <div class="prob"><div class="pnum">②<br/>集电极电流</div><div class="pbody"><h3>$I_{{CM}} \\geq V_{{CC}}/R_L$</h3><p>峰值电流约束</p></div></div>
          <div class="prob"><div class="pnum">③<br/>功耗</div><div class="pbody"><h3>$P_{{Tmax}} \\approx 0.2\\,P_{{ommax}}$</h3><p>在 U<sub>OM</sub>≈0.6V<sub>CC</sub> 时取得</p></div></div>
        </div>
        <div class="grid-2 mt">{img(16)}{img(18)}</div>
        {body(16, 17)}
      </div>
    </article>''')

    # 13 讨论
    A(f'''<article class="slide">
      <div class="topbar"><span class="sec">综合讨论</span><div class="tabs"><span class="ttab navy active">讨论一</span><span class="ttab ghost">讨论二</span><span class="ttab ghost">讨论三</span></div></div>
      <div class="rail"><div class="spine"></div><div class="tab-v">概述</div><div class="tab-v">典型电路</div><div class="tab-v">分析计算</div><div class="tab-v active">讨论</div></div>
      <div class="content">
        <div class="kicker">综合应用 · 识别 · 故障</div>
        <h2>综合讨论</h2>
        <div class="grid-3 mt">
          <div class="card"><h3>讨论一 · 运放+准互补OCL</h3><p>{tx(18, 90)}</p></div>
          <div class="card"><h3>讨论二 · 功放类型识别</h3><p>{tx(19, 90)}</p></div>
          <div class="card accent"><h3>讨论三 · 故障分析</h3><p>{tx(20, 90)}</p></div>
        </div>
        <div class="grid-3 mt">{img(18)}{img(19, 0)}{img(20)}</div>
        {body(18, 19, 20)}
      </div>
    </article>''')

    # 14 结尾
    A('''<article class="slide">
      <div class="closing">
        <div class="q">从"放大信号"到"驱动负载"——<br/>功率、效率与不失真，<b>一场工程性的权衡。</b></div>
        <div class="sub">第九章 功率放大电路 · 由 pptx-wzq 教材流水线生成</div>
      </div>
    </article>''')



def main() -> int:  # console
    from pptx_wzq.cli_common import banner, banner_end
    banner("pptx-deck")
    import argparse
    ap = argparse.ArgumentParser(prog="pptx-deck",
                                 description="生成教育蓝图风格教学 Deck")
    ap.add_argument("--dir", default="analog9",
                    help="产物目录（含 *_binding.json 与 images/）")
    ap.add_argument("--template", default="",
                    help="edu-blueprint 模板 HTML 路径（必需）")
    ap.add_argument("--out", default=None, help="输出 HTML 路径")
    args = ap.parse_args()
    _load_inputs(args.dir, args.out, args.template)
    slides = []
    A = slides.append
    _compose(P, D, slides, A)
    # ---- 组装（复用模板 style + 导航脚本，替换 slides）----
    base = BASE.read_text(encoding="utf-8")
    style = re.search(r"<style>.*?</style>", base, re.S).group(0)
    script = re.search(r"<script>.*?</script>", base, re.S).group(0)
    mathjax = '''<script>
    window.MathJax = { tex: { inlineMath: [['$','$']], displayMath: [['$$','$$']] } };
    </script>
    <script src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-svg.js" async></script>
    '''
    html = f'''<!doctype html>
    <html lang="zh-CN">
    <head>
    <meta charset="utf-8"/>
    <meta name="viewport" content="width=device-width, initial-scale=1"/>
    <meta name="generator" content="Lieflat HTML Deck"/>
    <meta name="template-origin" content="Lieflat HTML Deck template"/>
    <title>第九章 功率放大电路 · 教材精讲（全文版）</title>
    <link rel="preconnect" href="https://fonts.googleapis.com"/>
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin/>
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@400;500;700;900&family=Montserrat:wght@500;600;700&display=swap" rel="stylesheet"/>
    {style}
    <style>
    /* 补充：图片容器（教材图卡） */
    .grid-2 {{ display:grid; grid-template-columns:1fr 1fr; gap:14px; }}
    .grid-3 .ig-img, .grid-2 .ig-img {{ background:#fff; border:1px solid var(--line); border-radius:8px; padding:8px; display:flex; flex-direction:column; align-items:center; }}
    .ig-img img {{ width:100%; height:170px; object-fit:contain; background:#fff; border-radius:4px; }}
    .igcap {{ font-size:11px; color:var(--muted); margin-top:6px; line-height:1.55; text-align:left; max-height:52px; overflow:hidden; }}
    .grid-3 .ig-img img {{ height:120px; }}
    @media (max-width:900px){{ .grid-2{{ grid-template-columns:1fr; }} }}
    /* 补充：完整正文长文块 */
    .body-grid {{ display:grid; grid-template-columns:1fr 1fr; gap:12px 18px; margin-top:12px; }}
    .body {{ font-size:12.5px; line-height:1.72; color:var(--nearblack); text-align:justify;
            background:var(--mist); border:1px solid var(--line); border-radius:8px;
            padding:10px 14px; max-height:235px; overflow-y:auto; }}
    .body p {{ margin:5px 0; }}
    @media (max-width:900px){{ .body-grid{{ grid-template-columns:1fr; }} }}
    </style>
    {mathjax}
    </head>
    <body data-theme="light">
    <main class="stage" aria-label="第九章 功率放大电路">
      <section class="slides">
    {chr(10).join(slides)}
      </section>
      <div class="counter" id="counter">01 / 14</div>
      <div class="nav">
        <button id="prev" aria-label="Previous">‹</button>
        <button id="next" aria-label="Next">›</button>
      </div>
      <div class="progress"><i id="progress"></i></div>
    </main>
    {script}
    </body>
    </html>'''
    OUT.write_text(html, encoding="utf-8")
    print(f"[OK] deck 已生成：{OUT}  ({OUT.stat().st_size//1024} KB, {len(slides)} 张滑片)")


    banner_end("pptx-deck")
    return 0


if __name__ == "__main__":
    main()
