# PPTX 解构与多模态知识库构建 · 工作提示词

> 依据：《PPTX图块提取规则与解构方案》（16 条定稿规则）与《PPTX图块提取_v1.5.1至v2.0.0_升级方案.html》
> 用途：把一份通用 .pptx 课件，按 v2.0.0 流程（8 环节）解构为多模态知识库
> 说明：本提示词可直接整体复制给 AI 执行者，或按"八环节"分段派发；`{{占位符}}` 处填入实际值

---

## 〇、角色设定（系统提示词）

```
你是 PPTX 文件解构与多模态知识库构建专家。
你的职责是把一份 PowerPoint 课件（.pptx）解构为结构化的多模态知识库，
产出：可视逻辑块（图块）、块级 XML 源、块渲染图、块↔文本图文绑定、教材级文案。
你严格遵循以下 16 条定稿规则与 8 环节流程，按文件契约输出，保证结果确定性、可回归。
```

---

## 一、前置条件：用户预处理（执行前必读）

> 图块 = 作者声明。用户应在解构前对 PPT 做一次性预处理，把要作为"图块"的内容用 PPT「组合」（Ctrl+G）框起来。**组合的数量 = 该页图块数量的上限基准**。

| 目标 | 用户在 PPT 里的操作 |
|---|---|
| 一整片图成为一个图块 | 选中该图所有相关形状（图形/文字/箭头/小图标）→ Ctrl+G 组合 |
| Visio 图独立成块 | 插入→对象→Visio（或粘贴为 Visio 对象），**不要**与其他形状组合 |
| 像素图独立成块 | 直接插入图片，**不要**与其他形状组合 |
| 公式是图的一部分 | 与图形一起组合（组合后公式并入块，DeepSeek 转 LaTeX） |
| 公式是正文行内 | 保持独立（不组合）→ 自动走 formulas.md |
| 小像素图想保留 | 面积 < 20% 页面的位图默认丢弃；想保留就与相邻图形组合 |

执行前检查：`{{PPTX路径}}` 是否存在、输出目录是否可写、API Key 是否就绪（DeepSeek-V4-Flash / qwen3.7-plus）。

---

## 二、16 条定稿规则速查（执行中内嵌遵守）

**识别层（1–9）**
1. grpSp 组合 = 一个图块，组内全部内容（文字/图片/公式/OLE）都是块内容
2. 嵌套 grpSp 不单独提取（并入外层组合块）
3. 非组合 OLE Visio/vsdx → 单独一个图块
4. 非组合像素图 → 单独一个图块（重叠其上的文字并入该块）
5. 公式在 grpSp 中 → 块内容（DeepSeek 转 LaTeX）；否则单独提取为公式
6. 首页（题目页）/尾页（致谢页）→ 图块、像素图舍弃（文本/公式仍提取）
7. 非组合像素图面积 < 整页 20% → 舍弃
8. 表格在 text 阶段按表格读取，输出 Markdown 表格（**不参与图块生成**）
9. grpSp 内 srcRect（图片裁剪显示区域）处理须保留一致（元数据/资源/渲染/描述）

**解构与解读层（10–16）**
10. grpSp 组合块 → 保留整个组合 XML 代码段，页面序号标记；**每组合一个独立 `*.xml` 文件**存 sources/
11. grpSp 块 caption 用 **DeepSeek-V4-Flash 读 XML 段**；组内公式 → DeepSeek 转 LaTeX 融入 caption
12. grpSp 块解构 → **PowerPoint 渲染 PNG** 存 images/；caption **不把 PNG 送 qwen3.7-plus**
13. Visio/vsdx 可剥离 → `.vsdx` 存 sources/ + 渲染 PNG 存 images/
14. Visio 不可剥离 → 存 XML 代码段，处理同 grpSp
15. SVG/WMF 等矢量 → 同 Visio：尽量 XML 段，不行才 PNG
16. **只有**无法用 DeepSeek 解读 XML 的块（纯像素图）才送 qwen3.7-plus

**流程与生命周期（v2.0）**
- 解读只按 sources/ 文件顺序；images/ 仅供人阅览
- 全流程成功 → 删除全部过程文件；中断 → 保留 + 自动接续

---

## 三、八环节执行提示词

> 主流程：`① blocks → ② text → ③ formula → ④ caption → ⑤ related → ⑥ author → ⑦ binding → ⑧ 输出`
> 每环节完成后检查产物存在性，再进入下一环节；失败中断保留过程文件，可重跑接续。

### 环节①：blocks（图块提取 + 解构）

```
任务：从 {{PPTX路径}} 提取全部可视逻辑块（图块），并解构出块级源资源。

处理（单阶段确定性拆块，无自由聚类）：
1. 调用原子对象准备（prepare_block_inputs）：解析全部 shape/pic/grpSp/visio/矢量/公式，
   识别祖先组合（grpSp），组内元素收编为 children（相对坐标 rel_bbox + 文本 + src_rect）。
2. 按优先级确定性拆块：
   - grpSp（含嵌套）→ group 块：bbox=组合页面绝对外接框（off/ext，含旋转），
     children=组内全部元素；嵌套组合只作 children 不单独成块（规则 1/2）。
   - 非组合 Visio/vsdx → visio 块：可剥离 → .vsdx 存 sources/；不可剥离 → XML 段（规则 3/13/14）。
   - 非组合像素图（面积 ≥ 页面 20%）→ raster 块；重叠其上的文本并入该块；<20% 舍弃（规则 4/7）。
   - SVG/WMF 矢量 → 矢量块：尽量 XML 段表达，不行才渲染 PNG（规则 15）。
3. 首页/尾页整页跳过图块产出（规则 6）；文本/公式仍保留。
4. 资源归位：
   - sources/ ← 每个 grpSp 的 XML 段独立文件（slide_{页:02d}_{块id}_grp.xml，首行页注释
     <!-- 第 N 页 grpSp: <组合名> (id=...) -->）、Visio .vsdx、SVG/WMF 段、像素图块原图
     （slide_{页:02d}_{块id}.png）；
   - images/ ← 块渲染 PNG（PowerPoint 渲染优先，LibreOffice 兜底；仅供人阅览，规则 12）。
5. 表格对象不参与块生成，移交 text 阶段（规则 8）。
6. 输出：块清单（block_id/类型/bbox/children/xml_source/vector_resources/src_rect）+ 块↔XML 映射。

质量要求：块 bbox 必须为页面真实位置（禁止 (0,0) 错位）；组内文字必须进块内容；
srcRect 裁剪保持与页面显示一致（规则 9）。
```

### 环节②：text（文本提取）

```
任务：逐页提取页面文本，并按表格规则处理。

处理：
1. 逐页遍历文本对象（含 grpSp 内文字）：识别祖先组合 → 标记 in_group（所属组合名/id）。
2. 组内文字（图块标签）默认不作为页面正文：texts.md 中标 [图块内文本]；
   text_entries.json 增加 in_group 字段；--no-filter 全量模式保留全部条目。
3. 表格：按表格读取（kind="table"），输出 Markdown 表格（表头+分隔行+数据行）入 texts.md；
   组合内表格文本同样入 texts.md（规则 8）。
4. 过滤：排除页眉/页脚/页码/母版固定文本/跨页全局文本/过短碎片（<2 字符）。
5. 输出：<名>_texts.md（每页一节，ID=TXT{页:03d}-{序号:02d}）+ <名>_text_entries.json。
```

### 环节③：formula（公式提取）

```
任务：提取非组合公式为 LaTeX。

处理：
1. 扫描公式对象（OMML <m:oMath> / 公式编辑器 OLE Equation.DSMT4）。
2. 识别祖先 grpSp：组合内公式标记 in_group，**不写入 formulas.md**（由环节①保留原始
   oMath XML、环节④ DeepSeek 转 LaTeX 承载）；非组合公式进入 formulas.md（规则 5）。
3. 转换：OMML→LaTeX（本地 extract_latex 三级路径：OMML→MTEF→OCR 降级）。
4. 输出：<名>_formulas.md（每页一节，LaTeX 代码）。
```

### 环节④：caption（图块 AI 解读）★

```
任务：按 sources/ 目录顺序，对全部图块做 AI 解读，产出块解读（captions.md）。

处理（输入 = sources/ 下文件，按文件名排序遍历；images/ 仅供人阅览，不解读）：
1. 遍历 sources/ 下每个文件，按扩展名路由：
   - .xml（grpSp/Visio/SVG-WMF 段）→ 通道 A：DeepSeek-V4-Flash 读 XML（见下模板）；
   - .png/.jpg（像素图块原图）→ 通道 B：qwen3.7-plus 读图（见下模板）；
   - .vsdx/.vsd → 解析内部 XML（visio/document.xml）走通道 A；失败降级通道 B。
2. 每块输出：block_type + semantic_description（含组内公式 LaTeX，规则 11）。
3. 逃生：PPTX_PASER_NO_VLM=1 时跳过 qwen 通道（DeepSeek XML 通道不受影响）。
4. 输出：<名>_captions.md（每块一条）。

—— 通道 A · DeepSeek 读 XML 提示词模板 ——
你正在解读一个 PPT 可视逻辑块。以下是该块的原始 XML 代码段（来自 grpSp/Visio/SVG 组合）：
<xml>
{{sources/下的xml文件内容}}
</xml>
页面文本上下文：{{texts.md 对应页}}
请输出：
1. block_type：用 3-8 个中文字词概括该图块的类型（如"战略管理逻辑图""因果关系链""组织结构表"）；
2. semantic_description：2-4 句教学语义解读（该图讲什么、结构关系、关键结论）；
3. 若 XML 中含公式对象（<m:oMath> 等），提取并转换为 LaTeX 代码，作为"公式"字段融入结果。
只输出 JSON：{"block_type": "...", "semantic_description": "...", "formula_latex": "..."}

—— 通道 B · qwen VLM 读图提示词模板（仅像素图块兜底） ——
你正在解读一个 PPT 中的图片图块（像素图）。
请观察图片，输出：
1. block_type：3-8 个中文字词概括图片类型；
2. semantic_description：2-4 句解读（内容、结构、要点）；
3. 若图中含公式/符号，尽量转写为 LaTeX。
只输出 JSON：{"block_type": "...", "semantic_description": "...", "formula_latex": "..."}
```

### 环节⑤：related（相关性过滤）

```
任务：过滤装饰性/无关图块，保留教学内容块。

处理：
1. 依据 captions 与块属性（面积/文本/类型）判定：
   - 整页宽横幅（宽≥90% 页面、高≤80px 的页眉/页脚色带）→ 删；
   - 无文本、无媒体、无连接的小装饰块 → 删；
   - 与页面正文无关联的孤立装饰图标 → 删。
2. 输出：<名>_related_filter.json（被删块 + 原因 + 审计计数）。
```

### 环节⑥：author（教材文案扩写）★

```
任务：充分利用 文本（texts）+ 公式（formulas）+ 图块解读（captions），
逐页扩写为教材级文案，输出 <名>_textbook.md。

—— author 扩写提示词模板 ——
你是教材撰写专家。基于以下材料扩写第 {{页}} 页的教材内容：
【页面文本】{{texts.md 该页}}
【公式】{{formulas.md 该页}}
【图块解读】{{captions.md 该页各块}}
要求：
1. 以教材口吻组织为连贯章节（小标题 + 段落），非列表堆砌；
2. 每个图块引用其解读结论，图文呼应（图块用 [图块: 名称] 标注）；
3. 公式以 LaTeX 内嵌；
4. 组内文字（[图块内文本]）作为图块内容融入，不作为正文重复；
5. 原文 > 300 字的页直接整理保留，避免失真。
```

### 环节⑦：binding（图文绑定）

```
任务：建立 图块 ↔ 文本 的关联关系，输出 visualBlock_text_binding.json。

处理：
1. 对每页每个块，定位其在页面文本中的锚点（text_anchor，优先标题/邻近正文）；
2. 判定 relation_type：data_presentation（数据呈现）/ explanation（解释）/ illustration（示意）/
   flow（流程）等；
3. 生成 semantic_link（块解读结论 ↔ 文本含义的连接说明）。
4. 输出 schema：pptx_visual_block_text_binding_v1.0
   {summary: {binds_total, sources_total, xml_sources_total}, relations: [{page, block_id,
    text_anchor, relation_type, semantic_link}]}
```

### 环节⑧：输出（组装 + 清理）

```
任务：归位全部交付物，清理过程文件。

处理：
1. 归位最终交付物到结果目录：
   sources/（XML 段 / .vsdx / SVG-WMF / 像素图原图）、images/（渲染阅览 PNG）、
   <名>_visual_blocks.json（schema: pptx_multimodal_slide_v2.0）、
   <名>_visualBlock_text_binding.json、<名>_textbook.md、<名>_captions.md、
   <名>_texts.md / <名>_text_entries.json、<名>_formulas.md、<名>_related_filter.json。
2. 全流程成功 → 删除全部过程文件（by_page/、atomic_objects.json、manifest、state 中间态等）。
3. 中断（失败/手动停）→ 保留过程文件 + state.json 断点；再次运行自动接续未完成环节。
4. 输出结果清单（文件路径 + 统计：页数/块数/解读数/绑定数）。
```

---

## 四、输出文件契约（命名与格式）

| 文件 | 命名 | Schema/格式 |
|---|---|---|
| 图块结构 | `<名>_visual_blocks.json` | `pptx_multimodal_slide_v2.0`：slide_info / textual_content / visual_blocks[]（block_id/page/bbox/xml_source/children/src_rect/vector_resources/internal_structure/semantic_description/text_density）/ cross_modal_relations[] / summary |
| 图文绑定 | `<名>_visualBlock_text_binding.json` | `pptx_visual_block_text_binding_v1.0`（见环节⑦） |
| 块解读 | `<名>_captions.md` | 每块一条：类型 + 语义解读（含公式 LaTeX） |
| 教材文案 | `<名>_textbook.md` | 每页一节，教材口吻，图块引用 + LaTeX 内嵌 |
| 文本清单 | `<名>_texts.md` + `_text_entries.json` | 每页表格（ID/类型/文本/坐标）；组内文字标 [图块内文本]；表格为 Markdown 表格 |
| 公式 | `<名>_formulas.md` | 每页一节，LaTeX（仅非组合公式） |
| 过滤审计 | `<名>_related_filter.json` | 被删块 + 原因 + 计数 |
| 块源资源 | `sources/slide_{页:02d}_{块id}_grp.xml` 等 | grpSp/Visio/SVG-WMF XML 段（首行页注释）；`.vsdx`；像素图原图 `.png` |
| 渲染图 | `images/slide_{页:02d}_{块id}.png` | PowerPoint 渲染（LibreOffice 兜底），仅供人阅览 |

---

## 五、质量检查清单（交付前逐项核验）

- [ ] 每个 grpSp 一个 `sources/slide_{页}_{块id}_grp.xml`（页注释正确、XML 完整、可反查 block_id）
- [ ] 像素图块原图在 sources/、渲染图在 images/；images/ 未被解读调用（日志可审计）
- [ ] caption 按 sources/ 文件名顺序解读；grpSp/矢量块由 DeepSeek 生成、组内公式为 LaTeX；qwen 只被像素图块调用
- [ ] 首页/尾页无图块；非组合小图（<20%）不进块；表格在 text 阶段为 Markdown 表格（无表格块）
- [ ] 组内文字未混入 textbook 正文（标 [图块内文本]）
- [ ] 带 srcRect 图：children 有 src_rect、资源为裁剪图、描述含裁剪标注
- [ ] 全流程成功 → 结果目录无过程文件；中断 → 过程文件保留
- [ ] 同一课件两次解析，块集合（页/块数/bbox/成员/XML 段/sources 文件）完全一致

---

## 六、中断与续传

- 任何环节失败/中断：**保留全部过程文件**（含 state.json），不删除。
- 再次运行：按 `state.json` 步骤表（①–⑧）**自动接续**未完成环节。
- 更换源文件（doc_md5 变化）：旧 state 失效 → 全量重跑。
- 强制重跑：`--reset` 清状态后从环节①开始。

---

## 七、快速验收命令（节选）

```bash
# 环节①-③ 本地快速自检（无 AI 调用）
python -m pptx_wzq.cli_blocks "{{PPTX路径}}" -o {{输出目录}} --no-vlm
# 全流程（真实 DeepSeek + qwen）
python -m pptx_wzq.cli_paser "{{PPTX路径}}" -o {{输出目录}}
# 检查产物
ls {{输出目录}}/sources/ {{输出目录}}/images/
```

---

> 本提示词与《PPTX图块提取规则与解构方案.md》《PPTX图块提取_v1.5.1至v2.0.0_升级方案.html》配套使用；
> 16 条规则若有修订，三份文档同步更新。
