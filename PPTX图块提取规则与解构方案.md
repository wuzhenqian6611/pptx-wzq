# PPTX 图块提取规则与解构方案（通用）

> 对象：pptx-wzq 图块处理（`blocks` 步骤，`visual_blocks.extract_blocks` + `extract_pptx_images` 原子对象 + `cli_blocks` 解构/解读管线）
> 适用范围：**通用 PPTX 课件 → 多模态知识库**（不依赖特定课件）
> 文档 = ① 用户定稿的「预处理 + 图块提取/解构/解读规则」16 条 + ② 基于规则的失败根因分析与修改方案

---

## 0. 一页看懂：定稿规则（16 条）

> **预处理（解构前对 PPT 做的操作）→ 把「图块」从"算法猜测"变成"作者声明"。**
> 规则分两层：**A. 图块识别（组合即块、非组合按类型确定性处理）+ B. 图块解构与解读（XML 优先、模型按块路由）。**

### 0.1 图块识别规则（1–9）

| # | 规则 | 一句话解读 |
|---|---|---|
| 1 | **一个 grpSp 组合即是一个图块**，组合内所有内容（含文字/图片/公式/OLE）都读取为该图块的内容 | 组合 = 块的最高优先级容器，整体不可拆 |
| 2 | **嵌套在另一个 grpSp 中的组合，不单独提取** | 嵌套组合 = 外层图块的子结构，并入外层内容 |
| 3 | **OLE Visio / vsdx 对象 → 单独一个图块**，除非被包含在其他 grpSp 中 | Visio 工程图本身即完整图块；组内则作为块内容 |
| 4 | **像素图若不在 grpSp 中 → 单独提取为一个图块** | 每张位图都是独立可预期的图块；重叠其上的文字并入该块 |
| 5 | **公式若在 grpSp 中 → 作为图块内容；否则单独提取为公式** | 公式归宿唯一：组合内 → 块内容（DeepSeek 转 LaTeX）；组合外 → formulas.md |
| 6 | **第一页（题目页）与最后一页（致谢页）：图块内容、像素图等舍弃** | 封面/封底不产图块；文本/公式仍正常提取 |
| 7 | **未包含在 grpSp 中的像素图，面积 < 整页 20% → 舍弃** | 小像素图（装饰图标等）不进块 |
| 8 | **表格对象在 text 文本提取阶段按表格读取**，输出 Markdown 表格文本（不参与图块生成） | 表格 = 文本类内容，统一走 text 阶段；组合内表格随 XML 段保留 |
| 9 | **grpSp 组合中若有 srcRect（图片裁剪显示区域）处理，须保留一致** | 组内图片的 srcRect 裁剪必须贯穿：元数据、资源、渲染、描述全程一致 |

**识别优先级**：`grpSp > Visio > 像素图（≥20% 页面）> 矢量（SVG/WMF）`——blocks 步骤**只产确定性图块**（组合/Visio/像素图/矢量），无自由聚类；**shape/connector 不单独成块**（文本框由 text 阶段提取，连接符作块内附属或忽略）；**表格移交 text 阶段**；公式不参与块；**首页/尾页整页跳过图块产出**。

### 0.2 图块解构与解读规则（10–16）

| # | 规则 | 一句话解读 |
|---|---|---|
| 10 | **grpSp 组合图块 → 保留整个组合的 XML 代码段**，对每段做**页面序号标记**；**每个 grpSp XML 段形成一个独立的 `*.xml` 文件**，存 `sources/` | 组合的源语言表达（结构/文字/位置/填充全保留），每组合一文件 = 块的"可解读源" |
| 11 | **grpSp 块 caption 用 DeepSeek-V4-Flash 解读**（直接输入 XML 代码段）；组内**公式对象由 DeepSeek 解读成 LaTeX 代码**，融入 caption 结果 | 文本模型读 XML，公式转 LaTeX，语义更准 |
| 12 | **grpSp 块解构时调用 PowerPoint 渲染成 PNG** 存 `images/`；但 caption 解读**不把 PNG 输入 qwen3.7-plus** | 渲染产物 ≠ 解读输入；grpSp 块解读走 DeepSeek XML 通道，VLM 不参与 |
| 13 | **OLE Visio / vsdx 对象，能剥离成 vsdx 文档的 → `.vsdx` 文件存 `sources/`，并渲染 PNG 存 `images/`** | 可编辑原生源 + 可视渲染双产物 |
| 14 | **OLE Visio / vsdx 对象，不能剥离成 vsdx 的 → 存成 XML 代码段**，处理方法与 grpSp 一致（页标记 + 独立 XML 文件 + DeepSeek 解读） | 退化路径：XML 段兜底 |
| 15 | **其他 SVG、WMF 等矢量图对象 → 与 Visio 相同处理**：尽量用 XML 代码段表达，不行才处理成 PNG | 矢量优先 XML，栅格化兜底 |
| 16 | **只有无法用 DeepSeek-V4-Flash 解读 XML 代码段的图块**（如纯像素图块），**才将像素图/PNG 渲染结果送 qwen3.7-plus 做 caption** | 解读模型按块路由：DeepSeek-XML 优先，qwen-VLM 兜底 |

**解读通道路由**：`grpSp / Visio-XML / SVG-WMF 矢量 → DeepSeek-V4-Flash（读 XML 段，公式转 LaTeX）`；`纯像素图 / 无 XML 可读的块 → qwen3.7-plus（读 PNG）`。

> **解读输入约定（v2.0 定稿）**：AI 解读**只按顺序读取 `sources/` 目录下的文件**（文件名排序，按扩展名路由模型）；`images/` 目录的文件**仅供人阅览解读效果**，不进入解读管线。

---

## 1. 背景与失败根因（通用，代码级）

图块提取失败的根因不是聚类算法本身，而是上游数据层缺陷（以下为**通用代码问题**，任何含 grpSp 组合或叠图文字的课件都会触发）：

### 根因 A：grpSp 组合被暴力拆散、组内坐标未变换 → 图块错位、残缺

三个遍历器都用 `root.iter()` 递归把 grpSp 内元素当独立对象，但坐标解析**只取元素自身 `<a:off>`，未叠加组合偏移**：

| 位置 | 现状代码 | 缺陷 |
|---|---|---|
| `extract_pptx_images.py` `_find_xfrm` | `elem.find(".//a:off")` 直接取 | 组内元素拿到的是**相对坐标**（未叠加组合 off） |
| `extract_pptx_images.py` `_xfrm_xy` | 同上 | 同上 |
| `extract_pptx_images.py` `iter_pictures` | `root.iter(PIC)` 递归 | 组内 `p:pic` 被当成独立顶层图片，坐标错位 |
| `extract_pptx_images.py` `iter_native_shapes` | `root.iter(SP)` 递归 | 组内 `p:sp` 被当成独立形状，坐标错位 |
| `extract_texts.py` `_iter_slide_texts` | `root.iter(SP)` 递归 | 组内文字被当成页面正文，坐标同样错位 |
| `_collect_atomic_objects` | 只处理 pic/sp/cxnSp/table，**不识别 `<p:grpSp>`** | 组合本身不产生原子对象，成员失去归属 |

**后果**：组内元素被定位到页面左上角（相对坐标区域），与真实位置相差一个组合偏移；组内底图与页面其他元素重叠 → 区域生长聚类把错误元素聚到一起或拆散，图块必然"无效"。

### 根因 B：组内文字被当成「页面正文纯文本」

`extract_texts._iter_slide_texts` 递归进 grpSp，把图块标签文字作为普通文本对象提取为 `TXT###` 条目 → 进入 `textbook.md` 正文。它们从未与所属图块建立关联。

### 根因 C：文本墙规则误伤「叠加在图形上的文字」→ 图块内容整块丢失

`visual_blocks.extract_blocks` 用 `max_shape_text=10` 把超过 10 字的文本框**整体剔除**。而大量课件中"大文本框 + 底图区域重叠"是常见构图（大段图块说明文字叠在示意图上）→ 核心内容全部丢失。

### 根因 D：cli_blocks 自举失败 + 过期提示

`cli_blocks` 缺 `atomic_objects.json` 时仍提示"请先运行 pptx-img"——`pptx-img` 已下线，提示过期且自举路径未生效。

### 根因 E：旋转组合的坐标变换

组合可带 `rot`（旋转）。逐元素做绝对坐标变换需要旋转矩阵，复杂且易错 → 定稿规则以「组合整体成块（bbox 用组合外接框）」+「XML 段保留原生结构」直接规避。

### 根因 F：Visio 记录不生成原子对象（直接影响规则 3/13/14）

`_collect_atomic_objects` 只有 `picture / chart / formula_ole / formula_omath` 分支，**没有 `visio` 分支**——而 `_emit_visio` 产出的记录 `kind="visio"`。后果：**Visio 对象在原子对象阶段被静默丢弃，当前版本 Visio 不进任何块**。规则 3/13/14 落地前必须补上。

---

## 2. 规则详解与指导意见

### 2.1 识别规则（1–9）边界情况

| 场景 | 指导意见 |
|---|---|
| **嵌套组合**（规则 2） | 嵌套 grpSp 递归展开为外层 children，**children 保留分层结构**（可支撑块内拓扑），但顶层只产出最外层组合的块 |
| **组合旋转/缩放/翻转** | 组合的 `off/ext` 已是旋转后的外接框 → 块 bbox 直接用，**不做逐元素变换**；children 存 `rel_bbox` 仅供审计/拓扑；原生结构由 XML 段（规则 10）承载 |
| **组内 Visio/像素图**（规则 1 统领 3/4） | 组合内的 Visio/像素图是块内容：`output_file` 进 children，渲染图天然含它们，不再单独成块 |
| **重叠文字并入**（规则 4 的补丁，关键） | 非组合像素图"单独成块"不等于"文字丢弃"：**文本中心落在像素图 bbox 内 → 并入该 raster 块作为标签**；仅当文本与任何图形都不重叠时才作为自由文本处理 |
| **公式归属**（规则 5，四环节接力） | 组合内公式**不进入 formula 步骤的输出**，链路：**① blocks** 随组合保留原始 `<m:oMath>` XML（在 grpSp XML 段内）→ **③ formula** 识别祖先 grpSp 标记 `in_group`、**从 formulas.md 排除** → **④ caption** DeepSeek 读 XML 段时识别公式对象转 LaTeX 融入块 caption → **⑥ author** 扩写消费块解读（含公式 LaTeX）。归宿唯一：组合内 → 块 caption（LaTeX）；组合外 → formulas.md |
| **Visio 现状缺陷** | `_collect_atomic_objects` 必须新增 `visio` 分支（根因 F），否则规则 3/13/14 无法实现 |
| **首页/尾页**（规则 6） | 第 1 页与最后一页整页跳过图块产出（文本/公式仍提取）；判断默认按页序，增强可选文本启发（尾页含"谢谢/致谢/Thanks"）兜底；`skip_cover_pages` 可关或 `--skip-pages` 显式指定 |
| **小像素图**（规则 7） | 非组合 raster 面积 < 20% 页面 → 丢弃；**被舍弃小图上的文本不并入**（图已弃，文本由 text 阶段承载）；组合内小图不受限（规则 1 统领） |
| **表格**（规则 8，移交 text 阶段） | 表格**不参与块生成**：text 步骤按表格读取，输出 **Markdown 表格**（表头+分隔行+数据行）入 texts.md / text_entries.json（`kind="table"`）；组合内表格随 XML 段保留，文本仍入 texts.md |
| **普通 shape/connector（非组合）** | **不再聚块**：文本框（shape）内容由 text 阶段提取；连接符（箭头）若属图块（在组合内）随块保留，否则忽略；两者都不单独成块 |
| **srcRect 一致性**（规则 9） | 三环节：① 元数据：children 携带 `src_rect`；② 资源：带 srcRect 的图裁剪落盘（by_page PNG / sources/ 复制），保证导出图 = 页面所见；③ 描述：语义描述附加 srcRect 标注。分工：渲染图 = 整页渲染 + bbox 裁剪（天然含裁剪效果）；资源图 = 按 srcRect 裁剪落盘；manifest 记录 `src_rect`/`cropped` |

### 2.2 解构与解读规则（10–16）详解

| 场景 | 指导意见 |
|---|---|
| **XML 独立文件**（规则 10） | **每个 grpSp 段一个独立 `*.xml` 文件**存 `sources/`，命名 `slide_{页:02d}_{块id}_grp.xml`（如 `slide_07_blk_01_grp.xml`）；文件内首行注释 `<!-- 第 N 页 grpSp: <组合名> (id=...) -->` 标记页面序号；段本体 = `<p:grpSp>…</p:grpSp>` 原始 XML（含组内全部元素/文字/公式/填充）；一组合一文件，命名即来源，可独立取用 |
| **DeepSeek 读 XML**（规则 11） | caption 输入 = XML 段 + 页面文本上下文，模型 = DeepSeek-V4-Flash（语义模型）；组内 `<m:oMath>` 等公式对象由 DeepSeek 直接转 LaTeX 代码，融入 caption（可结合本地 `extract_latex` 做 OMML→LaTeX 预转换作为参照） |
| **PowerPoint 渲染**（规则 12） | grpSp 块解构时**调用 PowerPoint COM 渲染 PNG** 存 `images/`（本机 Office 场景，pywin32 + ExportAsFixedFormat PDF + PyMuPDF 逐页）；**该 PNG 不送 qwen3.7-plus 解读**。**本项目不再依赖 LibreOffice**（用户本机有 Microsoft Office 即可）。 |
| **Visio 可剥离**（规则 13） | `_emit_visio` 已按容器类型落盘 `.vsdx`（zip）/`.vsd`（OLE2）→ 存 `sources/`；渲染 PNG 存 `images/`；caption 若可剥离 XML 内部结构（visio/document.xml）也可走 DeepSeek |
| **Visio 不可剥离**（规则 14） | 无有效容器的 Visio OLE → 存 XML 代码段（`<p:oleObj>` 段 + 关系信息），页标记 + 独立 XML 文件，处理与 grpSp 一致（DeepSeek 解读） |
| **SVG/WMF 等矢量**（规则 15） | 与 Visio 相同处理：优先 XML 代码段（`<a:blip>`/`<p:pic>` 段 + 矢量源内嵌结构）表达；不行才渲染 PNG |
| **qwen VLM 兜底**（规则 16） | **仅当块无 XML 可读**（纯像素图块、无法剥离的位图块等）→ 将像素图/PNG 送 qwen3.7-plus（VLM）做 caption；识别：块内无 grpSp/矢量/OLE 结构成员 |

**模型路由总结**：

```
块类型                      解读通道                     渲染产物
─────────────────────────────────────────────────────────────────
grpSp 组合块      → DeepSeek-V4-Flash 读 XML 段   → PowerPoint 渲染 PNG（images/）
                  （组内公式转 LaTeX 融入 caption）  + 独立 XML 文件（sources/，每组合一个）
Visio/vsdx 块     → 可剥离：.vsdx（sources/）+ PNG（images/）
                  → 不可剥离：XML 段，同 grpSp
SVG/WMF 矢量块    → 尽量 XML 段（sources/），不行才 PNG
纯像素图块        → qwen3.7-plus 读像素图原图（兜底）→ PNG（images/，仅供阅览）
表格（非块）      → text 阶段 → Markdown 表格（texts.md）  （不参与块/解读/渲染）
```

---

## 3. 规则驱动的修改方案

### 3.0 管线流程（v2.0 定稿，8 环节）

> 与 v1.5.1 的差异：**blocks（图块提取）提到最前**，text/formula 紧随其后供解读取上下文；**AI 解读独立成环节**且只读 `sources/`；**author 扩写充分消费 文本+公式+块解读**；binding 与输出收尾。

| # | 环节 | 职能 | 输入 → 产出 | 关键约定 |
|---|---|---|---|---|
| ① | **blocks**（图块提取） | 确定性拆块 + 解构 | pptx → `sources/`（XML 段 / .vsdx / SVG-WMF / **像素图原图**）+ `images/`（渲染 PNG）+ 块清单 | 规则 1-9；**单阶段确定性拆块**（grpSp/Visio/像素图/矢量，无自由聚类）；像素图块资源入 sources/ |
| ② | **text**（文本提取） | 页面正文 + **表格** | pptx → texts.md / text_entries.json | 组内文字标 `[图块内文本]`（in_group）；**表格在此按 Markdown 表格输出（规则 8，不参与块）** |
| ③ | **formula**（公式提取） | **仅非组合公式** | pptx → formulas.md | **组合内公式识别 `in_group` 后排除**（不写 formulas.md），由 ① 保留 oMath XML、④ DeepSeek 转 LaTeX |
| ④ | **caption**（图块 AI 解读） | 逐块解读 | **`sources/` 目录按文件名顺序** → captions.md | DeepSeek 读 XML（组内公式转 LaTeX）/ qwen 兜底读像素图；**images/ 仅供阅览不解读** |
| ⑤ | **related**（相关性过滤） | 删装饰块 | captions + 块清单 → related_filter.json | 审计计数 |
| ⑥ | **author**（教材扩写） | 教材级文案 | **文本 + 公式 + 图块解读** → textbook.md | 充分利用 ①②③④ 产物 |
| ⑦ | **binding**（图文绑定） | 块↔文本关联 | 块 + 文本 + 解读 → visualBlock_text_binding.json | relation_type / semantic_link |
| ⑧ | **输出**（组装） | 知识库成型 | 归位 sources/images + visual_blocks.json + binding + textbook/captions | **成功后清理全部过程文件** |

**文件生命周期（关键）**：
- **流程全部完成** → `_organize` 归位最终交付物（sources/ images/ 5 个 JSON/MD），**删除全部过程中间文件**（by_page/、atomic_objects.json、manifest、state 中间态等）。
- **中途中断（失败/手动停止）** → **不删除**过程文件，保留 `state.json` 断点。
- **再次运行** → `state.json` 断点续传**自动接续**未完成环节（已具备，保持并纳入新流程顺序）。

### 3.1 原子对象层：group / visio / raster / formula 归属

**位置**：`extract_pptx_images.py` `_collect_atomic_objects` + 新增 `iter_groups`

| 对象 | 规则 | 原子对象产出 |
|---|---|---|
| 顶层 grpSp | 1 | `kind="group"`，bbox=组合绝对 bbox，`text`=组内文本合并，`children`=组内全部元素（递归，含 rel_bbox / output_file / src_rect / 嵌套结构） |
| 嵌套 grpSp | 2 | 仅作为外层 children 出现，不单独产原子对象 |
| 非组合 Visio OLE | 3/13/14 | `kind="visio"`，output_file=`.vsdx/.vsd`（**新增分支**，修根因 F）；可剥离/不可剥离标志 |
| 非组合像素图 | 4+7 | `kind="raster"`；阶段 A 按 `面积 ≥ 20% 页面` 门槛筛后独立成块（<20% 丢弃） |
| 组合内公式 | 5/11 | 并入 group children + group text（DeepSeek 转 LaTeX 供 caption） |
| 非组合公式 | 5 | `kind="formula"`（现有行为保留，不参与块） |
| 非组合 shape/connector | — | **不再产块**：文本框内容由 text 阶段提取；连接符（箭头）若在组合内随块保留，否则忽略 |
| 表格 | 8 | **移交 text 阶段**：不产原子对象块；texts.md 输出 Markdown 表格（表头+分隔行+数据行） |
| 首页/尾页全部对象 | 6 | 该页不产任何块（文本/公式仍提取） |

新增字段（全部可选，向后兼容旧 atomic_objects.json）：`kind="group"`、`children[]`、`rel_bbox`、`src_rect`、`in_group`（文本标记）。

### 3.2 块生成：单阶段确定性拆块（无自由聚类）

**位置**：`visual_blocks.extract_blocks` 每页处理

- **确定性拆块**：按优先级逐一处理——`grpSp → group 块`、`visio → visio 块`、`raster（面积 ≥ 20% 页面）→ raster 块（重叠文本并入；<20% 舍弃，规则 7）`、`SVG/WMF 矢量 → 矢量块`；每个对象独立成块，**不做区域生长/自由聚类**。
- **不再处理**：非组合 shape（文本框 → text 阶段）、connector（非组合忽略）、表格（→ text 阶段）；`_region_grow` 自由聚类逻辑**删除**。
- **页级跳过（规则 6）**：`skip_cover_pages=True` 时，第 1 页与最后一页整页不产块。
- **资源归位（v2.0 定稿）**：所有块的**可复用资源全部入 `sources/`**——grpSp/Visio/SVG-WMF 的 XML 段与矢量源、**像素图块的原图**（`sources/slide_{页}_{块id}.png`）；`images/` 只放渲染/阅览用 PNG（PowerPoint 渲染，供人查看解读效果）。
- 块 ID 顺序 = 按 z_index（页内稳定输出）。

### 3.3 XML 段导出（新，规则 10/14/15）

**位置**：`cli_blocks`（新增导出函数，`_export_block_resources_and_binding` 同层）

1. `iter_groups` 遍历时缓存每个 grpSp 的**原始 XML 段**（`etree.tostring(grpSp_el)`，含命名空间）。
2. **每组合一个独立文件**写 `sources/slide_{页:02d}_{块id}_grp.xml`：文件首行 `<!-- 第 N 页 grpSp: <组合名> (id=...) -->` 注释标记页面序号，其后为 `<p:grpSp>…</p:grpSp>` 原始 XML 段（含命名空间声明）。
3. Visio 不可剥离段（规则 14）、SVG/WMF XML 段（规则 15）各自独立成文件（命名 `slide_{页:02d}_{块id}_ole.xml` / `_vec.xml`），处理与 grpSp 一致。
4. 段与块关联：`sources/` 文件名 ↔ `block_id`（页+组合 id 双向映射，记录到块 dict 的 `xml_source` 字段），供 caption 取用与 binding 关联。
5. **资源图片落盘（rldimg）**：解析 XML 段内 `r:embed`/`r:link`（rIdX）→ 读对应 slide 的 rels 映射到 `ppt/media/` → 复制到 `sources/rldimg/slide_{页}_{块id}_{原媒体名}`；块 `assets.rldimg` 记录路径，binding summary 记 `rldimg_total`。

### 3.4 渲染通道：PowerPoint 优先（规则 12/13）

**位置**：`extract_pptx_images.py` 渲染上下文 / `cli_blocks` 块渲染

1. **grpSp 块 PNG**：唯一渲染通道 = **PowerPoint COM**：`Open(ReadOnly=True, WithWindow=False)` → `ExportAsFixedFormat(pdf, 2=PDF, ...)` 完整参数 → PyMuPDF 逐页转 PNG（pdftoppm 兜底）→ bbox 裁剪；存 `images/`。**不再回退 LibreOffice**。
2. **Visio 块 PNG**：`.vsdx` 落盘后渲染 PNG（同通道）。
3. 该 PNG **仅供 images/ 交付与人工查看**，不进入 grpSp/矢量块的 caption 解读输入（规则 12/16）。

### 3.5 caption 解读路由（规则 11/16，独立环节④）

**位置**：`cli_blocks` `describe_block` 调用处 + `cli_paser` blocks 步骤模型选择；v2.0 将解读拆为**独立步骤**（流程④，见 §3.0）

1. **输入 = `sources/` 目录顺序**：解读器遍历 `sources/` 下文件（按文件名排序），按扩展名路由模型——`.xml`（grpSp/Visio/SVG-WMF 段）→ DeepSeek；`.png/.jpg`（像素图块原图）→ qwen；`.vsdx/.vsd` → 解析内部 XML 或渲染后 qwen；**`images/` 目录不参与解读，仅供人阅览**。
2. **DeepSeek 通道（默认）**：输入 = XML 段 + 页面文本上下文（text 步骤产物）；提示词要求：组内公式 → LaTeX 代码融入；输出块类型 + 语义描述。
3. **qwen VLM 通道（兜底）**：仅像素图块（sources/ 下无 XML 可读）→ 才调 **qwen3.7-plus** 读该图（现有 `describe_block` 保留给此通道）。
4. 语义增强（输出环节）维持 DeepSeek；`PPTX_PASER_NO_VLM=1` 逃生通道语义更新：跳过 qwen VLM（DeepSeek XML 通道不受影响）。

### 3.6 文本归属分流

**位置**：`extract_texts.py` `_iter_slide_texts` + `_sp_meta`

1. 文本对象标记 `in_group`（识别祖先 grpSp，记录所属组合名/id）。
2. **默认（filter 模式）**：组内文本不作为页面正文进入 textbook 合并，标注为 `[图块内文本]`（texts.md 保留可见，text_entries.json 增加 `in_group` 字段）。
3. 组内文字归宿 = 块 `internal_structure.nodes[].text` / caption / XML 段（规则 10 后文本天然在段内）。
4. `--no-filter` 全量模式仍输出全部条目。

### 3.7 BLOCK_RULES 配置

```python
BLOCK_RULES = {
    # —— 识别层（1-9）——
    "group_as_block": True,          # 1: grpSp 整体成块
    "nested_group_merge": True,      # 2: 嵌套组合并入外层
    "visio_standalone": True,        # 3: 非组合 Visio 独立成块
    "raster_standalone": True,       # 4: 非组合像素图独立成块
    "raster_overlap_text": True,     # 4补: 重叠文本并入 raster 块
    "formula_in_group": True,        # 5: 组内公式并入块内容
    "skip_cover_pages": True,        # 6: 首页/尾页跳过图块产出
    "raster_min_area_ratio": 0.20,   # 7: 非组合像素图面积 < 页面 20% 丢弃
    "table_markdown": True,          # 8: 表格在 text 阶段输出 Markdown 文本（不参与块）
    "src_rect_consistent": True,     # 9: srcRect 裁剪一致性
    # —— 解构/解读层（10-16）——
    "grpSp_xml_export": True,        # 10: grpSp XML 段导出（每组合一个独立 XML 文件到 sources/）
    "caption_deepseek_xml": True,    # 11: grpSp/矢量块 caption 用 DeepSeek 读 XML（公式转 LaTeX）
    "render_powerpoint": True,       # 12: grpSp/Visio 块用 PowerPoint 渲染 PNG（仅 PowerPoint，不再回退 LibreOffice）
    "visio_extract_vsdx": True,      # 13: Visio 可剥离 → .vsdx 存 sources/ + PNG 存 images/
    "visio_xml_fallback": True,      # 14: Visio 不可剥离 → XML 段，同 grpSp
    "vector_xml_first": True,        # 15: SVG/WMF 优先 XML 段，不行才 PNG
    "vlm_fallback_only": True,       # 16: qwen VLM 仅兜底（XML 不可解读的块才走）
    # —— 流程/生命周期（v2.0 定稿）——
    "raster_to_sources": True,       # 像素图块原图入 sources/（images/ 仅渲染阅览图）
    "caption_read_sources_only": True,  # 解读只按顺序读 sources/ 文件；images/ 不解读
    "cleanup_on_success": True,      # 全流程成功后删除全部过程文件（只留交付物）
    "resume_on_interrupt": True,     # 中断保留过程文件 + state.json 断点；再跑自动接续
}
```

### 3.8 顺手修复

- `cli_blocks` 缺 atomic_objects.json 的提示改为"自动准备原子对象"，去掉已下线的 `pptx-img` 文案（修根因 D）。
- `_collect_atomic_objects` 新增 `visio` 分支（修根因 F）。
- `PPTX_PASER_NO_VLM` 逃生通道语义更新（只跳过 qwen VLM，不跳过 DeepSeek XML）。

---

## 4. 实施步骤（建议顺序，对齐 v2.0 流程 ①–⑧）

| 步骤 | 内容 | 涉及文件 | 验收 |
|---|---|---|---|
| 1 | 新增 `iter_groups`（递归枚举 grpSp：name/id/off/ext/chOff/chExt/rot/原始 XML 段） | `extract_pptx_images.py` | 单测：枚举出各 grpSp，XML 段完整 |
| 2 | `_collect_atomic_objects`：group 原子对象 + children 收编 + **visio 分支**（可剥离/不可剥离标志） | `extract_pptx_images.py` | group 对象 bbox 正确、children 含文字/底图/src_rect；visio 不再被丢弃 |
| 3 | **流程① blocks**：**单阶段确定性拆块**（grpSp/visio/raster/矢量，**删除自由聚类**）+ **资源归位（XML 段/矢量源/像素图原图 → sources/；渲染图 → images/）** | `visual_blocks.py` / `cli_blocks.py` | 组页各 1 个完整块；sources/ 含块全部可复用资源；无自由聚类产物 |
| 4 | **流程②③ text/formula**：`in_group` 标记 + 组内文本默认不进正文 + **表格在 text 阶段输出 Markdown**；formula 排除组内公式（in_group） | `extract_texts.py` | texts.md 无图块内文字正文条目、含 Markdown 表格；formulas.md 无组内公式；entries 有 in_group |
| 5 | **流程④ caption**：解读器**按 sources/ 文件名顺序**遍历，路由 DeepSeek（XML，公式转 LaTeX）/ qwen（像素图兜底）；images/ 不解读 | `cli_blocks.py` / `cli_paser.py` | grpSp 块 caption 由 DeepSeek 生成且公式为 LaTeX；纯像素块走 qwen；sources/ 顺序遍历可审计 |
| 6 | **流程⑤⑥⑦ related/author/binding**：过滤 → 教材扩写（消费文本+公式+块解读）→ 图文绑定 | `cli_related.py` / `cli_author.py` / `cli_bind.py` | textbook 充分含块解读信息；binding 关联正确 |
| 7 | **流程⑧ 输出 + 生命周期**：归位交付物；**成功 → 删全部过程文件**；中断 → 保留 + state.json 续传（含新流程顺序） | `cli_paser.py` | 成功后无 by_page/atomic_objects 等过程文件；中断后重跑自动接续 |
| 8 | srcRect 一致性（规则 9）+ PowerPoint 渲染（规则 12/13） | `extract_pptx_images.py` / `cli_blocks.py` | 带 srcRect 图一致；images/ 渲染 PNG 正常 |
| 9 | 回归：无组合课件 + 规则开关测试 | 全链路 | 见 §5 |

## 5. 验收标准（按定稿规则断言）

1. **识别（1-9）**：grpSp 页每页 1 个完整 group 块（bbox 非 (0,0)）、组内文字齐全；首页/尾页无块；非组合小图（<20%）不进块；**表格在 text 阶段为 Markdown 表格（无表格块）**；带 srcRect 图元数据/资源/描述一致。
2. **解构（10/12/13 + 像素图入 sources/）**：每个 grpSp 一个 `sources/slide_{页}_{块id}_grp.xml`（含页注释、XML 段完整）；**像素图块原图在 sources/、渲染图在 images/**；可剥离 Visio 的 `.vsdx` 在 sources/、PNG 在 images/。
3. **解读（11/16 + sources/ 顺序）**：caption 步骤**按 sources/ 文件名顺序**逐个解读；grpSp/矢量块由 DeepSeek-V4-Flash 生成、组内公式为 LaTeX；**qwen3.7-plus 只被像素图块调用**（调用日志可审计）；images/ 无解读调用。
4. **生命周期**：全流程成功 → 结果目录**无任何过程文件**（by_page/、atomic_objects.json、manifest 等已清理）；人为中断 → 过程文件保留；**再次运行 → 自动接续**未完成环节（state.json 驱动，按新流程顺序）。
5. **回归**：无组合课件在 `raster_standalone=False` 下块数与旧版一致；`=True` 时仅"图+文字聚合块"拆分为图块+文本块，语义不丢失。
6. **确定性**：同一课件两次解析，块集合（页/块数/bbox/成员/XML 段/sources 文件）完全一致。

## 6. 风险与兼容性

| 风险 | 说明 | 对策 |
|---|---|---|
| 嵌套组合深度 | 组合套组合 | `iter_groups` 递归 + children 分层；深度无上限 |
| raster 独立成块改变无组合课件行为 | 图片+标题类聚合块被拆分 | `raster_standalone` 开关 + 回归对比 |
| 旋转/缩放组合子元素精确坐标 | 逐元素变换复杂 | 组整体成块规避 + XML 段保留原生结构 |
| atomic_objects.json schema 变化 | 新增字段 | 全部可选；旧文件读取路径不变 |
| 组内文字转块内容后 textbook 缺字 | 正文减少 | 由 caption/XML 段承载，不丢失；`--no-filter` 可保留全量 |
| 墙例外过宽 | 正文误并入装饰图 | 例外仅限"文本中心落在图形 bbox 内" |
| 首页/尾页误判（规则 6） | 封面恰好有重要图、末页非致谢 | `skip_cover_pages` 可关；`--skip-pages` 覆盖；文本启发增强 |
| raster 20% 丢弃误伤（规则 7） | 有语义的小图被丢 | 组合即豁免；阈值可配置；丢弃写入审计 |
| srcRect 双重处理（规则 9） | 资源图与渲染图不一致 | 渲染图=整页渲染+bbox 裁剪；资源图=按 srcRect 裁剪；manifest 记标志 |
| PowerPoint 渲染依赖 Office | 无 Office 环境 | 已无 LibreOffice 兜底：用户本机必须安装 Microsoft Office；无 Office 时返回 None（优雅降级，不阻塞 XML/DeepSeek 通道） |
| DeepSeek XML 解读失败 | 段过大/异常结构 | 按块降级 qwen VLM（规则 16 兜底）；重试/截断策略 |
| XML 文件数量 | 组合多时 sources/ 文件数增加 | 每文件小（单组合段），命名即来源；与现有 `sources/` 矢量资源命名体系一致 |
| 成功清理误删交付物（生命周期） | cleanup 边界不清 | 只清理 `过程文件/` 与已知中间态（by_page/atomic_objects/manifest）；sources/、images/、5 个 JSON/MD 白名单保护 |
| 中断续传状态错乱 | 新流程顺序变更后旧 state 兼容 | state.json 步骤表按新流程 ①–⑧ 重写；旧 state 检测（doc_md5）不匹配则全量重跑 |

## 7. 预处理操作指南（供使用者/README）

| 目标 | 用户在 PPT 里的操作 |
|---|---|
| 一整片图成为一个图块 | 选中该图所有相关形状（图形/文字/箭头/小图标）→ **Ctrl+G 组合** |
| Visio 图独立成块 | 插入→对象→Visio（或粘贴为 Visio 对象），**不要**与其它形状组合 |
| 像素图独立成块 | 直接插入图片，**不要**与其它形状组合 |
| 公式是图的一部分 | 与图形一起组合（组合后公式并入块，DeepSeek 转 LaTeX） |
| 公式是正文行内 | 保持独立（不组合）→ 自动走 formulas.md |
| 嵌套组合 | 有意为之才用：嵌套组合是外层块的子结构；若想分开，移出外层组合 |
| 首页/尾页 | 无需操作：解析器默认跳过封面/封底的图块产出；确有要保留的图用 `--skip-pages` |
| 小像素图不想被舍弃 | 面积 < 20% 页面的位图默认丢弃；想保留就与相邻图形**组合** |
| 表格 | 无需操作：表格统一在文本提取阶段按表格读取，自动输出 Markdown 表格（不参与图块） |

> 预处理的本质 = 给解析器打"块边界"标注；**组合的数量 = 该页图块数量的上限基准**，可据此验收。

---

## 8. 实施状态（2026-08-24 落地进度，真机验证《战略管理逻辑体系》20 页）

### 8.1 已实现并真机验证 ✅

| 能力 | 验证结果（《战略管理逻辑体系》20 页） |
|---|---|
| **group 原子对象**（规则 1/2） | 7/8/9、16/17/18/19 各 1 个 group 对象，bbox=组合绝对外接框（非 (0,0)）；嵌套组合（第 7 页"组合 3"内嵌"group 2"）递归收编；children 含文字/底图/表格 |
| **visio 分支**（规则 3/13/14，修根因 F） | `_collect_atomic_objects` 新增 `kind="visio"` 分支（本样本无 Visio，代码就位） |
| **单阶段确定性拆块**（无自由聚类） | `_region_grow` 调用删除；每页 1 个 `group_diagram` 块，组内文字齐全（含第 19 页三大文本框）；shape/connector/表格不产块 |
| **首页/尾页跳过**（规则 6） | 第 1/20 页 0 块，其余组合页 1 块 |
| **raster ≥20% 门槛**（规则 7） | 页面小图标（23×21px）不进块 |
| **XML 段导出**（规则 10） | `sources/slide_{页}_{块id}_grp.xml` 7 个（页注释 + 原始 XML，含嵌套组合） |
| **in_group 文本标记**（规则 1/5） | texts.md 组内文字标 `[图块内文本]`（'战略哲学'/'三个基本问题'等），页面标题保持正文 |
| **表格 Markdown**（规则 8） | `_table_text` 输出 Markdown 表格（表头+分隔行+数据行），text 阶段独立表格块 |
| **formula 排除组内公式**（规则 5） | `extract_latex` 跳过祖先含 grpSp 的 OMML（本样本 0 公式） |
| **binding/资源默认导出** | `_export_block_resources_and_binding` 不再依赖 `--semantic-model`，默认产出（binding 7 条 / sources 7 个 / xml_sources_total=7） |
| **rldimg 资源图片落盘** | `sources/rldimg/` 存 grpSp XML 段引用的媒体资源（第 7/8/9、16/17/18 底图 + 19 页 4 张）；rldimg_total=10、sources_total=17 |
| **渲染图** | `images/slide_{页}_{块id}.png` 7 张（现有整页渲染+bbox 裁剪通道；PowerPoint COM 优先为待办） |

### 8.2 已实现 ✅（第二轮）

| 项 | 说明 |
|---|---|
| **caption 解读路由**（规则 11/16） | cli_blocks 新增 `--caption-sources`：按 sources/ 文件顺序解读——.xml → DeepSeek 读 XML（组内公式转 LaTeX）；.png 像素图原图 → qwen VLM；images/ 不解读。实测 7/7 覆盖（6 块 deepseek_xml + 1 块兜底） |
| **DeepSeek 长 XML 空响应对策** | 实测 API 对 >2000 字符 XML 静默返回空 → 本地压缩（结构标签统计+组内文字+公式标记）≤2000 字符 + 3 次重试；仍空 → qwen 读渲染图兜底 → 规则模板二级兜底 |
| **cli_paser 8 环节流程** | STEPS 改为 `blocks→text→formula→caption→related→author→blocks_json`；blocks 步骤强制 `--no-vlm`（结构）；新增 caption 独立步骤（`--caption-sources --semantic-model deepseek-v4-flash`） |
| **生命周期** | `_organize` 成功即**删除全部过程文件**（不保留 过程文件/），只留交付物；texts/formulas 归位结果目录；中断（异常路径不 organize）保留 work + state 续传 |
| **端到端验证** | cli_paser 全流程（--skip related,author）9 分钟跑通：20 页 7 块 → caption 7 块 DeepSeek 解读 → blocks_json enrich → 结果目录仅 images/sources/5 文件 |

### 8.3 已实现 ✅（第三轮 · 4 项收尾）

| 项 | 说明 |
|---|---|
| **PowerPoint COM 渲染优先**（规则 12） | `render_pptx_pages` 仅走 `_render_pptx_pages_com`：pywin32 调 PowerPoint.ExportAsFixedFormat（PDF 格式常量 2，完整 14 参数）→ PyMuPDF 逐页 PNG（pdftoppm 兜底）。**完全移除 LibreOffice/soffice 渲染路径**。验证：20 页 20.2s，第 7 页 group 块裁剪 1604×625 px，视觉保真 100%（战略哲学/商道/天道/人道）。 |
| **OLE 公式组内排除**（规则 5） | `iter_ole_formulas(skip_in_group=True)` 新增组内判定（parent map 判祖先 grpSp）；`extract_latex` OLE 分支传 True——Equation.DSMT4 在组合内不进 formulas.md（OMML 此前已排除） |
| **XML 段原生前缀**（规则 10 可读性） | `_extract_grp_segments` 从 slide 原文按深度计数提取 `<p:grpSp>…</p:grpSp>` 段（保留 p:/a:/r: 前缀，嵌套正确处理）；`_collect_atomic_objects` 的 xml_segment 优先原文段，ET 序列化（ns0）兜底 |
| **enrich 与 caption 的 block_type merge** | cli_blocks 新增 `--prev-blocks`（caption 步骤产物）：合并其 block_type/semantic_description（caption_source 非空）到当前块；`enrich_semantics` 输出保留 old_sd 的 block_type/vlm_caption/formula_latex/caption_source（merge 而非覆盖）；cli_paser blocks_json 步骤传 `--prev-blocks` |
| 验证状态 | 全部 py_compile 通过；`_extract_grp_segments`/`_collect_atomic_objects` 内存级验证正常（第 7 页提取 1 段、2 对象）；**端到端写盘验证受本会话沙箱写权限限制未完成**（建议本机终端执行验证命令） |

### 8.4 待办 ⏳

| 项 | 说明 |
|---|---|
| （4 项收尾已全部完成，见 §8.3） | — |
| 端到端回归验证 | 收尾改动（COM 渲染/OLE 排除/前缀/merge）需在本机终端或沙箱授权后跑 cli_paser 全流程回归确认 |

### 8.5 新增配置（BLOCK_RULES 扩展，代码已接）

```python
"skip_cover_pages": True,          # 规则6（已实现）
"raster_min_area_ratio": 0.20,     # 规则7（已实现）
"raster_to_sources": True,         # 像素图块原图入 sources/（assets 已接）
"caption_read_sources_only": True, # 解读只读 sources/（路由待联调）
"cleanup_on_success": True,        # 成功后清理（cli_paser 待接）
"resume_on_interrupt": True,       # 中断续传（cli_paser 待接）
```
