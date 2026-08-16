# PPTX 多模态知识库构建管线 · 六项需求修改方案与工作提示词

> 作者：吴振谦 · 日期：2026-08-16
> 适用范围：`pptx-wzq`（`src/pptx_wzq/`）六步管线 `pptx-img / pptx-formula / pptx-text / pptx-caption / pptx-author / pptx-bind` + 编排器 `pptx-paser`
> 目标：在不推翻现有架构（薄壳 CLI + 核心库 + 过滤后处理）的前提下，落地 6 项新需求。

> **✅ 实施状态（2026-08-16）：7 项任务已全部实施并通过验证。**
> 涉及文件：`extract_pptx_images.py`（vsdx/坐标）、`extract_texts.py`（文本ID+坐标）、
> `img_filter.py`（vsdx 放行）、`cli_img.py`（矢量规范化）、`cli_author.py`（300字直出）、
> `cli_bind.py`（绑定增强+resume）、`cli_caption.py`（resume）、`cli_related.py`（新增）、
> `cli_paser.py`（日志+断点续传）、`pyproject.toml`（pptx-related 入口）、`README.md`。
> 验证：真实课件 231 页文本坐标 916/918、图片 305 条全含坐标、vsdx/vsd 合成识别、
> 300 字直出 0 API、binding schema/relation 回退、related 删除+审计、state 状态机、
> paser `--dry-run` 续跑计划。待真机验证项见文末「风险与验证点」。

---

## 一、现状与差距对照

| # | 需求 | 现状 | 差距（要改什么） |
|---|---|---|---|
| 1 | vsdx 矢量图直接存 vsdx | `VECTOR={emf,wmf,svg}`，未识别 vsdx；OLE 嵌入的 Visio 对象未解包 | 新增 vsdx 类型识别 + 从 `ppt/embeddings/` 解包落盘 `.vsdx` |
| 2 | 其他矢量图转 svg/wmf | 矢量图仅保留原文件（wmf/emf/svg），或 `--rasterize-vector` 栅格化 PNG | 规范化：emf/wmf/svg → 统一 svg（失败回退 wmf），不再默认栅格化 |
| 3 | 页面文字 >300 字不扩写 | `cli_author` 无条件生成 ≥300 字文案 | author 增加"原文超 300 字→直出原文"分支，省 Token |
| 4 | 绑定文档含图片ID/文本ID/坐标/50字逻辑关系 | `binding.json` 仅 `{page,text,images[{file,caption}]}` | 重构 binding schema + 新增 DeepSeek 逻辑关系陈述 |
| 5 | 图文相关性过滤（剔除 logo/作者/单位等） | 无此环节，caption 全部保留 | 新增 `pptx-related` 步骤：DeepSeek 判相关，无关者删除图+解释 |
| 6 | 日志文件 + 断点续传（自动续跑） | 靠手动 `--skip`；caption 仅增量落盘 | 新增 `pipeline.log`/`state.json` + 启动自检续跑 |

---

## 二、逐条修改方案

### 需求 1 · vsdx 矢量图提取

**识别来源**（OOXML 中 Visio 对象的标准嵌入方式）：
- 幻灯片 XML 中的 `<p:graphicFrame>` 内 `<p:oleObj r:id="rIdN" progId="Visio.Drawing.15" />`（`.vsdx`）或 `progId="Visio.Drawing.11"`（旧 `.vsd`）；
- 真实数据在 `ppt/embeddings/oleObjectN.bin`（经 `ppt/_rels/slideN.xml.rels` 的 `rIdN` 定位）；
- 同一 OLE 对象通常伴随一张 `<p:pic>`（EMF 预览图，供人看/供 AI 解读）。

**落盘规则**：
1. 新增 `kind="visio"`，`original_format="vsdx"`（或 `.vsd`）；
2. 读取 `oleObjectN.bin` 前 4 字节判断容器：
   - 魔数 `PK\x03\x04`（zip）→ 含 `visio/document.xml` 则为 `.vsdx`，直接**改名落盘为 `<stem>_p<page>_visio<seq>.vsdx`**；
   - 魔数 `D0CF11E0`（OLE 复合文档）→ 旧 `.vsd`，落盘 `.vsd`（如需 `.vsdx` 可复用 `svg-to-visio-vsdx` 技能用 Visio COM 另存，本需求仅要求"存成 vsdx 文件"时优先原格式已是 vsdx 的情形）；
3. 其 EMF 预览图按需求 2 转 svg，作为该 visio 对象的"可读配图"（供 caption/教材使用），并在 `ImageRecord` 上以 `preview_file` 关联。

**验收**：`out/by_page/` 与 `manifest.json` 出现 `kind="visio"` 且 `output_file` 以 `.vsdx` 结尾的记录；EMF 预览图同时生成 svg。

### 需求 2 · 其他矢量图规范化 svg/wmf

- `VECTOR` 集合维持 `{emf, wmf, svg}` 语义（不含 vsdx，vsdx 走需求 1）；
- 提取后处理新增 `_normalize_vectors`：
  1. 对 `records` 中 `original_format in VECTOR` 且未被过滤的项，优先用 **LibreOffice `soffice --convert-to svg`**（或 Inkscape `--export-type=svg`）转 svg；
  2. 转换失败（无 soffice/inkscape 或转出为空）→ **回退保留 wmf 原文件**（`--vector-out {svg,wmf}` 可切换偏好）；
  3. 转换成功则更新 `rec.output_file` 为 `.svg` 并记 `converted_to_png=False`（仍是矢量）；
- 不再把曲线 WMF 默认栅格化为 PNG 作为"收录图"（`_process_vectors` 中曲线渲染分支改为"转 svg"；公式版 WMF 仍走 PowerPoint 渲染 + OCR，逻辑不变）。

**验收**：`images/`（或 `by_page/`）内矢量对象为 `.svg` 或 `.wmf`，无意外 PNG 化；`manifest.json` 的 `original_format` 与 `output_file` 后缀可溯源。

### 需求 3 · 页面文字 >300 字不扩写

- `cli_author.py` 在逐页生成前增加**字数预检**：取该页 `texts.md` 原始文本（含公式行），去除空白后 `len(text) > 300` → 该页**不调用 DeepSeek**，直接以原文作为该页教材文案，加一行标注 `> 本页原文已超 300 字，直接提取，未作扩写。`；
- 不足 300 字才走现有 DeepSeek 扩写逻辑（≥300 字教材口吻）；
- 阈值可配：`--no-expand-threshold 300`。

**验收**：含长文本页的 PPT，`_textbook.md` 中该页内容 == 原文（无扩写），且 Token 消耗下降；`pipeline.log` 记录"直出页数"。

### 需求 4 · 图文绑定文档增强

**新 binding.json schema**（向后兼容：保留 `summary` 与 `pages` 结构，扩展字段）：

```json
{
  "stem": "xxx",
  "pages": [
    {
      "page": 4,
      "texts": [
        {"text_id": "P4-T1", "text": "…", "x": 100, "y": 120, "w": 640, "h": 80}
      ],
      "images": [
        {
          "image_id": "IMG0012",
          "file": "slide_04_pic_05.png",
          "page": 4,
          "position": {"x": 96, "y": 220, "w": 448, "h": 300},
          "caption": "…(qwen 诠释，原文保留)…",
          "relation": "本图以……与正文……构成……（约 50 字逻辑关系，DeepSeek 生成）"
        }
      ],
      "text": "第 4 页文案全文…",
      "has_image": true
    }
  ],
  "summary": { "pages": 10, "images_total": 12, "pages_with_image": 8 }
}
```

**落地要点**：
1. **图片 ID**：沿用 caption 的 `IMGxxxx` 全局编号（`pptx-caption` 已生成），bind 时从 captions.md 回填；
2. **文本 ID**：`extract_texts.py` 为每条文本生成页内稳定 ID `P<page>-T<seq>`（见需求 4 前置改造），写入 `texts.md` 与 `text_entries.json`；
3. **坐标**：
   - 图片：`ImageRecord.x/y`（幻灯片左上，px）+ 新增 `shape_w/shape_h`（来自 `<a:xfrm><a:ext cx cy>`，EMU→px，即幻灯片上的实际显示宽高）；现有 `width/height` 是图片像素尺寸，二者并存，`position` 用 shape 宽高；
   - 文本：`extract_texts.py` 的 `_sp_meta` 增读 `<a:xfrm>` 的 `off/ ext`，产出 `x/y/w/h`；
4. **50 字逻辑关系**：`cli_bind.py` 新增 DeepSeek 调用（`deepseek-v4-flash`），输入 = 该图 `caption`（qwen 诠释）+ 该页 `texts` 原文，输出 ≤60 字的中文逻辑关系陈述，写入 `relation` 字段；失败则回退 `caption` 首句；
5. **分页定位**：bind 不再仅靠文件名 `slide_NN`，改以 `manifest.json` 的 `page + output_file` 为主键，兼容 vsdx/svg 等新类型。

### 需求 5 · 图文相关性过滤

- **新增子命令 `pptx-related`**（在 `pptx-caption` 之后、`pptx-author` 之前执行）：
  1. 读 `captions.md` 全部条目 + 每图对应页 `texts.md` 原文；
  2. 逐图调用 **DeepSeek（deepseek-v4-flash）** 判断"该图（据 qwen 诠释）与本页正文是否在知识/教学上相关"，输出结构化判定 `{keep: bool, reason: "相关/logo/作者信息/单位名称/项目类别/每页重复装饰/其他"}`；
  3. **判定为无关**（logo、作者署名、单位/项目名称、每页一致的装饰条、页脚二维码等）→ 从 `images/`（及 `by_page/`）、`captions.md` 中**删除该图及解释**，并记录到 `related_filter.json`（含被删原因，可审计、可恢复）；
  4. 输出过滤后的 `captions.md`（覆盖或另存 `_captions_kept.md`），并更新 `manifest.json` 标记 `related="drop"` 的条目；
- **判据提示词要点**（写入系统提示）：优先保留"讲解知识点、原理、电路、结构、数据、流程、示例"的图；判定为可删的典型信号：品牌 logo、作者/教师姓名、单位全称与校徽、课题/项目编号、每页重复出现的页眉页脚装饰、纯装饰分隔线、二维码、联系方式。
- 支持 `--keep-all`（跳过过滤，调试用）。

**验收**：同页 logo/作者/单位信息不再出现在最终 `images/`、`captions.md`、`binding.json` 中；`related_filter.json` 可查每张被删图的页号、文件与原因。

### 需求 6 · 日志文件 + 断点续传

**日志设计**（目标目录下，持久化）：

- `pipeline.log`（人类可读，追加写）：每步 `[时间] 步骤名 状态 耗时 产物`；
- `state.json`（机器可读，覆盖写）：流水线状态机，示例：

```json
{
  "pptx": "第九章功率放大电路.pptx",
  "stem": "第九章功率放大电路",
  "started_at": "2026-08-16T10:00:00",
  "steps": {
    "img":     {"status": "done", "finished_at": "…", "output": "过程文件/img"},
    "formula": {"status": "done", "finished_at": "…"},
    "text":    {"status": "done", "finished_at": "…"},
    "caption": {"status": "partial", "done_images": 23, "total_images": 40},
    "related": {"status": "pending"},
    "author":  {"status": "pending"},
    "bind":    {"status": "pending"}
  }
}
```

**续跑逻辑**（`cli_paser.py` 启动时）：
1. 读 `out/state.json` + `pipeline.log` + 检测 `out/过程文件/` 下已生成的产物；
2. `status=="done"` 且产物存在 → 自动跳过该步（等价于自动 `--skip`）；
3. `status=="partial"` → 从断点续跑：`caption` 按 `done_images` 续；`author` 按已完成页续（`_textbook.md` 已有页跳过）；`bind` 按已绑定页续；
4. `status=="failed"` 或产物缺失 → 重跑该步；
5. **三处来源冲突**时以 `state.json` 为准，日志/产物只做校验与告警（提示用户"检测到日志与产物不一致，将以 state 续跑"）；
6. 每步完成后立即回写 `state.json`（原子写：临时文件 + rename），确保任意时刻中断可恢复；
7. 新增 `--reset` 强制从头重跑（清空 state 与旧产物，但保留用户确认），`--dry-run` 仅打印续跑计划不执行。

**验收**：跑一半 Ctrl+C 后重跑同一命令，自动跳过已完成步骤、从 caption 断点续跑，最终产物与一次跑完一致。

---

## 三、数据契约变更汇总（关键）

### 3.1 `ImageRecord`（`extract_pptx_images.py`）新增字段

| 字段 | 类型 | 含义 |
|---|---|---|
| `shape_w` / `shape_h` | int | 幻灯片上的显示宽高（EMU→px，来自 `<a:ext>`） |
| `ole_progid` | str | OLE 对象 progId（识别 Visio 用，如 `Visio.Drawing.15`） |
| `preview_file` | str | visio 对象关联的 EMF/SVG 预览图文件名（可空） |

（`x/y` 已有，语义=幻灯片左上角 px，保持不变。）

### 3.2 文本条目（`extract_texts.py`）新增

- 每条 `(kind,name,ph_type,text)` → 增加 `(text_id, x, y, w, h)`：
  - `text_id = f"P{page}-T{seq}"`（页内递增）；
  - 坐标来自 `<a:xfrm><a:off x y>` 与 `<a:ext cx cy>`（EMU→px）；
- 输出：
  - `texts.md`：每行 `P4-T1 | sp | 标题 | 文本`（追加坐标列可读格式）；
  - `text_entries.json`：含 `text_id/x/y/w/h` 的完整审计。

### 3.3 新产物清单

```
结果目录/
├─ images/                    教学图片集（含 .svg/.wmf/.vsdx 矢量）
├─ vectors/                   vsdx 与 svg/wmf 矢量归档（可选，独立于 images）
├─ <名>_captions.md           图片 AI 解读（已过滤无关图）
├─ <名>_related_filter.json   相关性过滤审计（被删图 + 原因）
├─ <名>_textbook.md           教材文案（300字页直出）
├─ <名>_binding.json          图文绑定（image_id/text_id/坐标/relation）
├─ state.json                 断点续传状态机
├─ pipeline.log               运行日志
└─ 过程文件/                  img/formula/text/by_page/manifest…
```

---

## 四、完整工作提示词（可直接交付智能体执行）

> 将以下整段作为智能体（如 WorkBuddy / CodeBuddy / 通用编码 Agent）的 System/任务提示词。

```text
# 角色
你是「PPT 多模态教学知识库管线」的高级工程师与提示词专家。你必须在不推翻现有架构的前提下，
对项目 pptx-wzq（Windows，Python，薄壳 CLI + 核心库）实施 6 项改造。项目根目录即工作目录，
源码在 src/pptx_wzq/，核心文件：extract_pptx_images.py、extract_texts.py、img_filter.py、
cli_img.py、cli_formula.py、cli_text.py、cli_caption.py、cli_author.py、cli_bind.py、
cli_paser.py、cli_common.py。

# 总约束
1. 保持现有「薄壳 CLI 调用核心库」的分层，不重写核心库整体，只做最小范围修改；
2. 所有对外 CLI 参数向后兼容：旧参数不删除，仅新增参数/新命令；
3. 退出码沿用 cli_common 约定：0 成功 / 1 处理异常 / 2 参数或环境错误；
4. 输出编码统一 UTF-8；坐标统一「EMU→px」（1 inch = 914400 EMU = 96 px）；
5. 每个新功能必须可被单元级验证，给出验证命令与预期产物路径；
6. 中文注释与 docstring，风格与现有代码一致。

# 任务一：vsdx 矢量图提取（改 extract_pptx_images.py + cli_img.py）
- 新增 kind="visio"，original_format="vsdx"（旧 .vsd 保留 .vsd）；
- 识别 <p:oleObj progId="Visio.Drawing.*">，经 rels 定位 ppt/embeddings/oleObjectN.bin；
- 读 bin 前 4 字节：PK 开头（zip 且含 visio/document.xml）→ 落盘 <stem>_p<page>_visio<seq>.vsdx；
  D0CF11E0（OLE 复合）→ 落盘 .vsd；
- 该 OLE 的 EMF 预览图（<p:pic>）按任务二转 svg，并在 ImageRecord.preview_file 关联；
- manifest.json 记录 kind/original_format/output_file/preview_file/坐标。

# 任务二：其他矢量图规范化 svg/wmf（改 extract_pptx_images.py + cli_img.py）
- VECTOR 语义仍为 {emf,wmf,svg}（不含 vsdx）；
- 新增 _normalize_vectors：emf/wmf/svg 优先 soffice --convert-to svg（或 Inkscape --export-type=svg）
  转 svg；失败回退保留 wmf 原文件；新增参数 --vector-out {svg,wmf}；
- 将原「曲线 WMF→PowerPoint 渲染 PNG」的收录分支改为「转 svg」；公式版 WMF 仍走渲染+OCR 不变。

# 任务三：300 字不扩写（改 cli_author.py）
- 逐页生成前预检：该页 texts.md 原文去空白后 > 300 字 → 不调用 DeepSeek，原文直出，
  加标注「> 本页原文已超 300 字，直接提取，未作扩写。」；阈值参数 --no-expand-threshold 300；
- 不足 300 字才走现有扩写逻辑。

# 任务四：文本 ID + 坐标（改 extract_texts.py）
- 每条文本产出 text_id=f"P{page}-T{seq}"（页内递增）与坐标 x/y/w/h（来自 <a:xfrm> off/ext，EMU→px）；
- texts.md 每行携带 text_id 与坐标；text_entries.json 含 text_id/x/y/w/h 完整审计。

# 任务五：图文绑定增强（改 cli_bind.py，需调用 DeepSeek）
- 重构 binding.json 为：pages[].texts[{text_id,text,x,y,w,h}]、
  images[{image_id,file,page,position{x,y,w,h},caption,relation}]、text、has_image；
- image_id 回填 captions.md 的 IMGxxxx；position 用 ImageRecord 的 x/y + 新增 shape_w/shape_h；
- 主键由 manifest.json（page+output_file）定位，兼容 vsdx/svg；
- 每图调用 deepseek-v4-flash 生成 ≤60 字「图片与原文本逻辑关系」写入 relation，
  输入 = 该图 caption + 该页 texts 原文；失败回退 caption 首句；
- 保留 summary 与 pages 顶层结构，向后兼容。

# 任务六：图文相关性过滤（新增 cli_related.py + 接入 paser）
- 读 captions.md + 对应页 texts.md，逐图调用 deepseek-v4-flash 判相关，
  输出 {keep: bool, reason}；
- 判定无关（logo/作者信息/单位名称/项目类别/每页重复装饰/页脚二维码/联系方式）→
  从 images/、by_page/、captions.md 删除该图及解释，写 <名>_related_filter.json 审计；
- 优先保留讲知识/原理/电路/结构/数据/流程/示例的图；支持 --keep-all 跳过。

# 任务七：日志 + 断点续传（改 cli_paser.py）
- 目标目录写 pipeline.log（追加）与 state.json（覆盖，steps 状态机 status∈
  {pending,partial,done,failed} + 计数）；
- 启动自检：done 且产物存在→跳过；partial→续跑（caption 按 done_images、author 按完成页、
  bind 按完成页）；failed/缺产物→重跑；
- state.json 原子写（临时文件+rename）；新增 --reset（重头）与 --dry-run（打印续跑计划）；
- 日志与产物冲突时以 state.json 为准并告警。

# 执行顺序与验收
按 任务一→二→三→四→五→六→七 顺序实施；每完成一个任务运行一次现有冒烟测试，
并在 README.md 补充新命令/新产物说明。全部完成后用一份真实 .pptx（含 vsdx、长文本页、
logo/作者页）端到端验证 6 项需求，输出验证清单（命令 + 产物路径 + 是否符合预期）。
禁止范围蔓延：不改与 6 项需求无关的代码路径。
```

---

## 五、落地顺序建议（给项目负责人）

1. **先做任务四（文本 ID + 坐标）与 ImageRecord 扩展（shape_w/h）**——这是任务五/六的数据地基；
2. 再做任务一、二（vsdx / 矢量规范化，纯本地，无 Token）；
3. 再做任务三（300 字直出，改动最小、收益立现）；
4. 再做任务五、六（binding 增强 + 相关性过滤，均依赖 DeepSeek，需联调提示词）；
5. 最后做任务七（日志 + 断点续传，横切所有步骤）；
6. 全程用 `pptx-paser --dry-run` 观察续跑计划，用真实课件回归。

**风险与验证点**：
- vsdx 的实际容器格式需在真实样本上验证（zip 型 vsdx vs OLE 复合型 vsd 的魔数判断）；
- `soffice`/`inkscape` 是否安装决定矢量转 svg 的可用性，需保留 wmf 回退；
- DeepSeek 逻辑关系/相关性两处提示词需用 3~5 页样例校准措辞与判据强度，避免误删有效图。
