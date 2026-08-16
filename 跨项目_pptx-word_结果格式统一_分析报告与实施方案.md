# pptx-wzq × word-wzq 结果文档格式差异分析与统一实施方案

> 分析人：WorkBuddy Agent · 日期：2026-08-16
> 分析对象：`pptx-wzq v1.2.0`（binding v3，PPT 管线）与 `word-wzq v2.2.0`（Word 管线）
> 样本：pptx 用《战略管理》课件 23 页实测产物；word 用 tpack4_v3（TMACK 论文）实测产物

---

## 第一部分 · 分析报告

### 1. 产出文档总览

| 文档 | pptx-wzq | word-wzq | 一致性 |
|---|---|---|---|
| `<名>_texts.md` 文本清单 | `## 第 N 页` + 4 列表格 | `## 第 N 节` + 3 列表格 | 部分一致 |
| `<名>_captions.md` 图片解读 | `### IMGxxxx — \`file\` ✅` + 文本 | 相同格式 | ✅ 一致 |
| `<名>_textbook.md` 教材文案 | `## 第 N 页`，直出用 blockquote 标注 | `## 第 N 节`，直出用 `<!-- 原文直出 -->` | 部分一致 |
| `<名>_formulas.md` 公式清单 | 按页 `## 第 N 页` | 按节 `## 第 N 节` | 术语差异 |
| `<名>_binding.json` 图文绑定 | v3 结构（13 字段图片条目） | v2.2 结构（13 字段图片条目） | **字段集一致** |

> 关键结论：**binding.json 的图片条目字段集两边已完全一致**（file/caption/source/kind/image_id/page/paragraph/text_id/w/h/x/y/position/relation），
> 这是上一轮 pptx v3 对齐 word 结构的结果。剩余差异集中在「取值格式」「生成方式」「文档类型固有差异」三类。

### 2. binding.json 详细差异

| # | 维度 | word-wzq | pptx-wzq | 差异类型 | 说明 |
|---|---|---|---|---|---|
| A1 | 顶层/页结构 | `{stem, pages[{page,text,images,has_image}], summary}` | 相同 | ✅ 一致 | |
| A2 | 图片条目字段集 | 13 字段 | 13 字段（同名） | ✅ 一致 | |
| A3 | **text_id 格式** | `S{节:02d}_P{段:02d}`（如 `S03_P27`） | `TXT{页:03d}-{序:02d}`（如 `TXT017-01`） | ❌ 不一致 | word 内部还两套：texts.md 用 TXT###-##、binding 用 S##_P## |
| A4 | paragraph 语义 | Word 段落号（images_meta 的 par_index） | 页内文本条目序号 | ⚠️ 语义不同 | 均表示"图所在文本块序号" |
| A5 | x/y | `-1`（内联图流式无绝对坐标） | 真实幻灯片坐标（px） | ⚠️ 文档类型固有 | Word 流式排版无画布坐标 |
| A6 | w/h | 图片物理尺寸（emu/9525） | shape 显示尺寸（shape_w/shape_h） | ❌ 语义不同 | 应统一为物理像素 |
| A7 | position 生成 | 确定性位置 + DeepSeek `position_role`（≤40 字）拼接 | 一次调用 LLM 生成 `relation`+`position` 两个 JSON 字段 | ❌ 生成方式不同 | word 区分"功能角色/逻辑关系"两维度 |
| A8 | relation 字数 | ≤50 字 | ≤60 字 | ❌ 规格不同 | |
| A9 | kind 取值 | `svg / vsdx / raster` | `raster / vector / visio` | ❌ 枚举不同 | |
| A10 | source | sources/ 目录矢量源文件名（.vsdx/.svg/.wmf/.emf） | PPT 包内 media 路径 | ⚠️ 语义不同 | 均为"溯源来源"，可保留各自语义 |
| A11 | summary | `{pages, images_total, pages_with_image}` | 另有 `relations/positions` 计数 | ❌ word 缺 2 项 | |
| A12 | 图片命名/页码语义 | `block_NN_pic_MM` / 节 | `slide_NN_pic_MM` / 页 | ⚠️ 源文档固有 | 命名无法统一 |

### 3. md 文档详细差异

#### 3.1 texts.md
| # | word-wzq | pptx-wzq | 差异 |
|---|---|---|---|
| B1 | 分块标题 `## 第 N 节` | `## 第 N 页` | 术语（源文档结构：Word 按标题分节 / PPT 按幻灯片分页） |
| B2 | 表头 3 列 `\| ID \| 类型 \| 文本 \|` | 表头 4 列 `\| ID \| 类型 \| 文本 \| 坐标 \|` | 列数不同 |
| B3 | text_id `TXT###-##` | `TXT###-##` | ✅ 一致 |
| B4 | 类型：标题/内容/表格行 | 类型：标题/内容 | 基本一致 |

#### 3.2 captions.md —— ✅ 完全一致（条目格式 `### IMGxxxx — \`file\` ✅` + 尾部统计；仅生成工具名前缀不同，可忽略）

#### 3.3 textbook.md
| # | word-wzq | pptx-wzq | 差异 |
|---|---|---|---|
| D1 | 文档标题 `# xxx_textbook 教材文案` | 相同 | ✅ 一致 |
| D2 | 节标题 `## 第 N 节` | `## 第 N 页` | 术语 |
| D3 | 直出标记 `<!-- 原文直出 -->`（HTML 注释） | `> 本页原文已超过 300 字，直接提取，未作扩写。`（blockquote） | 标记形式不同 |
| D4 | 直出内容保留 `[标题]…\| TXT…\| 内容 \| …` 表格行 | 直出内容为纯文本 | 内容呈现不同 |
| D5 | 统计行：`共 N 节，其中 M 节原文直出，K 批模型生成` | `共 N 页，其中 M 页原文直出` | 基本一致 |

#### 3.4 formulas.md —— 仅分块标题术语差异（`第 N 节` vs `第 N 页`），条目格式一致

### 4. 差异归类与"不可统一项"判断

- **✅ 可统一（取值/格式层）**：A3 text_id、A6 w/h、A7 position 生成、A8 relation 字数、A9 kind 枚举、A11 summary、B2 表头列数、D3/D4 直出标记与内容、D5 统计行。
- **⚠️ 文档类型固有（建议保留语义，文档注明）**：A5 x/y（Word 流式无坐标）、A10 source（来源体系不同）、A12 图片命名与"页/节"术语（源文档结构不同）。强行统一会丢失信息或产生误导（如 Word 文档标"第 N 页"与真实页码不符）。
- **建议统一策略**：页/节术语保留各自输出（word 出"节"、pptx 出"页"），但**解析器两边都已兼容**（word 的 split_sections 兼容"第 N 页/节"；pptx 解析"第 N 页"——如需互读，给 pptx 的 `_split_pages` 补"节"兼容即可）。

---

## 第二部分 · 统一实施方案

### 5. 统一后目标格式（规范）

**binding.json 图片条目（两项目一致）**：
```json
{
  "file": "block_03_pic_01.png",
  "caption": "…qwen 诠释…",
  "source": "block_03_pic_01.vsdx",
  "kind": "vsdx",
  "image_id": "IMG0001",
  "page": 3,
  "paragraph": 27,
  "text_id": "TXT003-01",
  "w": 145, "h": 137,
  "x": -1, "y": -1,
  "position": "位于第 3 节（第 27 段），该图对该节文字表达起『直观呈现本节所述 TPACK 框架』的作用",
  "relation": "该图以示意图说明 TPACK 知识耦合关系，与本节呈说明关系"
}
```
- `text_id`：统一 `TXT{节/页:03d}-{条目:02d}`（与各自 texts.md 的 TXT id 对应）
- `kind`：统一输出 `original_format` 直出（vsdx/svg/wmf/emf/png/jpg…）
- `w/h`：统一图片物理像素尺寸
- `x/y`：pptx 输出真实坐标；word 输出 -1（流式，文档注明）
- `position`：统一「位于第 N 节/页（第 M 段）＋ DeepSeek position_role（≤40 字）」拼接结构
- `relation`：统一 ≤50 字（DeepSeek）
- `summary`：统一 `{pages, images_total, pages_with_image, relations, positions}`

**texts.md / textbook.md**：
- texts.md 统一 4 列 `| ID | 类型 | 文本 | 坐标 |`（word 坐标列输出 -1）
- textbook.md 直出标记统一 `> 原文直出（超过 N 字，未扩写）。`；直出内容统一纯文本
- 分块标题：保留各自术语（节/页），解析器补双向兼容

### 6. 最小改动方案（逐文件）

> 原则：**只改薄壳 CLI 的格式化/生成层，不动 extract 核心与数据结构**；
> 以 word 的 binding 结构（position_role 分离）为基准，因为其维度划分更符合需求定义。

#### 6.1 word-wzq（约 35 行改动）

| 文件 | 改动 | 行数 |
|---|---|---|
| `cli_bind.py` | ① `_enrich_relations` 中 `text_id` 由 `S##_P##` 改为 `TXT###-##`：从 `images_meta` 的 `par_index` 映射到 texts.md 的 TXT 条目号（需 `cli_img`/`extract_docx` 在 `images_meta.json` 记录 `txt_id` 字段，或在 bind 内解析 `texts.md` 建立 段落→TXT 映射）；② `kind` 改为 `original_format` 直出；③ `summary` 补 `relations/positions` 计数；④ relation 截断 ≤50 字 | ~25 |
| `cli_img.py`/`extract_docx.py` | `images_meta.json` 每图记录 `txt_id`（图片所在段落的 TXT 条目号） | ~5 |
| `cli_text.py` | texts.md 表头补第 4 列「坐标」，值输出 `-1` | ~3 |
| `cli_author.py` | 直出标记 `<!-- 原文直出 -->` → `> 原文直出（超过 N 字，未扩写）。`；直出内容去表格行格式改纯文本 | ~4 |

#### 6.2 pptx-wzq（约 35 行改动）

| 文件 | 改动 | 行数 |
|---|---|---|
| `cli_bind.py` | ① `LINK_SYSTEM` 改为 word 的 `POSITION_RELATION_PROMPT` 同款（`position_role` ≤40 字 + `relation` ≤50 字分离判断）；`position` 由"一次调用返回整句"改为"确定性位置前缀 + role 拼接"；② `w/h` 改用 manifest 的 `width/height`（物理像素）替代 `shape_w/shape_h`；③ `kind` 改为 `original_format` 直出（保留 raster/vector 归类可作附加字段）；④ relation 截断改 ≤50 字 | ~30 |
| `cli_paser.py` | `_split_pages`/相关解析兼容「第 N 节」（供互读 word 产物） | ~3 |
| `cli_author.py` | 直出标注文案微调对齐（`> 原文直出（超过 300 字，未扩写）。`） | ~2 |

#### 6.3 不修改项（保留差异并文档说明）
- x/y：word=-1（流式）、pptx=真实坐标 —— 语义不同，各保留
- source：word=矢量源文件名、pptx=包内 media 路径 —— 均为溯源字段，语义不同可保留
- 图片命名（block_NN vs slide_NN）与"节/页"术语 —— 源文档固有
- captions.md —— 已一致

### 7. 实施步骤与验证

1. **先改 word-wzq**（6.1）：text_id 映射是唯一涉及提取层的改动，优先做；
2. **再改 pptx-wzq**（6.2）：position/kind/w/h 均为 bind 薄壳格式化改动；
3. **回归验证**（各用一份真实样本）：
   - word：`tpack4_v3` 重新 bind，检查 text_id 为 `TXT###-##`、summary 含 relations/positions；
   - pptx：《战略管理》重新 bind，检查 position 为「确定性位置 + 『role』」结构、w/h 为物理像素、relation ≤50 字；
4. 交叉验证：pptx 的 `cli_paser` 能解析 word 产物目录（节兼容）；
5. 版本同步：两项目 bump（word 2.2.0→2.3.0、pptx 1.2.0→1.3.0），PyPI/GitHub 同步发布。

### 8. 改动量结论

| 项目 | 涉及文件 | 预估改动 | 风险 |
|---|---|---|---|
| word-wzq | cli_bind / cli_img / cli_text / cli_author | ~35 行 | text_id 映射需 images_meta 加字段（小） |
| pptx-wzq | cli_bind / cli_paser / cli_author | ~35 行 | position 结构变化影响消费者（若有）需同步 |
| 合计 | 7 个文件 | **~70 行** | 低（均为薄壳格式化层） |

**结论**：两项目 binding 字段集已一致，统一只需 ~70 行薄壳层改动，不动任何核心提取逻辑；
推荐以 **word 的 binding 结构（position_role 分离）+ 统一的 `TXT###-##` text_id + `original_format` 直出 kind + 物理像素 w/h** 为公共规范，两边各改一个 bind 文件为主。
