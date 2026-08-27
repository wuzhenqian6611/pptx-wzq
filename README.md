# pptx-wzq · PPT 多模态教学知识库自动化构建

> **作者：吴振谦 · wuzhenqian@nbu.edu.cn · QQ：38328063**
> 本项目帮助高校教师把教学 PPT 自动转化为**图文并茂、可检索、可复用、可再加工**的多模态教学知识库。
> 当前版本：**3.1.3**（PyPI：https://pypi.org/project/pptx-wzq/ · GitHub：https://github.com/wuzhenqian6611/pptx-wzq）

**开发原因**：高校教师在课程建设与教材建设中，长期面临「课件里的大量图片、公式、文本散落各处，难以整理为规范、可复用、图文并茂的教学资源」的痛点——手工整理一份课程的图文知识库往往要耗费数周。本工具把 PPT 自动转化为图文并茂的多模态教学知识库，让教师从机械整理中解放出来。

---

## 目录

- [一、它能做什么](#一它能做什么)
- [二、安装与环境要求](#二安装与环境要求)
- [三、快速开始](#三快速开始)
- [四、命令一览](#四命令一览)
- [五、输入 PPTX 预处理要求](#五输入-pptx-预处理要求)
- [六、工作流程（8 环节）](#六工作流程8-环节)
- [七、输出文档体系与内部格式](#七输出文档体系与内部格式)
- [八、技术架构](#八技术架构)
- [九、图块识别规则（16 条）](#九图块识别规则16-条)
- [十、生命周期与可靠性](#十生命周期与可靠性)
- [十一、常见问题](#十一常见问题)
- [十二、版本演进](#十二版本演进)
- [十三、文档与许可](#十三文档与许可)

---

## 一、它能做什么

输入一份教学 PPT（.pptx），自动完成 **8 环节流水线**：图块提取 → 文本提取 → 公式提取 → 图块 AI 解读 → 相关性过滤 → 教材文案（整篇自主分章分节）→ JSON 组装 → 输出。

**核心能力：**

- **图块（可视逻辑块）优先**：以「组合即图块」为第一原则，作者用 PPT 组合工具绘制的逻辑图/示意图整体保留（原生 XML 段），而非拆散成碎片。
- **源语言级解构**：grpSp 组合保留原生 XML 段（p:/a:/r: 前缀）；Visio/vsdx 剥离原生文件；公式转 LaTeX——信息不降维。
- **模型按块路由**：有 XML 的块 → DeepSeek 读 XML（组内公式转 LaTeX）；纯像素图块 → qwen VLM 兜底；DeepSeek 空响应自动降级；**blocks_json 步骤 qwen 视觉兜底按需触发**（仅 caption 未解读的块读渲染图，避免重复解读）。
- **公式双通道**：组合内公式既保留为组合块内容（XML 段），又进 formulas.md（语义存档）；OLE 预览图（公式/Visio 快照）判定后**不污染图块体系**。
- **运行过程全透明**：所有模型调用（DeepSeek/qwen）的输入输出摘要实时打印（`[DeepSeek]`/`[qwen]` 前缀），用户可随时看到每一步在想什么、输出了什么。
- **教材级文案**：整个 PPT 视为一部完整教材，DeepSeek 自主划分若干章（`# 第X章 章名`）、每章若干节（`## 第X节 节名`，一节可含多页），每页内容首行标注所属章节；**500 字为限**（原文 ≤500 字扩写到不少于 500 字；>500 字直出整理，不改原意、不增字数）。
- **生命周期自管**：全流程成功即清理过程文件；中断保留断点、二次运行自动接续。

## 二、安装与环境要求

### 系统要求

| 项 | 要求 |
|---|---|
| 操作系统 | **Windows 10/11**（强烈建议；渲染通道支持 PowerPoint/WPS COM） |
| Python | 3.9 及以上 |
| 演示应用 | **Microsoft Office PowerPoint 或 WPS 演示**（渲染自动探测：有 Office 用 Office，只有 WPS 用 WPS；均无则渲染优雅降级） |
| API Key | `DEEPSEEK_API_KEY`（语义解读/教材文案）、`DASHSCOPE_API_KEY`（qwen VLM 兜底） |

### 安装

```bash
pip install pptx-wzq            # 核心
pip install "pptx-wzq[ocr]"     # 含本地公式 OCR（pix2tex，可选）
```

安装后自带完整文档（使用手册 + 技术分析）：

```python
from pptx_wzq import docs
print(docs.__path__)   # 查看 docs/ 目录（含 html/pdf/md）
```

> **权重来源与许可**：图片过滤内置的 YOLO 模型权重 `yolov5su.pt` 来自
> [ultralytics](https://github.com/ultralytics/ultralytics)（YOLO 官方），采用 **AGPL-3.0**
> 许可证随包分发；商用或闭源使用请自行评估其开源条款。

## 三、快速开始

```bash
# 一条命令跑完 8 环节（推荐）
pptx-paser "C:\课件\战略管理.pptx" -o "C:\输出\战略管理知识库"

# 跳过相关性与文案（省 Token）
pptx-paser "C:\课件\战略管理.pptx" -o out --skip related,author

# 断点续传（中断后重跑同命令自动接续）
pptx-paser "C:\课件\战略管理.pptx" -o out
```

## 四、命令一览

| 命令 | 功能 |
|---|---|
| `pptx-paser` | 总编排器：8 环节一条命令，断点续传 / 成功清理 |
| `pptx-blocks` | 图块提取+解构+渲染；`--caption-sources` 按 sources/ 顺序解读 |
| `pptx-text` | 文本提取（in_group / 表格 Markdown） |
| `pptx-formula` | 公式提取 → LaTeX |
| `pptx-caption` | 图片/图块解读（模型路由） |
| `pptx-related` | 块相关性过滤 + 审计 |
| `pptx-author` | 整篇分章分节教材文案（500 字为限） |
| `pptx-bind` | 图文绑定 JSON |
| `pptx-del` | **图块删除后处理**：`pptx-del images\slide_29_blk_01.png -all`，删除指定块及其全部关联（images/sources/JSON/binding/captions），等价于该组合不存在；默认备份 + dry-run + 一致性校验 |
| `pptx-html` / `pptx-deck` | 教材 HTML / 教学 Deck 导出 |

## 五、输入 PPTX 预处理要求

> **核心思想：用 PPT 自带工具给解析器打「块边界」标注。** 组合的数量 = 该页图块数量的上限基准，可据此验收。预处理不是必须的（无组合也能跑），但组合能让图块提取完全确定、可回归。

| 对象 | 预处理操作 | 解析器行为 |
|---|---|---|
| **逻辑图/示意图（要整体成块）** | 「开始 → 排列 → 组合」（Ctrl+G）把底图+文字框+箭头合成一个组合 | 整个组合 = 1 个 group 块，XML 段导出 sources/，渲染图存 images/ |
| **嵌套组合** | 有意为之才用：嵌套 = 外层块的子结构；想分开就移出外层 | 嵌套组合并入外层 children，不单独成块 |
| **Visio / vsdx 工程图** | **不要**组合进其他形状（否则被吞并）；保持独立 OLE 对象 | 独立成 visio 块：可剥离 → .vsdx；不可剥离 → XML 段 |
| **公式** | 行内公式保持独立（不组合）→ 自动进 formulas.md；想并入图块就移进组合 | 组合内公式随块转 LaTeX；非组合公式独立提取 |
| **首页/尾页** | 无需操作：默认按页序跳过第 1 页与末页的图块 | 封面/致谢不产块（文本/公式仍提取） |
| **小像素图（装饰图标）** | 小于整页 20% 默认舍弃；想保留 → 组合进相邻图形 | 组合内小图豁免；非组合小图丢弃并写入审计 |
| **表格** | 无需操作 | 输出 Markdown 表格（texts.md），不产块 |
| **带裁剪的图片（srcRect）** | 无需操作（自动保持显示一致） | 元数据/资源/渲染/描述全程按裁剪 |

## 六、工作流程（8 环节）

```
① blocks → ② text → ③ formula → ④ caption → ⑤ related → ⑥ author → ⑦ blocks_json → ⑧ 输出
```

| 环节 | 职能 | 消耗 | 要点 |
|---|---|---|---|
| ① blocks | 图块提取 + 解构 | 本地 | 单阶段确定性拆块（grpSp→group / visio / 像素图≥20% / SVG-WMF）；XML 段每组合一个文件导出 sources/；rldimg 资源图落盘；PowerPoint COM 渲染块 PNG |
| ② text | 文本提取 | 本地 | 排除页眉页脚/母版固定文本；组内文字标「[图块内文本]」；表格输出 Markdown |
| ③ formula | 公式提取 | 本地 | 非组合公式 → LaTeX（OMML/MTEF/OCR 三级）；**组合内公式也进 formulas.md**（双通道：组合块 XML 段保留 + 语义存档，标注「组合内公式」） |
| ④ caption | 图块 AI 解读 | DeepSeek+qwen | 按 sources/ 顺序：.xml → DeepSeek 读 XML（公式转 LaTeX）；.png → qwen 兜底；调用输入输出实时打印 |
| ⑤ related | 相关性过滤 | DeepSeek | 剔除 logo/作者/装饰块 → related_filter.json 审计；**并发判定 + 页面正文为空时保守保留**（防误删） |
| ⑥ author | 教材文案 | DeepSeek | 整篇自主分章分节（一节可含多页），每页标注章节；≤500 字扩写、>500 字直出整理 |
| ⑦ blocks_json | 组装 + 语义增强 | DeepSeek+qwen | visual_blocks.json（v2.0）+ semantic_description + 跨模态关系 + binding 导出；**qwen 视觉兜底仅对 caption 未解读的块触发**（读渲染图），DeepSeek 语义增强并发执行 |
| ⑧ 输出 | 归位 + 清理 | 本地 | 成功即删过程文件；中断保留断点自动接续 |

**模型路由（规则 11/16）**：DeepSeek-V4-Flash 读 sources/ 的 XML 段（grpSp/Visio-XML/SVG-WMF，公式转 LaTeX）；qwen3.7-plus 读像素图原图——仅纯像素图块 / DeepSeek 空响应兜底。

## 七、输出文档体系与内部格式

### 结果目录结构（成功运行后）

```
<名>_result/
├─ sources/                          # 图块源资源（解读唯一输入源）
│  ├─ slide_07_blk_01_grp.xml        # grpSp XML 段（页注释 + 原生前缀，规则10）
│  ├─ slide_05_blk_02.vsdx            # Visio 可剥离（规则13）
│  ├─ slide_05_blk_03_ole.xml         # Visio 不可剥离 → XML（规则14）
│  ├─ slide_08_blk_04_vec.xml         # SVG/WMF 矢量 XML（规则15）
│  ├─ slide_19_blk_01.png             # 像素图块原图（qwen 解读输入，规则16）
│  └─ rldimg/                         # grpSp XML 内 r:embed 引用的资源图片
├─ images/                           # 块渲染图（PowerPoint，仅供人阅览）
├─ <名>_visual_blocks.json             # 核心结构化（pptx_multimodal_slide_v2.0）
├─ <名>_visualBlock_text_binding.json  # 图文关联（pptx_visual_block_text_binding_v1.0）
├─ <名>_textbook.md                   # 教材文案（篇→章→节→页，每页标注章节）
├─ <名>_captions.md                   # 块解读（绑定 sources/ 文件名 + 解读通道）
├─ <名>_texts.md / _text_entries.json   # 文本清单（组内文字标[图块内文本]）
├─ <名>_formulas.md                     # 非组合公式（LaTeX）
└─ <名>_related_filter.json              # 相关性过滤审计
（v2.0 起成功即清理：无 过程文件/；中断时保留断点供续传）
```

### VisualBlock 内部格式（visual_blocks.json 的块对象）

```json
{
  "block_id": "blk_01", "page": 7, "block_type": "战略管理概念框架",
  "bbox": {"x": 142.3, "y": 226.4, "w": 1026.2, "h": 400.2},
  "z_index_range": [12, 45], "is_single": false,
  "text": "战略哲学 商道 天道 人道 …",
  "assets": {
    "xml_source": "./sources/slide_07_blk_01_grp.xml",
    "raster_png": null,
    "rldimg": ["./sources/rldimg/slide_07_blk_01_image4.png", "…"]
  },
  "internal_structure": {"nodes": ["…"], "edges": ["…"]},
  "semantic_description": {
    "block_type": "战略管理概念框架",
    "expression_goal": "展示战略管理概念框架的核心逻辑",
    "expression_role": "将抽象概念通过战略哲学/商道/天道/人道具象化…",
    "expression_features": ["概念框架", "层次结构", "关系图"],
    "vlm_caption": "该图块以“战略哲学”为中心…",
    "teaching_use": "教学辅助图示",
    "formula_latex": "", "caption_source": "deepseek_xml"
  },
  "member_obj_ids": ["…"], "vector_resources": []
}
```

### textbook.md 结构（v2.3 起）

```
# <名> 教材文案
# 第1章 战略管理概述          ← DeepSeek 自主命名
## 第1节 战略的概念与本质      ← 自主命名，一节可含多页
## 第 N 页
> 所属章节：第1章 战略管理概述 · 第1节 战略的概念与本质   ← 每页首行标注
（不少于 500 字正文…）
```

## 八、技术架构

| 模块 | 职责 | Token |
|---|---|---|
| `extract_pptx_images.py` | OOXML 原子对象提取（图片/形状/grpSp/表格/公式/图表）；XML 段原生提取；srcRect 裁剪；**OLE 预览图判定（preview_of：公式/Visio 快照不独立成块）**；PowerPoint/WPS COM 渲染（ProgID 自动探测 + DispatchEx 独立实例） | 0 |
| `extract_texts.py` | 页面文本（in_group 标记）、表格 Markdown | 0 |
| `visual_blocks.py` | 单阶段确定性拆块、块渲染、语义增强（DeepSeek **并发**）、跨模态关系、**qwen 视觉兜底按需触发（仅 caption 未解读块）** | DeepSeek+qwen |
| `cli_blocks.py` | XML 段导出、rldimg 复制、caption 路由、binding、captions.md | DeepSeek+qwen |
| `cli_author.py` | 整篇分章分节文案、500 字为限、自动分批（跨批章节延续） | DeepSeek |
| `cli_related.py` | 相关性过滤（**并发判定 + 正文为空保守保留**）+ 审计 | DeepSeek |
| `cli_paser.py` | 总编排、断点续传（state.json）、归位、成功清理 | — |

## 九、图块识别规则（16 条）

### 识别层（1-9）——「组合即图块声明」

| # | 规则 | 实现 |
|---|---|---|
| 1 | 一个 grpSp 组合即是一个图块，组合内一切内容读取为该图块内容 | kind="group"，children 递归收编 |
| 2 | 嵌套组合不单独提取 | 嵌套仅作外层 children |
| 3 | OLE Visio/vsdx 独立成块（除非在组合内） | kind="visio" 分支 |
| 4 | 非组合像素图单独成块；重叠文本并入 | raster 独立块 + 重叠文本并入 |
| 5 | 组合内公式作块内容；非组合公式独立提取 | formula 排除 in_group |
| 6 | 首页/尾页图块舍弃 | skip_cover_pages 整页跳过 |
| 7 | 非组合像素图 < 整页 20% 舍弃 | raster_min_area_ratio=0.20 |
| 8 | 表格仍读取为表格（Markdown） | 表格移交 text 步骤 |
| 9 | grpSp 内 srcRect 须全程一致 | children 携带 src_rect + 裁剪落盘 |

### 解构/解读层（10-16）

| # | 规则 | 实现 |
|---|---|---|
| 10 | grpSp 保留整段 XML，页标记，每组合一个独立 .xml 存 sources/ | `sources/slide_{页}_{块id}_grp.xml` |
| 11 | grpSp 块 caption 用 DeepSeek 读 XML；公式转 LaTeX | `_ds_read_xml` + 超长压缩 + 重试 |
| 12 | grpSp 块 PowerPoint 渲染 PNG 存 images/；PNG 不送 qwen | COM ExportAsFixedFormat PDF + PyMuPDF |
| 13 | Visio 可剥离 → .vsdx 存 sources/ + PNG 存 images/ | 原生剥离 + 渲染 |
| 14 | Visio 不可剥离 → XML 段，同 grpSp | `sources/slide_{页}_{块id}_ole.xml` |
| 15 | SVG/WMF 同 Visio：尽量 XML 段，不行才 PNG | `sources/slide_{页}_{块id}_vec.xml` |
| 16 | 仅无法 DeepSeek 解读的块才送 qwen | 模型路由 + 空响应降级链 |

## 十、生命周期与可靠性

- **成功即清理**：全流程成功后中间产物（by_page/atomic_objects.json/manifest/各步骤工作目录）全部删除，结果目录只留交付物；删除失败打印警告而非静默。
- **中断续传**：中断保留 work + `state.json`（步骤状态机 pending/running/partial/done/failed）；再运行跳过已完成、自动接续缺失；`doc_md5` 换源检测 → 全量重跑；author 支持缺失页补跑（`--pages`）。

## 十一、常见问题

| 问题 | 说明 |
|---|---|
| 没有 Office 会怎样？ | 渲染降级：无块渲染图（images/ 为空）；XML/JSON/文案等其余产物正常 |
| 只有 WPS 没有 Office？ | 自动探测 `Kwpp.Application` COM，用 WPS 演示渲染（ProgID 依次尝试，有 Office 优先） |
| 渲染报 0x80070002（文件不可用）？ | 大概率是 PPT 正被 PowerPoint/WPS 打开编辑，或相对路径未解析——v3.0.2 起强制绝对路径 + 完整错误提示 |
| 没有 API Key？ | 跳过对应解读/文案步骤；`PPTX_PASER_NO_VLM=1` 可跳过 VLM 全流程 0 Token |
| 中途中断？ | 重跑同命令自动接续（state.json + doc_md5 换源检测） |
| DeepSeek 对某 XML 段空响应？ | 自动降级 qwen 读渲染图 → 规则模板兜底，保证每块有解读 |
| blocks_json 步骤为什么慢？ | v3.0.0 起语义增强/相关性判定为普通生成 + 并发（实测 10~160 倍提速）；若仍慢请确认已升级 |
| 如何查安装版本与文档？ | `pptx-paser --version`；`from pptx_wzq import docs` |

## 十二、版本演进

| 版本 | 里程碑 |
|---|---|
| 1.5.0 | 六步管线重构（img 并入 blocks 自举）、三件套交付物、版本统一 |
| 2.0.0 | 16 条规则落地：组合即块、单阶段确定性拆块、XML 段导出、DeepSeek caption 路由、PowerPoint 渲染、_organize 修复 |
| 2.1.0 | captions 绑定 sources/ 源文件 + 通道标注；每页扩写 500 字 |
| 2.2.0 | textbook 规则改 500 字为限（不足扩写、超出直出整理 _tidy_direct） |
| 2.3.0 | Author 整篇自主分章分节（一节可含多页），每页标注章节 |
| 2.4.0 | 技术分析报告（html/pdf/md）随安装包分发 |
| 2.5.0 | 使用手册+技术分析合并文档（html/pdf）、README 全量更新 |
| 2.5.1 | 渲染静默降级修复：dependencies 补 pymupdf；渲染失败给明确警告 |
| 2.5.2 | images/ 不再被 caption/blocks_json 步骤清空（--skip-render）；渲染子进程改 DispatchEx 独立实例（Open 0x80070002） |
| 2.6.0 | **OLE 预览图判定**（pic 前 300 字符检测 oleObj → preview_of，公式/Visio 快照不独立成块，vector 块 63→0）；**组合内公式进 formulas.md**（双通道 + 标注） |
| 3.0.0 | **语义增强/关系生成提速**（去 thinking + 并发 8，实测 10 倍）；**模型调用实时打印**（[DeepSeek]/[qwen] 输入输出可见） |
| 3.0.1 | **related 相关性过滤提速**（并发 + 去 thinking，实测 162 倍）；**正文为空保守保留**（消除误删） |
| 3.0.2 | 渲染 Open 失败根因修复：pptx 强制 resolve 绝对路径（子进程 CWD 继承问题）+ 错误完整显示 |
| 3.1.0 | **WPS 渲染支持**：ProgID 自动探测（PowerPoint.Application → Kwpp.Application）+ ExportAsFixedFormat 简化参数回退 |
| 3.1.1 | **qwen 视觉兜底按需触发**：仅 caption 未解读的块读渲染图（调用 73→13）；图路径接通 images/（修复此前"从未真正读图"） |
| 3.1.2 | 文档体系更新：README/使用手册/技术分析同步 3.0.x~3.1.x 全部变更（渲染通道、性能优化、实时打印、按需兜底） |
| 3.1.3 | README 补「开发原因」；**PyPI 主页显示完整 README**（wheel METADATA 携带 long_description） |
| 3.1.4 | ExportAsFixedFormat E_FAIL 可读化 + 旧 PDF 占用警告 |
| 3.1.5 | author 大输出批截断根治：max_tokens + 缺页重试 + 诚实统计 |
| 3.2.0 | **新增 `pptx-del` 命令**：流水线完成后的图块删除后处理（清理用户预处理不干净的 grpSp 组合），删除块及其全部关联输出，默认备份 + dry-run + 一致性校验 |

## 十三、文档与许可

- **文档**：完整「使用手册与技术分析」（HTML/PDF）随安装包分发于 `src/pptx_wzq/docs/`，安装后可 `from pptx_wzq import docs` 定位；README（本文档）为全量文字版。
- **许可证**：MIT（本项目代码）；随包 YOLO 权重 `yolov5su.pt` 为 **AGPL-3.0**（见上文）。
