# pptx-wzq · PPT 多模态教学知识库自动化构建

> **作者：吴振谦 · wuzhenqian@nbu.edu.cn · QQ：38328063**
> 本项目帮助高校教师把教学 PPT 自动转化为图文并茂的多模态教学知识库。

把一份教学 PPT 自动转化为**可检索、可复用、可再加工**的多模态教学知识库：
图片集 + 图片 AI 理解 + 公式 LaTeX + 教材文案 + 图文绑定 JSON，并可一键导出
教材 HTML 与教学 Deck。

```
PPT(.pptx)
  → 文本提取（文本ID+坐标）  pptx-text
  → 公式提取        pptx-formula
  → 图片提取/过滤 + 原子对象（shape/connector/table）  pptx-img
  → 可视逻辑块解析 + Semantic Captioning（合并图片AI解读）  pptx-blocks
  → 可视逻辑块相关性过滤（剔除 logo/作者/单位等无关块）  pptx-related
  → 教材文案（原文超 300 字直出不扩写）  pptx-author
  → 可视逻辑块 JSON 组装（pptx_multimodal_slide_v2.0 全栈解析）  pptx-blocks
  → 一条命令编排（日志 + 断点续传）  pptx-paser
  → 教材 HTML/Deck  pptx-html / pptx-deck
```

## 安装

```bash
pip install pptx-wzq            # 核心
pip install "pptx-wzq[ocr]"     # 含本地公式 OCR（pix2tex，可选）
```

Windows 下会自动安装 pywin32（用于 PowerPoint 渲染矢量图）。
首次运行 `pptx-paser` 会自动补齐缺失组件、检查 API Key
（DeepSeek / 阿里云百炼）并引导注册。

> **权重来源与许可**：图片过滤内置的 YOLO 模型权重 `yolov5su.pt`
> 来自 [ultralytics](https://github.com/ultralytics/ultralytics)（YOLO 官方），
> 采用 **AGPL-3.0** 许可证随包分发；商用或闭源使用请自行评估其开源条款。

## 快速开始

```bash
# 一条命令跑完整流水线（自动装依赖 / 引导 Key / 预估 Token / 确认执行）
pptx-paser 第九章功率放大电路.pptx -o output

# 结果目录：
#   output/images/ + output/<名>_captions.md + <名>_textbook.md + <名>_binding.json

# 导出教材 HTML（MathJax 公式渲染）
pptx-html output -o output/第九章.html

# 分步执行 / 断点续跑
pptx-img 课件.pptx -o out
pptx-caption out/images -o cap.md --texts out/texts.md --formulas out/formulas.md
pptx-related out -o out/课件_captions.md --texts out/texts.md
pptx-paser 课件.pptx -o out --skip img,formula,text
pptx-paser 课件.pptx -o out --dry-run    # 预览续跑/执行计划（不执行）
pptx-paser 课件.pptx -o out --reset      # 强制从头重跑
```

> **断点续传**：`pptx-paser` 在结果目录写 `state.json`（步骤状态机）与
> `pipeline.log`（运行日志）。中途中断后重跑同一条命令，会自动跳过已完成
> 步骤、从断点续跑（图片解读按图、教材文案按页、图文绑定按页），
> 无需手动 `--skip`；`--dry-run` 可先预览计划。

## 指令一览

| 命令 | 功能 | 模型 | Token |
|---|---|---|---|
| `pptx-text` | 逐页文本/表格提取（排除页眉页脚），每条文本带 text_id 与幻灯片坐标 x/y/w/h | — | 0 |
| `pptx-formula` | 公式三路径提取（OMML / EQ3-MTEF / OCR）→ LaTeX | — | 0 |
| `pptx-img` | 图片提取 + 三路过滤 + WMF 识别/渲染 + 公式图片补提；Visio OLE 按容器存 `.vsdx`/`.vsd`；emf/wmf/svg 规范化 svg（失败回退 wmf）；记录幻灯片坐标；**收集原子对象（shape/connector/table → atomic_objects.json）** | — | 0 |
| `pptx-blocks` | **可视逻辑块全栈解析**：空间聚类（并查集）把每页拆成 1~6 块 → 块渲染 PNG → VLM 判定块类型 + Semantic Captioning（表达目标/作用/特征/图文描述/教学用途）→ 图/树拓扑 → 跨模态关系 → `visual_blocks.json`（pptx_multimodal_slide_v2.0 schema）+ 块级 captions.md；`--no-vlm` 纯规则模式 | qwen3.7-plus | 有 |
| `pptx-caption` | （旧）图片 AI 解读——已并入 `pptx-blocks`，保留兼容 | qwen3.7-plus | 有 |
| `pptx-related` | 可视逻辑块相关性过滤：DeepSeek 判定块与正文是否相关，无关块（logo/作者/单位/项目类别/重复装饰等）连同解读一并删除，写审计 json | deepseek-v4-flash | 有 |
| `pptx-author` | 教材文案（学科推断 + 整文生成 + 自适应分批）；某页原文超 300 字直接提取不扩写（`--no-expand-threshold`） | deepseek-v4-flash | 有 |
| `pptx-bind` | （旧）图文绑定 JSON——保留兼容，新流程由 `pptx-blocks` 的 `visual_blocks.json` 替代 | deepseek-v4-flash | 有 |
| `pptx-paser` | 整体编排（**text → formula → img → blocks → related → author → blocks_json**）：环境自检 / Key 引导 / Token 预估 / **日志+断点续传**（state.json / pipeline.log）/ `--dry-run` / `--reset` | — | — |
| `pptx-html` | 单文件教材 HTML（MathJax 公式 + PPT 原生标题） | — | 0 |
| `pptx-deck` | 教育蓝图风格教学 Deck（示例脚本） | — | 0 |

## API Key

| 用途 | 平台 | 模型 |
|---|---|---|
| 可视逻辑块 Semantic Captioning / 公式图片识别 | 阿里云百炼（bailian.console.aliyun.com） | qwen3.7-plus |
| 教材文案 / 学科判断 / 相关性判定 / 跨模态关系 | DeepSeek（platform.deepseek.com） | deepseek-v4-flash |

缺失时 `pptx-paser` 会打印注册引导并交互写入环境变量（`DEEPSEEK_API_KEY` /
`DASHSCOPE_API_KEY`）。

## 产物

```
结果目录/
├─ images/              可视逻辑块渲染图 + 原子图片集（slide_NN_blk_NN.png）
├─ sources/             矢量源文件归档（vsdx/svg/wmf/emf，可编辑资产）
├─ <名>_texts.md        文本清单（ID | 类型 | 文本 | 坐标；表格行类型）
├─ <名>_captions.md     可视逻辑块级 AI 解读（# images 图片 AI 解读，块为条目单位）
├─ <名>_textbook.md     教材文案（直出标注在节标题后；原文超 300 字直出）
├─ <名>_visual_blocks.json  可视逻辑块全栈解析（pptx_multimodal_slide_v2.0：
│                       slide_info/textual_content/visual_blocks（几何/拓扑/
│                       资源/类型/semantic_description）/cross_modal_relations/
│                       summary；替换原 binding.json）
├─ images_meta.json     图片元数据（图片→页/尺寸/来源）
├─ <名>_related_filter.json  可视逻辑块相关性过滤审计（被删块 + 原因）
├─ state.json           断点续传状态机（含 doc_md5/tool_version，换源提示）
├─ pipeline.log         运行日志（每步开始/完成/失败时间戳）
└─ 过程文件/            中间产物（by_page / manifest / texts / formulas / …）
```

## 应用

- 课程知识内容体系自动构建（章 → 节 → 知识点 → 公式 + 配图）
- 基于学情测验的个性化学习内容生成
- 教材二次开发与多形态出版（HTML / Deck / Word / PDF）
- 智能问答与 RAG 检索（binding.json 作索引，答案图文并茂）
- 试题与测验生成（按 Bloom 认知层次，真实配图出题）

## License

MIT
