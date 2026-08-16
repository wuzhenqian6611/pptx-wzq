# pptx-wzq · PPT 多模态教学知识库自动化构建

> **作者：吴振谦 · wuzhenqian@nbu.edu.cn · QQ：38328063**
> 本项目帮助高校教师把教学 PPT 自动转化为图文并茂的多模态教学知识库。

把一份教学 PPT 自动转化为**可检索、可复用、可再加工**的多模态教学知识库：
图片集 + 图片 AI 理解 + 公式 LaTeX + 教材文案 + 图文绑定 JSON，并可一键导出
教材 HTML 与教学 Deck。

```
PPT(.pptx)
  → 图片提取/过滤（vsdx 直接存 .vsdx，矢量规范化 svg/wmf）  pptx-img
  → 公式提取        pptx-formula
  → 文本提取（文本ID+坐标）  pptx-text
  → 图片AI解读      pptx-caption
  → 图文相关性过滤（剔除 logo/作者/单位等无关图）  pptx-related
  → 教材文案（原文超 300 字直出不扩写）  pptx-author
  → 图文绑定（图片ID/文本ID/坐标/逻辑关系）  pptx-bind
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
| `pptx-img` | 图片提取 + 三路过滤 + WMF 识别/渲染 + 公式图片补提；Visio OLE 按容器存 `.vsdx`/`.vsd`；emf/wmf/svg 规范化 svg（失败回退 wmf）；记录幻灯片坐标 | — | 0 |
| `pptx-formula` | 公式三路径提取（OMML / EQ3-MTEF / OCR）→ LaTeX | — | 0 |
| `pptx-text` | 逐页文本/表格提取（排除页眉页脚），每条文本带 text_id 与幻灯片坐标 x/y/w/h | — | 0 |
| `pptx-caption` | 图片 AI 解读（文档上下文模式，教材角度）；`--resume` 断点续跑 | qwen3.7-plus | 有 |
| `pptx-related` | 图文相关性过滤：DeepSeek 判定图片与正文是否相关，无关图（logo/作者/单位/项目类别/重复装饰等）连同解读一并删除，写审计 json | deepseek-v4-flash | 有 |
| `pptx-author` | 教材文案（学科推断 + 整文生成 + 自适应分批）；某页原文超 300 字直接提取不扩写（`--no-expand-threshold`） | deepseek-v4-flash | 有 |
| `pptx-bind` | 图文绑定 JSON：含 图片ID/文本ID/幻灯片坐标/DeepSeek 生成的 ≤60 字图文逻辑关系；`--resume` 按页续跑 | deepseek-v4-flash | 有 |
| `pptx-paser` | 整体编排：环境自检 / Key 引导 / Token 预估 / **日志+断点续传**（state.json / pipeline.log）/ `--dry-run` / `--reset` | — | — |
| `pptx-html` | 单文件教材 HTML（MathJax 公式 + PPT 原生标题） | — | 0 |
| `pptx-deck` | 教育蓝图风格教学 Deck（示例脚本） | — | 0 |

## API Key

| 用途 | 平台 | 模型 |
|---|---|---|
| 图片 AI 解读 / 公式图片识别 | 阿里云百炼（bailian.console.aliyun.com） | qwen3.7-plus |
| 教材文案 / 学科判断 | DeepSeek（platform.deepseek.com） | deepseek-v4-flash |

缺失时 `pptx-paser` 会打印注册引导并交互写入环境变量（`DEEPSEEK_API_KEY` /
`DASHSCOPE_API_KEY`）。

## 产物

```
结果目录/
├─ images/              教学图片集（PNG + SVG/WMF 矢量 + vsdx/vsd 归档）
├─ <名>_captions.md     图片理解（已剔除 logo/作者/单位等无关图）
├─ <名>_textbook.md     教材文案（每页一节；原文超 300 字页直接提取）
├─ <名>_binding.json    图文绑定 v3（按页 {page, text, images[{file, caption,
│                       source, kind, image_id, page, paragraph, text_id,
│                       w/h/x/y 坐标, position(图片作用, LLM), relation
│                       (图文逻辑关系, LLM)}]}；无图文本段不进入绑定）
├─ <名>_related_filter.json  图文相关性过滤审计（被删图 + 原因）
├─ state.json           断点续传状态机（步骤 status：pending/running/
│                       partial/done/failed）
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
