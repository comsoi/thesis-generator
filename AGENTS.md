# AGENTS.md

## 项目概述

Markdown → Typst → DOCX 论文自动构建系统，支持中英文双版本输出。

## 构建命令

```bash
make          # 完整构建 → result/thesis-cn.docx + result/thesis-en.docx
make gen-src  # markdown/*.md → src/main.typ
make content  # src/main.typ → result/main_content.docx（英文） / main_content_cn.docx（中文）
make gen-prelim     # 前置部分模板渲染
make merge-docs     # 合并中英文最终文档
make clean          # 清空 result/ 和 .cache/
```

## 流程图

使用 TikZ (XeLaTeX) 编写，源文件 `flowchart/*.tex`，`make gen-flowchart` 编译为 PNG 到 `figure/`。

在 markdown 中用 `images/` 路径引用（软链接 `markdown/images → ../figure`）。

## 环境

- `uv run` 管理 Python 脚本，禁止使用系统 python

## 变量命名

Makefile 变量统一风格：`OBJECT_TAG`，如 `PRELIM_TEMPLATE_CN`、`RENDER_SCRIPT`、`MERGE_SCRIPT`。

## 易犯错误

1. **Markdown 表格**：表题必须用 `: 标题`（Pandoc 语法）紧邻表格，**不能**写成 `表 1-1 xxx` 普通段落，否则不会被识别为表格。
2. **章节标题**：用 `# 标题`，`md2typst.py` 会自动剥离以下手写编号后交给 Typst 统一编号：

   | 层级 | 支持格式 | 正则 |
   |---|---|---|
   | `#` 一级 | `第一章 xxx`、`第N章 xxx`、`1 xxx`、`1. xxx` | `第[一二三...\d]+章` 或 `\d+\.?\s+` |
   | `##` 二级 | `1.1 xxx` | `\d+\.\d+\s+` |
   | `###` 三级 | `1.2.1 xxx` | `\d+\.\d+\.\d+\s+` |

   手写编号会导致 Typst 二次编号，产生 `1 1.1 xxx` 之类错误。
3. **图片路径**：markdown 中写 `![caption](images/xxx.png)`。`markdown/images` 是软链接指向 `../figure`，`md2typst.py` 自动把 `images/` 替换为 `figure/`。不要直接写 `figure/`。

## 外部素材导入

给定一个混乱的项目目录（下称 `<proj>`），AI 需将其内容整理为 thesis-generator 可用的 markdown 章节、图片、参考文献。以下为通用处理流程。

### 1. 扫描分类

先遍历 `<proj>`，将文件按类型分类：

| 类别 | 识别特征 | 用途 |
|------|---------|------|
| 分析报告 | `.md` 文件，含中文标题/段落 | 提取为论文章节草稿 |
| 图片 | `.png`/`.jpg`/`.pdf`，在 `tikz/`、`plots/`、`*_plots/` 下 | 归集到 `figure/` |
| 论文 PDF | `.pdf`，在 `papers/` 或散落顶层，文件名为人名/标题 | 提取 BibTeX 引用 |
| 源代码 | `.cpp`/`.h`/`.py`/`.m` | 参考算法实现，不导入论文 |
| 实验数据 | `.csv`/`.tsp`/`.txt`，在 `output/`/`data/` 下 | 不导入论文，但不可删除 |
| 编译产物 | `*_experiment`/`test_quick*`/`.o`/`.aux`/`.log`/`__pycache__/` | 噪声，删除或 gitignore |
| 构建文件 | `Makefile`/`CMakeLists.txt` | 保留在原位 |

### 2. 清理噪声

执行以下动作（不改动源代码和实验数据）：

- 删除 `.partial`（下载中断残留）、`.aux`、`.log`、`__pycache__/`
- 删除空目录
- 编译产物加入 `<proj>/.gitignore`（不删除文件本身）
- 删除重复文件，保留一份
- 报告 `.md` 中的图片引用路径暂时保留（后续步骤统一修正）

### 3. 提取论文章节

阅读 `<proj>` 中的所有分析报告 `.md`，判断内容覆盖哪些主题（如算法原理、数据集描述、实验设计、结果分析），然后按以下规则生成 `markdown/chapter-N-*.md`：

- **章节拆分**：一个 `.md` 覆盖一个独立主题，文件名 `chapter-{N}-{slug}.md`，N 从 2 开始（1 为绪论）
- **标题格式**：一级 `# 标题`、二级 `## 标题`，不得手写编号（由 Typst 自动编号）
- **表格标题**：用 `: 标题` 紧邻表格下一行，不要写成 `表 N-M xxx`
- **图片引用**：统一改为 `images/xxx.png`，实际图片同时拷贝到 `figure/`
- **公式**：LaTeX 格式不变，pandoc 自动转换
- **引用**：用 `[@key]` 格式，BibTeX 条目写入 `src/references.bib`
- **不要照搬原文**：草稿可能有口语化、重复、标题层级混乱等问题，需整理为正式论文语气

### 4. 图片归集

将 `<proj>` 中所有论文用到的图片拷贝到 `figure/`：

- 每个 `<proj>` 子目录下的 `.png` 只取最终版本（排除带 `-1`、`-v2`、`fix` 后缀的中间版本）
- 保留源文件名，不要重命名
- 在 markdown 中引用时统一用 `images/<filename>.png`

### 5. 参考文献提取

- 从 `<proj>` 中的论文 PDF 提取关键文献的 BibTeX 条目
- 写入 `src/references.bib`（追加，不覆盖已有条目）
- 优先提取被报告 `.md` 明确引用的论文
- BibTeX key 格式：`<第一作者姓氏><年份><关键词>`，如 `storn1997de`

### 6. 最终校验

完成上述步骤后，运行 `make` 验证构建是否通过。若 pandoc 报引用缺失，补全 `references.bib`；若图片 404，检查 `figure/` 中是否存在且 markdown 引用路径正确。
