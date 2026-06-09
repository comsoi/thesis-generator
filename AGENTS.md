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

1. **Markdown 表格**：表题用 `: 标题`（Pandoc 语法），与表格之间**留一空行**（紧邻会导致标题在上方时解析失败），**不能**写成 `表 1-1 xxx` 普通段落。
2. **章节标题格式**：用 `# 标题`，`md2typst.py` 会自动剥离以下手写编号后交给 Typst 统一编号：

   | 层级 | 支持格式 | 正则 |
   |---|---|---|
   | `#` 一级 | `第一章 xxx`、`第N章 xxx`、`1 xxx`、`1. xxx` | `第[一二三...\d]+章` 或 `\d+\.?\s+` |
   | `##` 二级 | `1.1 xxx` | `\d+\.\d+\s+` |
   | `###` 三级 | `1.2.1 xxx` | `\d+\.\d+\.\d+\s+` |

   手写编号会导致 Typst 二次编号，产生 `1 1.1 xxx` 之类错误。
3. **图片路径**：markdown 中写 `![caption](images/xxx.png)`。`markdown/images` 是软链接指向 `../figure`，`md2typst.py` 自动把 `images/` 替换为 `figure/`。不要直接写 `figure/`。
4. **行内公式闭合 `$` 后跟数字需空格**：`$\beta$1` 中 `$1` 被 Pandoc 当作美元金额而非公式分隔符，输出 `$$1`（公式丢失）。加空格：`$\beta$ 1`（输出 `β 1`）。若不想要空格，把数字并入公式：`$\beta 1$`（输出 `β1`）。注意 `$` 后跟字母无此问题（如 `$\beta$a` 正常）。

## 外部素材导入

给定一个混乱的项目目录（下称 `<proj>`），AI 需将其内容整理为 thesis-generator 可用的 markdown 章节、图片、参考文献。

**核心原则**：整理完成后，AI **必须先学习** `markdown/` 目录下现有章节的写法风格（摘要、致谢、绪论等），然后**删除这些模板文件**（`.md` 及 `flowchart/` 下的流程图），再基于 `<proj>` 的内容和学到的风格**重写**所有 markdown 章节。

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

### 2. 学习现有写法并清理模板

- **阅读现有模板**：通读 `markdown/` 下所有 `.md` 文件，掌握本项目的写作风格、标题格式、表格/图片/公式用法
- **学习流程图**：阅读 `flowchart/` 下现有的 `.tex` 源文件，掌握流程图的风格、布局方式、节点样式与连线规则
- **删除模板文件**：将 `markdown/` 下的所有 `.md` 文件（如 `abstract-*.md`、`acknowledgment.md`、`chapter-*.md`）及 `flowchart/` 下的流程图源文件全部删除
- 删除后 `markdown/` 目录应仅剩 `images/` 软链接

### 3. 提取论文章节

基于 `<proj>` 内容，参照已学习的风格重写所有章节，生成新的 `markdown/*.md`：

- **章节拆分**：按主题拆分，文件名 `chapter-{N}-{slug}.md`，N 从 2 开始（1 为绪论）
- **标题格式**：一级 `# 标题`、二级 `## 标题`，编号遵循本项目的章节标题格式规则
- **表格标题**：用 `: 标题`（Pandoc 语法），与表格之间留一空行
- **图片引用**：统一用 `images/xxx.png`，图片同时拷贝到 `figure/`
- **公式**：LaTeX 格式，注意行内公式 `$` 后跟数字需空格
- **引用**：用 `[@key]` 格式，BibTeX 条目追加到 `src/references.bib`
- **不要照搬原文**：需整理为正式论文语气，消除口语化、重复、层级混乱

### 4. 图片归集

- 从 `<proj>` 收集论文用到的图片，拷贝到 `figure/`
- 保留源文件名，引用路径统一为 `images/<filename>.png`

### 5. 参考文献提取

- 从 `<proj>` 中的论文 PDF 提取 BibTeX 条目，追加到 `src/references.bib`
- key 格式：`<第一作者姓氏><年份><关键词>`，如 `storn1997de`

### 6. 最终校验

运行 `make` 验证构建是否通过。若 pandoc 报引用缺失，补全 `references.bib`；若图片 404，检查路径。
