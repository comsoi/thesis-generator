# Thesis Assembler (Markdown → Word)

用 Markdown 撰写论文，经 Typst 排版引擎，最终输出带样式的 Word 学位论文文档。

## 流水线概要

1. 编写 Markdown 正文（`markdown/*.md`）和元数据（`config.json`）
2. `md2typst.py` 将 Markdown 转为 Typst（`src/main.typ`）
3. Pandoc + Lua 过滤器将 Typst 转为 Word 正文 DOCX
4. 注入目录、图索引、表索引域
5. 基于 Word 模板生成封面
6. 合并封面与正文

## 目录结构

### 根目录
- `Makefile`：构建入口，定义完整文档生成流程
- `config.json`：封面元数据及摘要/关键词，供模板渲染使用
- `README.md`：使用说明与架构文档

### 内容
- `src/main.typ`：Typst 主源文件（章节、图表、公式、参考文献）
- `src/header.typ`：Typst 头文件（`#set` 规则等）
- `markdown/`：Markdown 源文件，经由 `process/md2typst.py` 转换为 `src/main.typ`
- `src/references.bib`：BibTeX 参考文献

### 模板
- `template/gb.csl`：国标参考文献格式（中文正文使用）
- `template/ieee.csl`：IEEE 参考文献格式（英文正文使用）
- `template/clean_ref_en.docx`：英文 Pandoc 样式参考文档（`make content` / `make merge-docs-en`）
- `template/clean_ref_cn.docx`：中文 Pandoc 样式参考文档（`make content-cn` / `make merge-docs-cn`）
- `template/prelim-cn.docx`：中文前置部分模板（原创性声明、摘要等）
- `template/prelim-en.docx`：英文前置部分模板

### 处理脚本 (`process/`)

| 文件 | 作用 |
|---|---|
| `fix-ref.lua` | 交叉引用：`@tab:`/`@fig:`/`@sec:`/`@eq:` → Word REF 域，图表公式按章编号 |
| `fix-table-caption.lua` | 表题上移：将标题段落置于表格之前 |
| `fix-pagebrk.lua` | 分页符：`#pagebreak()` → Word 分页符 |
| `fix-cite-link.lua` | 文献引用：`[@citekey]` → 可点击跳转的数值上标 |

- `md2typst.py`：Markdown → Typst（剥离手动编号、补全标签前缀、路径替换）
- `gen-toc.py`：向正文 DOCX 注入目录、图索引、表索引域
- `substitute-txt.py`：用 `config.json` 渲染前置部分模板（`prelim-cn/en.docx`）
- `merge-docs.py`：合并封面与正文 DOCX
- `setup-thesis.py`：分节、页码、页眉清理等后处理
- `sync-abstract.py`：同步 `abstract-*.md` 到 `config.json`

### 输出
- `result/`：中间及最终产物
  - `main_content.docx` — 正文
  - `extended_content.docx` — 正文 + 目录
  - `prelim-cn.docx` / `prelim-en.docx` — 中/英文前置部分
  - `thesis-cn.docx` / `thesis-en.docx` — 最终合并文档

## Markdown 写作规范

`markdown/` 目录下的所有 `.md` 文件需遵循以下格式约定，以确保 `md2typst.py` 正确处理并生成合规的 Typst 源文件。

### 章节标题

使用 `#` / `##` / `###` 表示一级/二级/三级标题，**不要手动编号**。

```markdown
# Introduction

## Research Background

### Subsection Title
```

`md2typst.py` 会自动剥离形如 `第一章`、`1.1` 的手动编号，交由 Typst 的 `#set heading(numbering: "1.")` 自动编号。手动编号会导致双重编号或格式错误。

### 图片引用

使用标准 Markdown 图片语法，路径前缀必须是 `images/`：

```markdown
![caption](images/figure-name.png)
```

`md2typst.py` 会将 `images/` 替换为 `figure/`，以匹配 Typst 侧的实际图片目录。请勿使用 `figure/` 或其他路径前缀。

### 表格

标准 Markdown 表格，表题用 `: 标题` 紧邻表格上方或下方（中间不留空行）：

```markdown
| Column A | Column B |
|----------|----------|
| value1   | value2   |

: 表格标题
```

`fix-table-caption.lua` 会将表题提取为独立段落置于表格之前。

### 公式

- 行内公式：用 `$...$` 包裹
- 行间公式：用 `$$...$$` 包裹

```markdown
行内：$E = mc^2$

行间：

$$\hat{y}_i = \sum_{k=1}^{K} f_k(x_i)$$
```

行间公式会被 `md2typst.py` 自动分配 `<eq:ch{N}-{i}>` 标签，用于交叉引用。

### 文献引用

- 括号式引用：`[@citekey]`
- 叙述式引用：`Author[@citekey]`

```markdown
Pandoc 支持 Markdown 到 40 余种格式的转换[@macfarlane2022pandoc]。
```

`--citeproc` 将其渲染为数值上标，`fix-cite-link.lua` 添加超链接。

### 交叉引用

在正文中直接使用 `@` 前缀引用：

```markdown
参见 @sec:研究背景、@eq:ch3-1 和 @tab:ch3-1。
```

标签前缀由 `md2typst.py` 自动分配，通常**不需要手动编写**：

| 类型 | 前缀 | 示例 |
|---|---|---|
| 章节 | `<sec:...>` | `<sec:研究背景>` |
| 公式 | `<eq:...>` | `<eq:ch3-1>` |
| 图 | `<fig:...>` | `<fig:ch3-1>` |
| 表 | `<tab:...>` | `<tab:ch3-1>` |

### 摘要格式

**中文摘要**（`markdown/abstract-cn.md`）：

```markdown
# 摘要

正文段落……

**关键词**：关键词1；关键词2
```

**英文摘要**（`markdown/abstract-en.md`）：

```markdown
# Abstract

正文段落……

**Keywords**: keyword1; keyword2
```

`sync-abstract.py` 会自动提取摘要正文和关键词，同步到 `config.json` 的 `ABSTRACTCN` / `ABSTRACTEN` / `KWDCN` / `KWDEN` 字段，供封面模板渲染使用。关键词之间用中文分号 `；`（中文摘要）或英文分号 `;`（英文摘要）分隔。

## 使用方法

### 1. 环境准备

```bash
# pandoc
brew install pandoc
# Python 依赖
python3 -m pip install python-docx docxtpl
```

### 2. 编辑内容

**主要写作方式：Markdown**（`markdown/*.md`）。`make gen-src` 会自动将 Markdown 转为 Typst。也可以直接编辑 `src/main.typ`。

- **章节课文**：`markdown/chapter-1-introduction.md` ~ `chapter-5-conclusion.md`
- **致谢**：`markdown/acknowledgment.md`
- **附录**：`markdown/appendix.md`
- **摘要**：`markdown/abstract-cn.md`、`markdown/abstract-en.md`（保存后自动同步到 `config.json`）
- **元数据**：编辑 `config.json`（`TITLE_CN`、`TITLE_EN`、`SID`、`STNAME`、`INSTRUCTOR`、`MAJOR`、`SUBJECT`、`CNDATE`）
- **参考文献**：编辑 `src/references.bib`（正文中用 `Author[@citekey]` 引用，citeproc 自动生成文献列表）

### 3. 构建

```bash
make          # 完整构建 → result/thesis-cn.docx + result/thesis-en.docx
make all      # 同上
```

分步构建：

```bash
make gen-src           # markdown/*.md → src/main.typ
make content           # src/main.typ → result/main_content.docx（英文格式）
make content-cn        # src/main.typ → result/main_content_cn.docx（中文格式）
make extend-content    # 注入目录 → result/extended_content.docx
make extend-content-cn # 注入目录 → result/extended_content_cn.docx
make gen-prelim        # 生成中英文前置部分
make merge-docs        # 合并 → result/thesis-cn.docx + result/thesis-en.docx
make clean             # 清空 result/ 和 .cache/
make clean-config      # 重置 config.json 摘要字段
```

**中英文格式切换**：`make content` / `make merge-docs-en` 使用 `template/clean_ref_en.docx`（英文排版样式）；`make content-cn` / `make merge-docs-cn` 使用 `template/clean_ref_cn.docx`（中文排版样式，如中文字体、段落缩进等）。

### 4. 注意事项

- **首次打开必须更新域**：打开 `result/thesis-cn.docx` 或 `result/thesis-en.docx`，全选（`Ctrl+A`）按 `F9`，勾选"更新整个目录"。目录、图索引、表索引、交叉引用编号才能正确显示。
- 合并步骤不依赖 Microsoft Word，基于 `python-docx` 纯 Python 实现。

## 构建流程图

```
markdown/*.md (chapters + acknowledgment + appendix)
  → md2typst.py → src/main.typ (+ #bibliography())
  → pandoc + fix-ref.lua + fix-table-caption.lua + fix-pagebrk.lua + citeproc + fix-cite-link.lua
    --reference-doc template/clean_ref_en.docx
  → result/main_content.docx
  → gen-toc.py
  → result/extended_content.docx

  → pandoc (同上过滤器)
    --reference-doc template/clean_ref_cn.docx
  → result/main_content_cn.docx
  → gen-toc.py
  → result/extended_content_cn.docx

template/prelim-cn.docx + config.json
  → substitute-txt.py → result/prelim-cn.docx

template/prelim-en.docx + config.json
  → substitute-txt.py → result/prelim-en.docx

prelim-cn.docx + extended_content_cn.docx
  → merge-docs.py → setup-thesis.py → result/thesis-cn.docx

prelim-en.docx + extended_content.docx
  → merge-docs.py → setup-thesis.py → result/thesis-en.docx
```
