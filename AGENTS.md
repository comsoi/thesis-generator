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
