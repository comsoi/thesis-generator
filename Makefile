.PHONY: all gen-src content content-cn content-en extend-content extend-content-cn extend-content-en gen-prelim gen-prelim-cn gen-prelim-en merge-docs merge-docs-cn merge-docs-en gen-flowchart copy-images sync-config clean clean-config help

# 目录
SRC_DIR := ./src
MD_DIR := ./markdown
TPL_DIR := ./template
PROC_DIR := ./process
OUT_DIR := ./result
FLOWCHART_DIR := ./flowchart
FIG_DIR := ./figure
CACHE_DIR := ./.cache

# 构建工具
PYTHON := uv run

# 输出文件
MAIN_CONTENT_DOCX := $(OUT_DIR)/main_content.docx
MAIN_CONTENT_DOCX_CN := $(OUT_DIR)/main_content_cn.docx
EXTENDED_CONTENT_DOCX := $(OUT_DIR)/extended_content.docx
EXTENDED_CONTENT_DOCX_CN := $(OUT_DIR)/extended_content_cn.docx
PRELIM_CN := $(OUT_DIR)/prelim-cn.docx
PRELIM_EN := $(OUT_DIR)/prelim-en.docx
MERGED_CN := $(OUT_DIR)/thesis-cn.docx
MERGED_EN := $(OUT_DIR)/thesis-en.docx
TYPST_FILE := $(SRC_DIR)/main.typ

# 源文件
MD_FILES := $(wildcard $(MD_DIR)/*.md)
MD_ABSTRACT_CN := $(MD_DIR)/abstract-cn.md
MD_ABSTRACT_EN := $(MD_DIR)/abstract-en.md
HEADER_FILE := $(SRC_DIR)/header.typ
BIB_FILE := $(SRC_DIR)/references.bib
CSL_CN := $(TPL_DIR)/gb.csl
CSL_EN := $(TPL_DIR)/ieee.csl
REFERENCE_DOC_EN := $(TPL_DIR)/clean_ref_en.docx
REFERENCE_DOC_CN := $(TPL_DIR)/clean_ref_cn.docx
PRELIM_TEMPLATE_CN := $(TPL_DIR)/prelim-cn.docx
PRELIM_TEMPLATE_EN := $(TPL_DIR)/prelim-en.docx
CONFIG := ./config.json

# 脚本
MD_SCRIPT := $(PROC_DIR)/md2typst.py
TOC_SCRIPT := $(PROC_DIR)/gen-toc.py
RENDER_SCRIPT := $(PROC_DIR)/substitute-txt.py
SYNC_SCRIPT := $(PROC_DIR)/sync-abstract.py
MERGE_SCRIPT := $(PROC_DIR)/merge-docs.py
SETUP_SCRIPT := $(PROC_DIR)/setup-thesis.py
TABLE_STYLE_FIX_SCRIPT := $(PROC_DIR)/fix-table-style.py

# Lua 过滤器
LUA_FIX_REF := $(PROC_DIR)/fix-ref.lua
LUA_TABLE_CAPTION := $(PROC_DIR)/fix-table-caption.lua
LUA_FIX_PAGE := $(PROC_DIR)/fix-pagebrk.lua
LUA_FIX_CITE := $(PROC_DIR)/fix-cite-link.lua

# 其它
XELATEX := xelatex
TEX_FILES := $(wildcard $(FLOWCHART_DIR)/*.tex)
GEN_FLOWCHART_STAMP := $(CACHE_DIR)/.gen-flowchart.stamp

all: merge-docs-cn merge-docs-en

# --- 图片生成 ---

$(GEN_FLOWCHART_STAMP): $(TEX_FILES)
	mkdir -p $(FIG_DIR) $(CACHE_DIR)
	@for f in $(TEX_FILES); do \
		name=$$(basename "$$f" .tex); \
		echo "  XeLaTeX  $$f -> $(FIG_DIR)/$$name.png"; \
		cp "$$f" $(CACHE_DIR)/$$name.tex; \
		( cd $(CACHE_DIR) && $(XELATEX) -shell-escape \
			-interaction=nonstopmode -halt-on-error \
			$$name.tex > /dev/null 2>&1 ); \
		mv $(CACHE_DIR)/$$name-1.png $(FIG_DIR)/$$name.png 2>/dev/null; \
		rm -f $(CACHE_DIR)/$$name.* $(CACHE_DIR)/$$name-*.png; \
		done
	@touch $@

gen-flowchart: $(GEN_FLOWCHART_STAMP)

# --- 外部图片拷贝 ---

copy-images:
	mkdir -p $(FIG_DIR)
	cp ../outputs/*.png $(FIG_DIR)/ 2>/dev/null || true

# --- Markdown → Typst ---

$(TYPST_FILE): $(HEADER_FILE) $(MD_FILES) $(GEN_FLOWCHART_STAMP)
	$(PYTHON) $(MD_SCRIPT) $(MD_DIR) -o $@ --header $(HEADER_FILE) --combine

gen-src: $(TYPST_FILE)

# --- Typst → DOCX (正文) ---

$(MAIN_CONTENT_DOCX): $(TYPST_FILE) $(REFERENCE_DOC_EN)
	mkdir -p $(OUT_DIR)
	pandoc $< --resource-path=$(SRC_DIR) --bibliography=$(BIB_FILE) --csl=$(CSL_EN) --lua-filter=$(LUA_FIX_REF) --lua-filter=$(LUA_TABLE_CAPTION) --lua-filter=$(LUA_FIX_PAGE) --citeproc --lua-filter=$(LUA_FIX_CITE) -o $@ --reference-doc $(REFERENCE_DOC_EN)
	$(PYTHON) $(TABLE_STYLE_FIX_SCRIPT) $@

$(MAIN_CONTENT_DOCX_CN): $(TYPST_FILE) $(REFERENCE_DOC_CN)
	mkdir -p $(OUT_DIR)
	pandoc $< --resource-path=$(SRC_DIR) --bibliography=$(BIB_FILE) --csl=$(CSL_CN) --lua-filter=$(LUA_FIX_REF) --lua-filter=$(LUA_TABLE_CAPTION) --lua-filter=$(LUA_FIX_PAGE) --citeproc --lua-filter=$(LUA_FIX_CITE) -o $@ --reference-doc $(REFERENCE_DOC_CN)
	$(PYTHON) $(TABLE_STYLE_FIX_SCRIPT) $@

content: $(MAIN_CONTENT_DOCX)
content-cn: $(MAIN_CONTENT_DOCX_CN)

# --- 扩展内容（目录等） ---

$(EXTENDED_CONTENT_DOCX): $(MAIN_CONTENT_DOCX)
	mkdir -p $(OUT_DIR)
	$(PYTHON) $(TOC_SCRIPT) $< $@

$(EXTENDED_CONTENT_DOCX_CN): $(MAIN_CONTENT_DOCX_CN)
	mkdir -p $(OUT_DIR)
	$(PYTHON) $(TOC_SCRIPT) $< $@

extend-content: $(EXTENDED_CONTENT_DOCX)
extend-content-cn: $(EXTENDED_CONTENT_DOCX_CN)

# --- 同步摘要 ---
sync-config:
	$(PYTHON) $(SYNC_SCRIPT) $(CONFIG) $(MD_ABSTRACT_CN) $(MD_ABSTRACT_EN)

# --- 前置部分 ---

$(PRELIM_CN): $(PRELIM_TEMPLATE_CN) $(CONFIG)
	mkdir -p $(OUT_DIR)
	$(PYTHON) $(RENDER_SCRIPT) $(PRELIM_TEMPLATE_CN) $(CONFIG) $@

$(PRELIM_EN): $(PRELIM_TEMPLATE_EN) $(CONFIG)
	mkdir -p $(OUT_DIR)
	$(PYTHON) $(RENDER_SCRIPT) $(PRELIM_TEMPLATE_EN) $(CONFIG) $@

gen-prelim-cn: $(PRELIM_CN)
gen-prelim-en: $(PRELIM_EN)
gen-prelim: gen-prelim-cn gen-prelim-en

# --- 合并：封面 + 正文 ---

$(MERGED_CN): $(PRELIM_CN) $(EXTENDED_CONTENT_DOCX_CN)
	mkdir -p $(OUT_DIR)
	$(PYTHON) $(MERGE_SCRIPT) $@ $(PRELIM_CN) $(EXTENDED_CONTENT_DOCX_CN)
	$(PYTHON) $(SETUP_SCRIPT) $@ --split-heading 1 --pg-start=-1,1 --merge 1-2 --clear-headers 1

$(MERGED_EN): $(PRELIM_EN) $(EXTENDED_CONTENT_DOCX)
	mkdir -p $(OUT_DIR)
	$(PYTHON) $(MERGE_SCRIPT) $@ $(PRELIM_EN) $(EXTENDED_CONTENT_DOCX)
	$(PYTHON) $(SETUP_SCRIPT) $@ --split-heading 1 --pg-start=-1,1 --merge 1-2 --clear-headers 1

merge-docs-cn: $(MERGED_CN)
merge-docs-en: $(MERGED_EN)
merge-docs: merge-docs-cn merge-docs-en


# --- 清理 ---

clean:
	@if [ -d $(OUT_DIR) ]; then \
		find $(OUT_DIR) -mindepth 1 -delete; \
	fi
	@rm -rf $(CACHE_DIR)

clean-config:
	@jq '.ABSTRACTCN = "" | .ABSTRACTEN = "" | .KWDCN = "" | .KWDEN = ""' $(CONFIG) > $(CONFIG).tmp && mv $(CONFIG).tmp $(CONFIG)

# --- 帮助 ---

help:
	@echo "make targets:"
	@echo ""
	@echo "  make all           完整构建 → result/thesis-cn.docx + result/thesis-en.docx"
	@echo "  make merge-docs    合并中英文版（merge-docs-cn + merge-docs-en）"
	@echo "  make merge-docs-cn 合并中文版（中文前置 + 中文格式正文）"
	@echo "  make merge-docs-en 合并英文版（英文前置 + 英文格式正文）"
	@echo ""
	@echo "  分步构建:"
	@echo "  make gen-flowchart flowchart/ → figure/   XeLaTeX 流程图 → PNG"
	@echo "  make copy-images   ../outputs/*.png → figure/   拷贝外部产出图"
	@echo "  make gen-src       markdown/*.md → src/main.typ  Markdown → Typst"
	@echo "  make content       src/main.typ → result/main_content.docx  英文正文排版"
	@echo "  make content-cn    src/main.typ → result/main_content_cn.docx  中文正文排版"
	@echo "  make extend-content   result/main_content.docx → result/extended_content.docx"
	@echo "  make extend-content-cn result/main_content_cn.docx → result/extended_content_cn.docx"
	@echo "  make gen-prelim     生成中英文前置部分（gen-prelim-cn + gen-prelim-en）"
	@echo "  make gen-prelim-cn  中文前置 → result/prelim-cn.docx"
	@echo "  make gen-prelim-en  英文前置 → result/prelim-en.docx"
	@echo ""
	@echo "  辅助:"
	@echo "  make sync-config   同步摘要 (edit markdown/abstract-*.md 后自动触发)"
	@echo "  make clean         清空 result/ 和 .cache/"
	@echo "  make clean-config  重置 config.json 摘要字段"
	@echo ""
	@echo "  目录约定:"
	@echo "  figure/                   图片唯一存储"
	@echo "  markdown/images/ → ../figure   md 侧引用（软链接）"
	@echo "  src/figure/       → ../figure   typst 侧引用（软链接）"
	@echo "  flowchart/                流程图源文件"
	@echo "  .cache/                   构建戳记"
	@echo ""
	@echo "  参考文档切换:"
	@echo "  英文正文: --csl=ieee.csl  --reference-doc=clean_ref_en.docx"
	@echo "  中文正文: --csl=gb.csl   --reference-doc=clean_ref_cn.docx"
