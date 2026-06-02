#!/usr/bin/env python3
"""
Markdown → Typst: pandoc + post-processing.

Key fix: strip manual numbering from headings so Typst's
`#set heading(numbering: "1.")` can auto-number them.

  = 第一章 绪论      →  = 绪论
  == 1.1 研究背景    →  == 研究背景
  === 1.2.1 子节     →  === 子节

Also fixes labels:
  - <slug> → <sec:slug>  on headings
  - tables → <tab:ch{N}-{i}>
  - figures → <fig:ch{N}-{i}>
  - display equations → <eq:ch{N}-{i}>

Usage:
    python md2typst.py chapter.md -o chapter.typ
    python md2typst.py thesis/ -o thesis-asm/src/main.typ --header header.typ
"""

import argparse
import re
import subprocess
import sys
from pathlib import Path


def pandoc_md_to_typst(md_path: Path) -> str:
    result = subprocess.run(
        ["pandoc", "-f", "markdown", "-t", "typst", str(md_path)],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(f"pandoc failed on {md_path}:\n{result.stderr}")
    return result.stdout


def _collect_block(lines: list[str], start: int) -> tuple[list[str], int]:
    block = [lines[start]]
    depth = 1
    i = start + 1
    while i < len(lines) and depth > 0:
        block.append(lines[i])
        depth += lines[i].count("(")
        depth -= lines[i].count(")")
        i += 1
    return block, i


def _strip_heading_number(title: str, level: int) -> str:
    """Remove manual Chinese/Arabic numbering from heading text."""
    if level == 1:
        # = 第一章 绪论  →  绪论
        m = re.match(r'^第[一二三四五六七八九十\d]+章\s+(.+)$', title)
        if m:
            return m.group(1)
        # = 1 Introduction / = 1. Introduction →  Introduction
        m = re.match(r'^\d+\.?\s+(.+)$', title)
        if m:
            return m.group(1)
    elif level == 2:
        # == 1.1 研究背景  →  研究背景
        m = re.match(r'^\d+\.\d+\s+(.+)$', title)
        if m:
            return m.group(1)
    elif level == 3:
        # === 1.2.1 子节  →  子节
        m = re.match(r'^\d+\.\d+\.\d+\s+(.+)$', title)
        if m:
            return m.group(1)
    return title


def postprocess(typst: str, chapter_no: int) -> str:
    # md 侧用 images/ 引用，typst 侧统一用 figure/
    typst = typst.replace('image("images/', 'image("figure/')
    lines = typst.splitlines()
    out: list[str] = []

    tab_i = fig_i = eq_i = 0
    i = 0

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # ---- Heading + label on next line ------------------------------
        hm = re.match(r'^(=+)\s+(.+)$', stripped)
        if hm and i + 1 < len(lines):
            nxt = lines[i + 1].strip()
            lm = re.match(r'^<([^>]+)>$', nxt)
            if lm:
                level = len(hm.group(1))
                title = _strip_heading_number(hm.group(2).strip(), level)
                label = lm.group(1)
                if not label.startswith(("sec:", "tab:", "fig:", "eq:")):
                    label = f"sec:{label}"
                out.append(f"{'=' * level} {title} <{label}>")
                i += 2
                continue

        # ---- Tables & Figures -----------------------------------------
        if stripped.startswith("#figure("):
            block, i = _collect_block(lines, i)
            block_text = "\n".join(block)
            if "kind: table" in block_text:
                tab_i += 1
                block[-1] = block[-1].rstrip()
                if block[-1].endswith(")"):
                    block[-1] = block[-1][:-1] + f")<tab:ch{chapter_no}-{tab_i}>"
            elif 'image("' in block_text:
                fig_i += 1
                block[-1] = block[-1].rstrip()
                if block[-1].endswith(")"):
                    block[-1] = block[-1][:-1] + f")<fig:ch{chapter_no}-{fig_i}>"
            out.extend(block)
            continue

        # ---- Display equations -----------------------------------------
        # Typst display math 以 "$ " 开头、" $" 结尾；行内公式无空格，避免误判
        if (
            stripped.startswith("$ ")
            and stripped.endswith(" $")
            and len(stripped) > 20
        ):
            eq_i += 1
            out.append(f"{stripped} <eq:ch{chapter_no}-{eq_i}>")
            i += 1
            continue

        out.append(line)
        i += 1

    return "\n".join(out)


def convert_file(md_path: Path, out_path: Path, chapter_no: int = 1):
    raw = pandoc_md_to_typst(md_path)
    fixed = postprocess(raw, chapter_no)
    out_path.write_text(fixed, encoding="utf-8")
    print(f"  [{chapter_no}] {md_path} → {out_path}")


def main():
    parser = argparse.ArgumentParser(description="Markdown → Typst via pandoc + post-process")
    parser.add_argument("input", help="Input .md file or directory")
    parser.add_argument("-o", "--output", help="Output .typ file or directory")
    parser.add_argument("--header", help="Typst header to prepend (combine mode)")
    parser.add_argument("--combine", action="store_true", help="Combine all .md into one .typ")
    args = parser.parse_args()

    inp = Path(args.input)

    if inp.is_file():
        out = Path(args.output) if args.output else inp.with_suffix(".typ")
        convert_file(inp, out, chapter_no=1)
        return

    if not inp.is_dir():
        sys.exit(f"Error: {inp} not found")

    md_files = sorted(inp.glob("chapter-*.md"))
    if not md_files:
        md_files = sorted(inp.glob("*.md"))
    if not md_files:
        sys.exit(f"Error: no .md files in {inp}")

    if args.combine:
        parts: list[str] = []
        if args.header:
            header_path = Path(args.header)
            if header_path.exists():
                parts.append(header_path.read_text(encoding="utf-8"))
            else:
                parts.append(args.header)  # treat as raw text
        for idx, f in enumerate(md_files, start=1):
            parts.append(postprocess(pandoc_md_to_typst(f), chapter_no=idx))

        # 尾部内容：致谢、附录
        extra_files = [
            ("acknowledgment.md", "acknowledgment"),
            ("appendix.md", "appendix"),
        ]
        for fname, label in extra_files:
            extra_path = inp / fname
            if extra_path.exists():
                raw = pandoc_md_to_typst(extra_path)
                typst = postprocess(raw, chapter_no=len(md_files) + 1)
                parts.append(f"{typst}")

        # 参考文献指令：告诉 pandoc citeproc 在致谢/附录之后填入文献列表
        parts.append("\n#bibliography(\"references\")\n")

        out = Path(args.output) if args.output else Path("combined.typ")
        text = parts[0] if parts else ""
        for i, p in enumerate(parts[1:]):
            sep = "\n\n#pagebreak()\n\n" if i > 0 else "\n\n"
            text += sep + p
        out.write_text(text, encoding="utf-8")
        print(f"Combined {len(md_files)} files → {out}")
    else:
        out_dir = Path(args.output) if args.output else inp
        out_dir.mkdir(parents=True, exist_ok=True)
        for idx, f in enumerate(md_files, start=1):
            convert_file(f, out_dir / f.with_suffix(".typ").name, chapter_no=idx)


if __name__ == "__main__":
    main()
