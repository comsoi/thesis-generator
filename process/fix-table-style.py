#!/usr/bin/env python3
"""
修复 DOCX 中表格样式引用不匹配的问题。

Pandoc 的 DOCX writer 在 custom-style 属性中写入的是样式“名称”，
而 Word 的 w:tblStyle 按样式 ID 解析。如果参考文档里三线表样式的 ID
不是该名称（例如 ID 为 "62"，名称为 "CenterTable"），Word 会回退到
默认样式，导致表格失去三线格效果。

本脚本：
1. 把 document.xml 中所有表格的 w:tblStyle 值从“名称”映射为对应的
   styles.xml 中的 table 样式 ID；
2. 删除 Pandoc 为 figure 生成的 FigureTable 外壳（只含一个嵌套表格的
   单单元格外壳），避免同一表格被渲染成两个嵌套表。
"""

import argparse
import os
import shutil
import tempfile
import zipfile
from lxml import etree

WPML = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W = '{%s}' % WPML


def build_table_style_map(styles_root):
    """name -> styleId for table styles."""
    name_to_id = {}
    for s in styles_root.iter(W + 'style'):
        if s.get(W + 'type') != 'table':
            continue
        style_id = s.get(W + 'styleId')
        name_el = s.find(W + 'name')
        if name_el is None:
            continue
        name = name_el.get(W + 'val')
        if name:
            name_to_id[name] = style_id
    return name_to_id


def is_figure_wrapper_table(tbl):
    """判断是否为 Pandoc 为 figure 生成的 FigureTable 外壳。"""
    pr = tbl.find(W + 'tblPr')
    if pr is None:
        return False
    ts = pr.find(W + 'tblStyle')
    if ts is None:
        return False
    if ts.get(W + 'val') not in ('FigureTable',):
        return False
    rows = list(tbl.iter(W + 'tr'))
    if len(rows) != 1:
        return False
    cells = list(rows[0].iter(W + 'tc'))
    if len(cells) != 1:
        return False
    # 统计该单元格内的直接表格数
    nested = [c for c in cells[0] if c.tag == W + 'tbl']
    if len(nested) != 1:
        return False
    return True


def fix_docx(input_path, output_path):
    # 如果原地修改，先写入临时文件再覆盖，避免 SameFileError
    if os.path.abspath(input_path) == os.path.abspath(output_path):
        fd, tmp_path = tempfile.mkstemp(suffix='.docx', prefix='fix-table-style-')
        os.close(fd)
        try:
            _do_fix(input_path, tmp_path)
            shutil.move(tmp_path, output_path)
        except Exception:
            os.remove(tmp_path)
            raise
    else:
        _do_fix(input_path, output_path)


def _do_fix(input_path, output_path):
    shutil.copy(input_path, output_path)

    with zipfile.ZipFile(output_path, 'r') as zin:
        styles_root = etree.fromstring(zin.read('word/styles.xml'))
        doc_root = etree.fromstring(zin.read('word/document.xml'))

    name_to_id = build_table_style_map(styles_root)

    body = doc_root.find(W + 'body')
    for tbl in list(body.iter(W + 'tbl')):
        # 1. 名称 -> ID 映射
        pr = tbl.find(W + 'tblPr')
        if pr is not None:
            ts = pr.find(W + 'tblStyle')
            if ts is not None:
                val = ts.get(W + 'val')
                if val in name_to_id and val != name_to_id[val]:
                    ts.set(W + 'val', name_to_id[val])

        # 2. 删除 FigureTable 外壳，把内层表格提升到 body
        if is_figure_wrapper_table(tbl):
            cell = next(tbl.iter(W + 'tc'))
            inner_tbl = next((c for c in cell if c.tag == W + 'tbl'), None)
            if inner_tbl is not None:
                tbl.addprevious(inner_tbl)
                tbl.getparent().remove(tbl)

    # 重写 zip
    with zipfile.ZipFile(output_path, 'r') as zin:
        items = [(n, zin.read(n)) for n in zin.namelist()]
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n, data in items:
            if n == 'word/document.xml':
                data = etree.tostring(doc_root, xml_declaration=True, encoding='UTF-8', standalone=True)
            zout.writestr(n, data)


def main():
    parser = argparse.ArgumentParser(description='Fix DOCX table style references')
    parser.add_argument('input', help='input .docx')
    parser.add_argument('output', nargs='?', help='output .docx (default overwrite input)')
    args = parser.parse_args()
    output = args.output or args.input
    fix_docx(args.input, output)
    print(f'fixed table styles: {args.input} -> {output}')


if __name__ == '__main__':
    main()
