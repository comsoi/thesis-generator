#!/usr/bin/env python3
"""
论文 DOCX 后处理：节操作与页码控制。

功能：
    - 在指定位置插入分节符（按标题级别、页码）
    - 合并相邻节
    - 设置每节的起始页码和页码格式
    - 清除指定节的页眉/页脚

用法:
    python setup-thesis.py <input.docx> [output.docx] [选项]

选项:
    --split-heading N     在第 N 个顶级标题前插入分节符 (默认 N=1)
                          顶级标题 = outlineLvl=0 或 style name 含 "heading 1"
                          例如：N=1 在"第一章 绪论"前分节
    --split-page N        在物理第 N 页处插入分节符 (需 Word 渲染过)
                          依赖 Word 保存时写入的 <w:lastRenderedPageBreak> 标记
    --merge I-J[-np]      合并第 I 到第 J 节 (0-based)
                          默认插入分页符；加 -np 后缀则不插入
    --pg-start S,N        设置第 S 节起始页码为 N
    --pg-format S,FMT     设置第 S 节页码格式
                          decimal / upperRoman / lowerRoman / none
                          none = 从页眉页脚中移除 PAGE 域
    --clear-footers I,J   清除指定节的页脚内容
    --clear-headers I,J   清除指定节的页眉内容

默认行为（无选项时）：
    1. 在第 1 个顶级标题前插入 nextPage 分节符
    2. 合并 Section 1 和 2（带分页符）
    3. 最后一节页码 start=1
    4. Section 1 无页码（none）
"""

import sys
import os
import re
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree as _etree
from docx.opc.part import Part
from docx.opc.packuri import PackURI
from docx.opc.constants import RELATIONSHIP_TYPE as RT


def _resolve_si(doc, raw):
    """解析节索引，-1 表示最后一节。"""
    si = int(raw)
    return len(doc.sections) + si if raw.startswith('-') else si


# 操作分发表：name → handler(doc, value)
DISPATCH = {
    'split_heading':  lambda d, v: split_before_heading(d, int(v)),
    'split_page':     lambda d, v: split_at_page(d, int(v)),
    'merge':          lambda d, v: _dispatch_merge(d, v),
    'pg_start':       lambda d, v: _dispatch_pg_start(d, v),
    'pg_format':      lambda d, v: _dispatch_pg_format(d, v),
    'clear_footers':  lambda d, v: clear_section_footers(d, *[int(x) for x in v.split(',')]),
    'clear_props':    lambda d, v: clear_doc_properties(d),
    'clear_headers':  lambda d, v: clear_section_headers(d, *[int(x) for x in v.split(',')]),
}


def _dispatch_merge(doc, v):
    page_break = not v.endswith('-np')
    if v.endswith('-np'):
        v = v[:-3]
    start, _, end = v.partition('-')
    merge_sections(doc, int(start), int(end) if end else None, page_break=page_break)


def _dispatch_pg_start(doc, v):
    si, val = v.split(',')
    set_page_start(doc, _resolve_si(doc, si), int(val))


def _dispatch_pg_format(doc, v):
    si, fmt = v.split(',')
    set_page_format(doc, _resolve_si(doc, si), fmt.strip())


def main(doc_path, output_path=None, actions=None):
    """按 actions 列表顺序依次执行操作。"""
    if actions is None:
        actions = []
    doc = Document(doc_path)

    for name, value in actions:
        handler = DISPATCH.get(name)
        if handler:
            handler(doc, value)

    output = output_path or doc_path
    os.makedirs(os.path.dirname(output) or '.', exist_ok=True)
    doc.save(output)
    print(f"setup-thesis: {doc_path} → {output}")


# =========================================================
# 节操作（分节、合节）
# =========================================================

def split_before_heading(doc, n=1):
    """在第 n 个顶级标题段落前插入 nextPage 分节符。

    标题判定规则（_is_top_heading）：
    - 段落的 style 在 styles.xml 中 outlineLvl=0，或
    - style 的 name 含 "heading 1"（大小写不敏感）

    Pandoc 生成的文档用 styleId=2 (name="heading 1", outlineLvl=0) 表示
    一级标题，Word 原生文档则用 styleId="Heading1"。本函数兼容两种命名方式。

    分节符附加在标题段落的**前一个段落**上（Word 约定：段落级 sectPr 定义
    该段之上内容的分节属性）。
    """
    body = doc.element.body
    found = 0
    for p in body.iter(qn('w:p')):
        pPr = p.find(qn('w:pPr'))
        if pPr is None:
            continue
        pStyle = pPr.find(qn('w:pStyle'))
        if pStyle is None:
            continue
        if not _is_top_heading(doc, pStyle.get(qn('w:val')) or ''):
            continue

        found += 1
        if found != n:
            continue

        # 找到标题的前一个段落元素
        prev = p.getprevious()
        while prev is not None and prev.tag != qn('w:p'):
            prev = prev.getprevious()
        if prev is None:
            return

        _insert_section_break(doc, prev, 'nextPage')
        return


def _is_top_heading(doc, style_id):
    """判断 styleId 是否为顶级标题。

    两种判定方式（满足其一即可）：
    1. outlineLvl=0：Word 标准的"第 1 级大纲"（Pandoc 生成的 heading 1 用此标记）
    2. style name 含 "heading 1"：Word 原生 Heading 1 样式
    """
    for s in doc.styles.element:
        if s.tag == qn('w:style') and s.get(qn('w:styleId')) == style_id:
            pPr = s.find(qn('w:pPr'))
            if pPr is not None:
                ol = pPr.find(qn('w:outlineLvl'))
                if ol is not None and ol.get(qn('w:val')) == '0':
                    return True
            name_el = s.find(qn('w:name'))
            if name_el is not None:
                name = (name_el.get(qn('w:val')) or '').lower()
                if 'heading 1' in name:
                    return True
            break
    return False


def split_at_page(doc, page_num):
    """在第 page_num 页的起始处插入分节符。

    前提：文档需先在 Word 中打开并保存，让 Word 写入 lastRenderedPageBreak
    标记。Pandoc 直接生成的 DOCX 不含此标记，无法按页分节。

    lastRenderedPageBreak 是一个 <w:br w:type="page"/> 的替代标记，
    存储在 <w:r> 内，表示 Word 渲染时该位置发生了分页。
    第 N 页的起始 = 第 N-1 个 lastRenderedPageBreak 所在段落。
    """
    body = doc.element.body
    page_count = 0
    for p in body.iter(qn('w:p')):
        for r in p.iter(qn('w:r')):
            if r.find(qn('w:lastRenderedPageBreak')) is not None:
                page_count += 1
                if page_count == page_num - 1:
                    _insert_section_break(doc, p, 'nextPage')
                    return


def merge_sections(doc, start_section, end_section=None, page_break=True):
    """合并节：Section start_section ~ end_section 合为一节。

    page_break=True（默认）：删除 sectPr 前插入 <w:br w:type="page"/>，
    保证合并后换页效果仍在。False：仅删除 sectPr，内容连续。

    CLI: --merge 1-2（默认分页） / --merge 1-2-np（无分页）
    """
    if end_section is None:
        end_section = len(doc.sections) - 1
    remove_from = start_section
    remove_to = end_section - 1

    body = doc.element.body
    idx = 0
    for p in body.findall(qn('w:p')):
        pPr = p.find(qn('w:pPr'))
        if pPr is None:
            continue
        sectPr = pPr.find(qn('w:sectPr'))
        if sectPr is None:
            continue
        if remove_from <= idx <= remove_to:
            if page_break:
                pb_p = OxmlElement('w:p')
                r = OxmlElement('w:r')
                br = OxmlElement('w:br')
                br.set(qn('w:type'), 'page')
                r.append(br)
                pb_p.append(r)
                p.addnext(pb_p)
            pPr.remove(sectPr)
        idx += 1
        if idx > remove_to:
            break


def _insert_section_break(doc, paragraph, break_type='nextPage'):
    """在指定段落的 pPr 中附加一个分节符，从 body-level sectPr 继承页眉/页脚。

    break_type: 'nextPage'（默认）| 'oddPage' | 'evenPage' | 'continuous'
    """
    from copy import deepcopy

    pPr = paragraph.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        paragraph.insert(0, pPr)
    sectPr = OxmlElement('w:sectPr')
    type_el = OxmlElement('w:type')
    type_el.set(qn('w:val'), break_type)
    sectPr.append(type_el)

    # 从 body-level sectPr 拷贝页眉/页脚引用，保证分节后新节仍有页眉
    body_sectPr = doc.element.body.find(qn('w:sectPr'))
    if body_sectPr is not None:
        for child in body_sectPr:
            if child.tag in (qn('w:headerReference'), qn('w:footerReference'),
                             qn('w:pgSz'), qn('w:pgMar'), qn('w:cols'), qn('w:docGrid')):
                sectPr.append(deepcopy(child))

    pPr.append(sectPr)



def insert_page_breaks_before_headings(doc):
    """在每个顶级标题（除第一个外）前插入手动分页符。

    适用于各章合并在同一节时，保证每章从新页开始。
    """
    found = 0
    for p in doc.element.body.iter(qn('w:p')):
        pPr = p.find(qn('w:pPr'))
        if pPr is None:
            continue
        pStyle = pPr.find(qn('w:pStyle'))
        if pStyle is None:
            continue
        if not _is_top_heading(doc, pStyle.get(qn('w:val')) or ''):
            continue

        found += 1
        if found == 1:
            continue  # 第一个标题跳过（已有分节符）

        prev = p.getprevious()
        while prev is not None and prev.tag != qn('w:p'):
            prev = prev.getprevious()
        if prev is None:
            continue
        run = OxmlElement('w:r')
        br = OxmlElement('w:br')
        br.set(qn('w:type'), 'page')
        run.append(br)
        prev.append(run)


# =========================================================
# 页码操作
# =========================================================

def set_page_start(doc, section_idx, start):
    """设置第 section_idx 节的起始页码。

    操作 sectPr 中的 <w:pgNumType w:start="...">。
    """
    sections = doc.sections
    if section_idx >= len(sections):
        return
    _set_pg_num_type(sections[section_idx]._sectPr, start=start)


def set_page_format(doc, section_idx, fmt):
    """设置第 section_idx 节的页码格式。

    fmt 取值：
    - decimal / upperRoman / lowerRoman：Word 标准页码格式
    - none：删除该节页眉/页脚中的 PAGE 域，实现"本节无页码"
    """
    sections = doc.sections
    if section_idx >= len(sections):
        return
    if fmt == 'none':
        _remove_page_fields_from_section(doc, section_idx)
    else:
        _set_pg_num_type(sections[section_idx]._sectPr, fmt=fmt)


def clear_page_numbers(doc, section_idx):
    """清除指定节的页码（等价于 set_page_format(section_idx, 'none')）。"""
    set_page_format(doc, section_idx, 'none')


def _set_pg_num_type(sectPr, start=None, fmt=None):
    """操作 sectPr 中的 <w:pgNumType> 元素。

    设置 start（起始页码）和 fmt（页码格式：decimal/upperRoman/lowerRoman）。
    """
    pgNumType = sectPr.find(qn('w:pgNumType'))
    if pgNumType is None:
        pgNumType = OxmlElement('w:pgNumType')
        sectPr.insert(0, pgNumType)
    if start is not None:
        pgNumType.set(qn('w:start'), str(start))
    if fmt is not None:
        pgNumType.set(qn('w:fmt'), fmt)


def _remove_page_fields_from_section(doc, section_idx):
    """从指定节的所有页眉/页脚中移除 PAGE 域。

    如果页眉/页脚 part 被多节共享，先 clone 一份独立的给当前节，
    避免影响其他节的页码显示。
    """
    sections = doc.sections
    sec = sections[section_idx]

    for ref_tag in (qn('w:headerReference'), qn('w:footerReference')):
        for ref in sec._sectPr.findall(ref_tag):
            rId = ref.get(qn('r:id'))
            if not (rId and rId in doc.part.rels):
                continue
            if _count_sections_using_rId(sections, rId) > 1:
                rId = _clone_hdr_ftr_part(doc, rId)
                ref.set(qn('r:id'), rId)
            _remove_page_fields(doc.part.rels[rId].target_part._element)


def _count_sections_using_rId(sections, target_rId):
    """统计有多少节引用了相同的 rId。"""
    count = 0
    for sec in sections:
        for ref_tag in (qn('w:headerReference'), qn('w:footerReference')):
            for ref in sec._sectPr.findall(ref_tag):
                if ref.get(qn('r:id')) == target_rId:
                    count += 1
    return count


def _clone_hdr_ftr_part(doc, original_rId):
    """克隆页眉/页脚 part 到包中，返回新 rId。"""
    from lxml import etree as _etree
    from docx.opc.part import XmlPart as _XmlPart
    from docx.opc.packuri import PackURI as _PackURI

    orig_part = doc.part.rels[original_rId].target_part
    reltype = doc.part.rels[original_rId].reltype

    key = 'header' if 'header' in reltype else 'footer'
    ct = f'application/vnd.openxmlformats-officedocument.wordprocessingml.{key}+xml'

    elem = _etree.fromstring(orig_part._element.xml.encode('utf-8'))
    existing = {p.partname for p in doc.part.package.parts}
    i = 1
    while _PackURI(f'/word/{key}{i}.xml') in existing:
        i += 1

    new_part = _XmlPart(
        partname=_PackURI(f'/word/{key}{i}.xml'),
        content_type=ct,
        element=elem,
        package=doc.part.package,
    )
    j = 1
    while f'rId{j}' in doc.part.rels:
        j += 1
    new_rId = f'rId{j}'
    doc.part.rels.add_relationship(reltype, new_part, new_rId)
    return new_rId


def _remove_page_fields(element):
    """从页眉或页脚元素中移除所有 PAGE 域。

    Word 的 PAGE 域结构（在 OOXML 中）：
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText> PAGE </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>当前页码</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>

    本函数找到 begin→end 之间包含 PAGE 指令的连续 run，全部删除。
    """
    for para in element.iter(qn('w:p')):
        runs = list(para.iter(qn('w:r')))
        i = 0
        while i < len(runs):
            run = runs[i]
            fldChar = run.find(qn('w:fldChar'))
            if fldChar is not None and fldChar.get(qn('w:fldCharType')) == 'begin':
                # 找到一个域的开始，收集连续 run 直到 end
                field_runs = [run]
                is_page_field = False
                j = i + 1
                while j < len(runs):
                    field_runs.append(runs[j])
                    instr = runs[j].find(qn('w:instrText'))
                    if instr is not None and 'PAGE' in (instr.text or ''):
                        is_page_field = True
                    fldChar2 = runs[j].find(qn('w:fldChar'))
                    if fldChar2 is not None and fldChar2.get(qn('w:fldCharType')) == 'end':
                        break
                    j += 1
                if is_page_field:
                    for r in field_runs:
                        r.getparent().remove(r)
                    runs = list(para.iter(qn('w:r')))
                    i = 0
                    continue
            i += 1


# =========================================================
# 页眉/页脚清除
# =========================================================

def clear_doc_properties(doc):
    """清除文档属性中的作者和修改者信息。"""
    doc.core_properties.author = ''
    doc.core_properties.last_modified_by = ''


def clear_section_footers(doc, *indices):
    """清除指定节的页脚内容。用法：clear_section_footers(doc, 0, 1, 2)"""
    for idx in indices:
        if idx >= len(doc.sections):
            continue
        sec = doc.sections[idx]
        sec.footer.is_linked_to_previous = False
        for para in sec.footer.paragraphs:
            para.clear()


def _create_empty_header(doc):
    """创建独立的空白 header part 并返回新 rId。"""
    NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    hdr_el = _etree.Element(f'{{{NS}}}hdr')
    _etree.SubElement(hdr_el, f'{{{NS}}}p')

    blob = _etree.tostring(hdr_el, xml_declaration=True, encoding='UTF-8', standalone=True)

    # 找下一个可用的 header part 编号
    max_n = 0
    for rel in doc.part.rels.values():
        try:
            m = re.match(r'/word/header(\d+)\.xml', str(rel.target_part.partname))
            if m:
                max_n = max(max_n, int(m.group(1)))
        except Exception:
            pass

    partname = PackURI(f'/word/header{max_n + 1}.xml')
    content_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'

    header_part = Part(partname, content_type, blob, doc.part.package)
    return doc.part.relate_to(header_part, RT.HEADER)


def clear_section_headers(doc, *indices):
    """清除指定节的全部页眉（default/even/first）。用法：clear_section_headers(doc, 0, 1)"""
    for idx in indices:
        if idx >= len(doc.sections):
            continue
        sec = doc.sections[idx]
        for ref in list(sec._sectPr.findall(qn('w:headerReference'))):
            rId = ref.get(qn('r:id'))
            if not rId or rId not in doc.part.rels:
                continue

            hdr_el = doc.part.rels[rId].target_part._element

            # 检查此 header 是否被其他节共享
            shared = any(
                other_sec is not sec and
                any(other_ref.get(qn('r:id')) == rId
                    for other_ref in other_sec._sectPr.findall(qn('w:headerReference')))
                for other_sec in doc.sections
            )

            if shared:
                # 共享页眉：创建独立空白 header，替换当前节的引用
                ref_type = ref.get(qn('w:type'))
                new_rId = _create_empty_header(doc)
                new_ref = OxmlElement('w:headerReference')
                new_ref.set(qn('r:id'), new_rId)
                if ref_type:
                    new_ref.set(qn('w:type'), ref_type)
                ref.addnext(new_ref)
                sec._sectPr.remove(ref)
            else:
                # 仅本节使用，直接清空
                for child in list(hdr_el):
                    hdr_el.remove(child)


# =========================================================
# CLI
# =========================================================

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="论文 DOCX 节操作与页码控制")
    parser.add_argument("input", help="输入 docx 文件")
    parser.add_argument("output", nargs="?", default=None, help="输出 docx 文件 (默认覆盖输入)")
    parser.add_argument("--split-heading", metavar="N",
                        help="在第 N 个顶级标题前插入分节符")
    parser.add_argument("--split-page", metavar="N",
                        help="在物理第 N 页处插入分节符 (需 Word 渲染过)")
    parser.add_argument("--merge", metavar="I-J",
                        help="合并第 I 到第 J 节 (0-based)")
    parser.add_argument("--pg-start", metavar="S,N",
                        help="设置第 S 节起始页码为 N")
    parser.add_argument("--pg-format", metavar="S,FMT",
                        help="设置第 S 节页码格式 (decimal/upperRoman/lowerRoman/none)")
    parser.add_argument("--clear-footers", metavar="I,J,...",
                        help="清除指定节的页脚")
    parser.add_argument("--clear-props", action="store_true",
                        help="清除文档属性中的作者和修改者")
    parser.add_argument("--clear-headers", metavar="I,J,...",
                        help="清除指定节的页眉")

    args = parser.parse_args()

    # 按 sys.argv 顺序构建 actions，执行顺序与传参顺序一致
    actions = []
    parsed = {k: v for k, v in vars(args).items()
              if k not in ('input', 'output') and v is not None}
    cls = {a.option_strings[0][2:]: a.dest for a in parser._actions
           if a.option_strings}  # --split-heading → split_heading
    for a in sys.argv[1:]:
        name = a.lstrip('-')
        if '=' in name:
            name = name.split('=', 1)[0]
        dest = cls.get(name, '')
        if dest in parsed:
            actions.append((dest, str(parsed.pop(dest))))

    main(args.input, args.output, actions)
