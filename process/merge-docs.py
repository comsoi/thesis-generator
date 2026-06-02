#!/usr/bin/env python3
"""
DOCX 装配器：按输出顺序拼接多个文档。

以指定文件为样式源（master），将其它文件内容 prepend/append 到 master。

核心功能：
- 样式合并：prepend 时，实际使用的同名同类型样式覆盖 master；
           同名不同类型 → 分配新 styleId 并重映射段落引用
- 编号合并：重映射 numId / abstractNumId 避免冲突
- 页眉/页脚保留：将非 master 文件的 header/footer 引用合并到 master body sectPr
- 奇偶页眉：同步 evenAndOddHeaders 设置到 master settings
- 页码起始：保留 pgNumType（如 start=1）到分节符
- rId 冲突处理：重新分配与 master 冲突的 rId，跳过当前文件的 rId 避免自碰撞
- 书签去重、图形 id 重编号、页眉/页脚/图片 part 复制

用法:
    python merge-docs.py <output> <input1> <input2> ...
    样式以最后一个输入文件为准。
"""

import sys
from docx import Document
from docx.oxml.ns import qn
from io import BytesIO
from copy import deepcopy
from lxml import etree
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart

WPML = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


class Assembler:
    """DOCX 装配器：将多个文档按顺序拼接为一个。

    style 参数指定以第几个文件（0-based）作为样式源（master）。
    负数表示从末尾倒数（-1 = 最后一个文件）。
    """

    def __init__(self, style=0):
        self._files = []
        self._style = style
        self._master = None

    def add(self, path):
        """添加文件到装配队列（按添加顺序排列）。"""
        self._files.append(path)
        return self

    def save(self, output):
        """执行装配并保存。

        1. 以 style 指定的文件打开 master
        2. master 之前的文件按逆序 prepend（保证最终顺序正确）
        3. master 之后的文件按正序 append
        4. 保存到 output
        """
        files = self._files
        if len(files) < 2:
            import shutil
            shutil.copy2(files[0], output)
            return

        si = self._style if self._style >= 0 else len(files) + self._style
        if not 0 <= si < len(files):
            raise ValueError(f"style index {self._style} out of range")

        self._master = Document(files[si])

        for f in reversed(files[:si]):
            self._insert(f, 'prepend')
        for f in files[si + 1:]:
            self._insert(f, 'append')

        self._master.save(output)
        self._log(files, si, output)

    @staticmethod
    def _log(files, si, output):
        """输出构建日志。"""
        names = ' | '.join(
            f + (" [style]" if i == si else "")
            for i, f in enumerate(files)
        )
        print(f"assembler: {names} → {output}")

    # =========================================================
    # 单文件插入
    # =========================================================

    def _insert(self, path, position):
        """将单个文件插入 master。

        position: 'prepend' 插入到 master 内容之前
                  'append'  插入到 master 内容之后（body sectPr 之前）

        处理流程：样式 → 编号 → settings → rId 上下文 → sectPr 合并 → body 插入
        """
        source = Document(path)
        style_map = self._merge_styles(source, position)
        self._merge_numbering(source)
        self._merge_settings(source)
        ctx = self._build_ctx(source)
        ctx['style_map'] = style_map
        self._merge_body_sectpr_headers(source, ctx['rid_map'])
        self._insert_body(source, position, ctx)

    def _merge_body_sectpr_headers(self, source, rid_map):
        """将 source body sectPr 中的 header/footer 引用合并到 master body sectPr。

        问题：master（如 pandoc 生成的正文）通常没有页眉定义。
        不加此合并，正文页不会显示页眉。

        按 (tag, type) 去重，避免重复引用。通过 rid_map 重映射 rId。
        """
        s_sectpr = source.element.body.find(qn('w:sectPr'))
        if s_sectpr is None:
            return
        m_sectpr = self._master.element.body.find(qn('w:sectPr'))
        if m_sectpr is None:
            return

        existing = set()
        for el in m_sectpr:
            if el.tag in (qn('w:headerReference'), qn('w:footerReference')):
                existing.add((el.tag, el.get(qn('w:type'))))

        for el in s_sectpr:
            if el.tag in (qn('w:headerReference'), qn('w:footerReference')):
                key = (el.tag, el.get(qn('w:type')))
                if key not in existing:
                    clone = deepcopy(el)
                    if rid_map:
                        Assembler._remap_rids(clone, rid_map)
                    m_sectpr.append(clone)
                    existing.add(key)

    # =========================================================
    # 样式
    # =========================================================

    def _merge_styles(self, source, position):
        """合并样式。

        prepend：source 实际使用的样式覆盖同名同类型 master 样式。
                 同名不同类型 → 分配新 styleId 避免冲突。
        append：master 已有样式不受影响。

        返回 style_map (old_id → new_id) 用于重映射 body 中的 pStyle 引用。
        """
        style_map = {}
        used_ids = self._used_style_ids(source) if position == 'prepend' else set()
        master_el = self._master.styles.element
        existing = {c.get(qn('w:styleId')): c for c in master_el
                    if c.tag == qn('w:style')}
        for child in source.styles.element:
            if child.tag == qn('w:style'):
                sid = child.get(qn('w:styleId'))
                if sid in existing:
                    if sid in used_ids:
                        if self._same_style_type(child, existing[sid]):
                            master_el.replace(existing[sid], deepcopy(child))
                            existing[sid] = child
                        else:
                            new_id = self._next_style_id(existing)
                            clone = deepcopy(child)
                            clone.set(qn('w:styleId'), new_id)
                            master_el.append(clone)
                            existing[new_id] = clone
                            style_map[sid] = new_id
                else:
                    master_el.append(deepcopy(child))
                    existing[sid] = child
            elif child.tag == qn('w:docDefaults'):
                if master_el.find(child.tag) is None:
                    master_el.append(deepcopy(child))
        return style_map

    @staticmethod
    def _next_style_id(existing):
        """返回不在 existing 中的新 styleId（从现有最大数字 ID+1 开始）。"""
        max_id = max((int(k) for k in existing if k.isdigit()), default=0)
        i = max_id + 1
        while str(i) in existing:
            i += 1
        return str(i)

    @staticmethod
    def _same_style_type(a, b):
        """两个 style 元素是否同类型。防止字符样式覆盖段落样式（如 Title Char 覆盖 First Paragraph）。"""
        return a.get(qn('w:type')) == b.get(qn('w:type'))

    @staticmethod
    def _used_style_ids(doc):
        """返回文档中实际使用的 styleId（body + 页眉/页脚内的段落/字符/表格样式）。"""
        used = set()

        def scan_p(p_elem):
            pPr = p_elem.find(qn('w:pPr'))
            if pPr is not None:
                ps = pPr.find(qn('w:pStyle'))
                if ps is not None:
                    used.add(ps.get(qn('w:val')) or '')
            for r in p_elem.iter(qn('w:r')):
                rPr = r.find(qn('w:rPr'))
                if rPr is not None:
                    rs = rPr.find(qn('w:rStyle'))
                    if rs is not None:
                        used.add(rs.get(qn('w:val')) or '')

        def scan_tbl(tbl_elem):
            tblPr = tbl_elem.find(qn('w:tblPr'))
            if tblPr is not None:
                ts = tblPr.find(qn('w:tblStyle'))
                if ts is not None:
                    used.add(ts.get(qn('w:val')) or '')
            for p in tbl_elem.iter(qn('w:p')):
                scan_p(p)

        # body: iter 递归扫描，覆盖表格内部的段落
        for p in doc.element.body.iter(qn('w:p')):
            scan_p(p)
        for tbl in doc.element.body.iter(qn('w:tbl')):
            scan_tbl(tbl)

        # 页眉/页脚单独递归扫描
        for rel in doc.part.rels.values():
            if 'header' in rel.reltype or 'footer' in rel.reltype:
                for p in rel.target_part._element.iter(qn('w:p')):
                    scan_p(p)
                for tbl in rel.target_part._element.iter(qn('w:tbl')):
                    scan_tbl(tbl)

        return used

    # =========================================================
    # 编号
    # =========================================================

    def _merge_numbering(self, source):
        """合并编号定义。

        重映射 abstractNumId 和 numId 避免与 master 冲突。
        映射结果存入 self._num_map 供后续 body 元素中的 numId 引用重映射。
        """
        src_part = self._numbering_part(source)
        if src_part is None:
            return
        dst_part = self._ensure_numbering_part()

        abs_ids = {int(an.get(qn('w:abstractNumId')))
                   for an in dst_part.findall(qn('w:abstractNum'))}
        num_ids = {n.get(qn('w:numId'))
                   for n in dst_part.findall(qn('w:num'))}

        # 复制 abstractNum，重映射 ID
        abs_map = {}
        for an in src_part.findall(qn('w:abstractNum')):
            aid = int(an.get(qn('w:abstractNumId')))
            if aid not in abs_ids:
                new_aid = self._next_int_id(abs_ids)
                clone = deepcopy(an)
                clone.set(qn('w:abstractNumId'), str(new_aid))
                dst_part.append(clone)
                abs_ids.add(new_aid)
                abs_map[aid] = new_aid

        # 复制 num，重映射 numId 和 abstractNumId 引用
        self._num_map = {}
        for n in src_part.findall(qn('w:num')):
            nid = n.get(qn('w:numId'))
            anc = n.find(qn('w:abstractNumId'))
            old_aid = int(anc.get(qn('w:val'))) if anc is not None else None

            new_nid = nid if nid not in num_ids else str(self._next_int_id(num_ids))
            num_ids.add(new_nid)

            clone = deepcopy(n)
            clone.set(qn('w:numId'), new_nid)
            if anc is not None and old_aid is not None and old_aid in abs_map:
                clone.find(qn('w:abstractNumId')).set(qn('w:val'), str(abs_map[old_aid]))

            dst_part.append(clone)
            self._num_map[int(nid)] = int(new_nid)

    @staticmethod
    def _numbering_part(doc):
        """获取 numbering.xml 的 lxml element，不存在返回 None。"""
        for rel in doc.part.rels.values():
            if 'numbering' in rel.reltype:
                return rel.target_part._element
        return None

    def _ensure_numbering_part(self):
        """获取或创建 master 的 numbering part。"""
        part = self._numbering_part(self._master)
        if part is not None:
            return part
        elem = etree.fromstring(
            '<?xml version="1.0"?>'
            '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
            .encode('utf-8'))
        p = XmlPart(
            partname=PackURI('/word/numbering.xml'),
            content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
            element=elem,
            package=self._master.part.package,
        )
        rId = self._next_rid()
        self._master.part.rels.add_relationship(
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
            p, rId)
        return elem

    @staticmethod
    def _next_int_id(existing):
        """返回不在 existing 集合中的最小正整数 id（字符串形式）。"""
        i = 1
        while str(i) in existing:
            i += 1
        return i

    # =========================================================
    # 文档设置（evenAndOddHeaders 等）
    # =========================================================

    @staticmethod
    def _settings_part(doc):
        """获取 settings.xml 的 lxml element，不存在返回 None。"""
        for rel in doc.part.rels.values():
            if rel.reltype == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings':
                return rel.target_part._element
        return None

    def _ensure_settings_part(self):
        """获取或创建 master 的 settings part。"""
        part = self._settings_part(self._master)
        if part is not None:
            return part
        elem = etree.fromstring(
            '<?xml version="1.0"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
            .encode('utf-8'))
        p = XmlPart(
            partname=PackURI('/word/settings.xml'),
            content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
            element=elem,
            package=self._master.part.package,
        )
        rId = self._next_rid()
        self._master.part.rels.add_relationship(
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',
            p, rId)
        return elem

    def _merge_settings(self, source):
        """同步 source 的 evenAndOddHeaders 设置到 master。

        问题：封面模板启用了奇偶页不同页眉，但 pandoc 生成的正文没有此设置。
        master 缺少此设置时，Word 不区分奇偶页，所有页只显示一种页眉。
        """
        src_settings = self._settings_part(source)
        if src_settings is None:
            return
        if src_settings.find(qn('w:evenAndOddHeaders')) is None:
            return
        dst_settings = self._ensure_settings_part()
        if dst_settings.find(qn('w:evenAndOddHeaders')) is None:
            etree.SubElement(dst_settings, qn('w:evenAndOddHeaders'))

    # =========================================================
    # 上下文（rId 映射、书签、docPr ID）
    # =========================================================

    def _build_ctx(self, source):
        """构建 source → master 的映射上下文。

        - rid_map: source 的 rId → master 的 rId（处理冲突重分配）
          * rId 不在 master 中：直接复用原 rId，复制 part
          * rId 在 master 中且 reltype 相同：直接复用（如图片共享）
          * rId 在 master 中但 reltype 不同：分配新 rId，复制 part
          * 分配新 rId 时跳过当前 source 中所有存在的 rId，避免自碰撞
            （如 even header=rId8 → new=rId9，但 rId9 也是 source 的 default header，
              rId 碰撞导致两个页眉引用了同一 part）
        - bookmarks: master 已有书签集合（用于重命名冲突书签）
        - next_id: 下一个可用的图形 docPr id
        - num_map: 编号 id 映射（来自 _merge_numbering）
        """
        master_rels = self._master.part.rels
        source_rels = source.part.rels

        # 收集 source body 中所有 rId 引用（包括 sectPr 中的页眉/页脚引用）
        referenced = set()
        for elem in source.element.body.iter():
            for ak, av in elem.attrib.items():
                if av.startswith('rId') and (ak.endswith('}id') or ak.endswith('}embed')):
                    referenced.add(av)

        # 构建 rId 映射，复制引用的 part 到 master
        rid_map = {}
        for rId in referenced:
            if rId not in source_rels:
                continue
            rel = source_rels[rId]
            if rId not in master_rels:
                self._copy_part(rel, rId)
                rid_map[rId] = rId
            elif rel.reltype == master_rels[rId].reltype:
                rid_map[rId] = rId
            else:
                new_rId = self._next_rid(referenced)
                self._copy_part(rel, new_rId)
                rid_map[rId] = new_rId

        # 收集 master 中已有书签名称
        bookmarks = set()
        for elem in self._master.element.body.iter():
            if elem.tag == qn('w:bookmarkStart'):
                bookmarks.add(elem.get(qn('w:name')))

        # 找到 master 中最大的图形 docPr id
        max_id = 0
        for elem in self._master.element.body.iter():
            dp = elem.get('id')
            if dp and dp.isdigit():
                max_id = max(max_id, int(dp))

        return {
            'rid_map': rid_map,
            'bookmarks': bookmarks,
            'next_id': max_id + 1,
            'num_map': getattr(self, '_num_map', {}),
        }

    # =========================================================
    # Body 元素插入
    # =========================================================

    def _insert_body(self, source, position, ctx):
        """将 source 的 body 子元素（段落、表格）插入到 master。

        prepend：插入到 master 第一个元素之前（封面 → 正文）
        append：插入到 master body sectPr 之前（正文末尾之后）

        最后一段附带 source body sectPr 衍生的分节符，保留页面尺寸、
        页眉/页脚引用和页码起始设置。
        已有段落级 sectPr 的段落则保持原样不修改。
        """
        m_body = self._master.element.body
        s_body = source.element.body
        s_sectpr = s_body.find(qn('w:sectPr'))
        m_sectpr = m_body.find(qn('w:sectPr'))

        if position == 'prepend':
            anchor = m_body[0] if len(m_body) > 0 else None

            def put(e):
                if anchor is not None:
                    anchor.addprevious(e)
                else:
                    m_body.append(e)
        else:
            def put(e):
                idx = -1 if m_sectpr is not None else len(m_body)
                m_body.insert(idx, e)

        # 找到 source 最后一个有效元素（段落或表格）
        last = None
        for child in s_body:
            if child.tag in (qn('w:p'), qn('w:tbl')):
                last = child

        for child in s_body:
            # body-level sectPr 不直接插入（其属性通过 _attach_sectpr 转移）
            if child is s_sectpr:
                continue
            elem = etree.fromstring(etree.tostring(child, encoding='unicode'))
            self._apply_ctx(elem, ctx)

            # 在最后一页附加从 body sectPr 派生的分节符
            if child is last and s_sectpr is not None:
                self._attach_sectpr(elem, s_sectpr, position, ctx['rid_map'])

            put(elem)

    # =========================================================
    # 元素后处理（rId / numId / 书签 / 图形 id 重映射）
    # =========================================================

    def _apply_ctx(self, elem, ctx):
        """对插入的每个元素应用所有上下文映射。"""
        self._remap_rids(elem, ctx['rid_map'])
        self._remap_num_ids(elem, ctx['num_map'])
        self._remap_style_ids(elem, ctx.get('style_map', {}))
        self._renumber_bookmarks(elem, ctx['bookmarks'])
        self._renumber_shape_ids(elem, ctx)

    @staticmethod
    def _remap_rids(elem, rid_map):
        """重映射元素中所有 rId 属性值。"""
        for el in elem.iter():
            for ak in list(el.attrib.keys()):
                if el.attrib[ak] in rid_map:
                    el.attrib[ak] = rid_map[el.attrib[ak]]

    @staticmethod
    def _remap_style_ids(elem, style_map):
        """重映射段落 pStyle 引用（冲突样式重新分配 ID 后更新）。"""
        if not style_map:
            return
        for p in elem.iter(qn('w:p')):
            pPr = p.find(qn('w:pPr'))
            if pPr is not None:
                ps = pPr.find(qn('w:pStyle'))
                if ps is not None:
                    val = ps.get(qn('w:val')) or ''
                    if val in style_map:
                        ps.set(qn('w:val'), style_map[val])

    @staticmethod
    def _remap_num_ids(elem, num_map):
        """重映射编号 id（w:numId 的 w:val 属性）。"""
        if not num_map:
            return
        for el in elem.iter():
            if el.tag == qn('w:numId'):
                v = int(el.get(qn('w:val')))
                if v in num_map:
                    el.set(qn('w:val'), str(num_map[v]))

    @staticmethod
    def _renumber_bookmarks(elem, existing):
        """重命名与 master 冲突的书签（追加 _1, _2 ...）。"""
        for el in elem.iter():
            if el.tag == qn('w:bookmarkStart'):
                name = el.get(qn('w:name'))
                if name in existing:
                    i = 1
                    while f"{name}_{i}" in existing:
                        i += 1
                    name = f"{name}_{i}"
                    el.set(qn('w:name'), name)
                existing.add(name)

    @staticmethod
    def _renumber_shape_ids(elem, ctx):
        """给图形 docPr id 递增编号，避免与 master 冲突。"""
        nid = ctx['next_id']
        for el in elem.iter():
            dp = el.get('id')
            if dp is not None and dp.isdigit():
                el.set('id', str(nid))
                nid += 1
        ctx['next_id'] = nid

    # =========================================================
    # Part 复制（图片、页眉、页脚）
    # =========================================================

    def _copy_part(self, rel, rId):
        """将 source 的 part 复制到 master 并用指定 rId 建立关系。"""
        reltype = rel.reltype
        if "image" in reltype:
            blob = rel.target_part.blob
            part = self._master.part.package.get_or_add_image_part(BytesIO(blob))
        elif "header" in reltype or "footer" in reltype:
            part = self._clone_hdr_ftr(rel.target_part, reltype)
        else:
            return
        self._master.part.rels.add_relationship(reltype, part, rId)

    def _next_rid(self, skip=None):
        """返回下一个可用的 rId。

        skip 参数用于排除当前 source 中已被引用的 rId，
        防止新分配的 rId 与 source 中尚未处理的 rId 冲突。
        """
        i = 1
        while f'rId{i}' in self._master.part.rels or (skip and f'rId{i}' in skip):
            i += 1
        return f'rId{i}'

    def _clone_hdr_ftr(self, src_part, reltype):
        """克隆页眉/页脚 part 到 master 包中（避免共享修改）。"""
        for key, prefix in [('header', 'header'), ('footer', 'footer')]:
            if key in reltype:
                elem = etree.fromstring(src_part._element.xml.encode('utf-8'))
                i = 1
                existing = {p.partname for p in self._master.part.package.parts}
                while PackURI(f'/word/{prefix}{i}.xml') in existing:
                    i += 1
                ct = f'application/vnd.openxmlformats-officedocument.wordprocessingml.{prefix}+xml'
                return XmlPart(
                    partname=PackURI(f'/word/{prefix}{i}.xml'),
                    content_type=ct,
                    element=elem,
                    package=self._master.part.package,
                )
        return None

    # =========================================================
    # sectPr（分节符属性）
    # =========================================================

    @staticmethod
    def _attach_sectpr(paragraph_elem, src_sectpr, position, rid_map=None):
        """在段落的 pPr 中附加分节符 sectPr。

        如果段落已有段落级 sectPr（如封面模板中"封面/原创声明"之间的分节符），
        保留原样不修改（这些是模板设计者有意设置的分节边界）。

        如果没有，从 source 的 body-level sectPr 派生新的分节符，
        拷贝页面尺寸、页边距、分栏、网格、页眉/页脚引用、页码起始类型。
        拷贝的 header/footer 引用通过 rid_map 重映射到 master 中的 rId。
        """
        pPr = paragraph_elem.find(qn('w:pPr'))
        if pPr is None:
            pPr = etree.SubElement(paragraph_elem, qn('w:pPr'))

        # 段落已有段落级分节符：保留模板原始设计
        if pPr.find(qn('w:sectPr')) is not None:
            return

        # 从 body-level sectPr 创建新的分节符
        dest = etree.SubElement(pPr, qn('w:sectPr'))
        for child in src_sectpr:
            if child.tag in (
                qn('w:headerReference'), qn('w:footerReference'),
                qn('w:pgSz'), qn('w:pgMar'), qn('w:cols'), qn('w:docGrid'),
                qn('w:pgNumType'),  # 页码起始（如 start=1 使第二节首页为奇数页）
            ):
                copy = deepcopy(child)
                if rid_map:
                    Assembler._remap_rids(copy, rid_map)
                dest.append(copy)


# =========================================================
# CLI
# =========================================================

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(f"用法: {sys.argv[0]} <output> <input1> <input2> [...]", file=sys.stderr)
        print("  样式以最后一个输入文件为准。", file=sys.stderr)
        sys.exit(1)

    output = sys.argv[1]
    inputs = sys.argv[2:]

    a = Assembler(style=-1)
    for f in inputs:
        a.add(f)
    a.save(output)
