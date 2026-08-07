#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
build_journal_docx.py —— 马理论科期刊投稿式排版（内置双栏期刊模板）

把结构化 markdown 草稿 + meta.json 元信息，生成符合内置双栏期刊模板的 Word 文档：
- 顺序编码制真尾注（endnoteReference 链接 endnotes.xml，可点击跳转、自动编号）
- 文中上标 [N] 方括号（vertAlign=superscript，三个 run：[ + endnoteReference + ]）
- 一文一引：每篇文献只在文中引用一次；检测到重复标注时保留首次、删掉后面的
- 自动重编号：按文中首次出现顺序重排编号，参考文献表同步重排
- "参考文献"标题做成一条不被引用的尾注（永远贴着尾注列表，不会分离跑页）
- 作者简介做成第一页真脚注（自动编号），文末另附作者简介+联系方式段
- 自动重编号/重复引用处理结果在 stdout 报告，供 agent 转述给用户

用法：
    python build_journal_docx.py --md 论文正文.md --meta meta.json --out 投稿版.docx
    可选：--no-english  强制不出英文部分（默认：meta 里有 english 块就出）

输入约定（markdown）：
    # 标题
    **摘要：** 摘要内容……
    **关键词：** 词1；词2；……
    ## 引言        ← 可选；按内置模板"引言"不排标题，只排其下段落
    正文段落，文中引用标注写成 [1] [2] ……
    ## 一、XXX     ← 一级标题（黑体）
    ### （一）XXX  ← 二级标题（楷体）
    #### 1. XXX    ← 三级标题（宋体）
    ## 结语
    **参考文献：**
    [1] 作者.题名[J].刊名,年(期):页码.
    ……

meta.json 字段：
    author       作者姓名
    affiliation  单位，如"（XX大学XX学院，北京 100000）"（括号自带）
    author_bio   作者简介（不含"作者简介："前缀），如"张三（2000—），男，……研究。"
    contact      文末联系方式，如"联系电话：13800000000，电子邮箱：xx@xx.edu.cn。"
    clc          中图分类号，默认 G641
    doc_code     文献标识码，默认 A
    fund         基金项目（可选，进第一页脚注）
    received     收稿日期（可选，进第一页脚注），如"2026-01-01"
    english      可选对象 {title, author, affiliation, abstract, keywords}

依赖：python-docx、lxml
"""

import argparse
import copy
import json
import re
import sys
import zipfile

try:
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.opc.part import Part
from docx.opc.packuri import PackURI
from lxml import etree

XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

# ==================== 内置双栏期刊模板（字号/字体规格） ====================
# 规格来自实际投稿验证过的期刊版式；新增模板时在此扩展。
TPL = {
    'page': dict(width=21, height=29.7, top=3, bottom=3, left=2, right=2, grid=312),
    'title':        dict(cn='宋体', en='Times New Roman', sz=15, bold=True),   # 小三
    'author':       dict(cn='楷体', en='楷体', sz=12, bold=False),              # 小四
    'affiliation':  dict(cn='宋体', en='Times New Roman', sz=10.5, bold=False), # 五号
    'abs_label':    dict(cn='黑体', en='Times New Roman', sz=9, bold=True),     # 小五
    'abs_content':  dict(cn='楷体', en='楷体', sz=9, bold=False),
    'clc_label':    dict(cn='黑体', en='Times New Roman', sz=9, bold=True),
    'clc_content':  dict(cn='宋体', en='Times New Roman', sz=9, bold=False),
    'h1':           dict(cn='黑体', en='Times New Roman', sz=10.5, bold=True),  # 一、
    'h2':           dict(cn='楷体', en='楷体', sz=10.5, bold=True),             # （一）
    'h3':           dict(cn='宋体', en='Times New Roman', sz=10.5, bold=True),  # 1.
    'body':         dict(cn='宋体', en='Times New Roman', sz=10.5, bold=False), # 五号
    'cite':         dict(sz=9),                                                  # 上标[N] 小五
    'ref_entry':    dict(cn='宋体', en='Times New Roman', sz=9),                 # 尾注条目 小五
    'ref_heading':  dict(cn='宋体', en='Times New Roman', sz=10.5, bold=True),
    'en_title':     dict(en='Times New Roman', sz=15, bold=True),
    'en_author':    dict(en='Times New Roman', sz=12),
    'en_affil':     dict(en='Times New Roman', sz=10.5),
    'en_abs':       dict(en='Times New Roman', sz=10.5),
    'fn_author':    dict(sz=7.5),                                                # 脚注 六号
    'bio_end':      dict(cn='宋体', en='Times New Roman', sz=9),                 # 文末作者简介
}


# ==================== 工具函数 ====================

def set_run_font(run, cn, en, sz, bold=False):
    run.font.size = Pt(sz)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0)
    r = run._element
    rPr = r.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        r.insert(0, rPr)
    rf = rPr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts')
        rPr.insert(0, rf)
    rf.set(qn('w:ascii'), en)
    rf.set(qn('w:hAnsi'), en)
    rf.set(qn('w:eastAsia'), cn)
    rf.set(qn('w:cs'), en)
    rf.set(qn('w:hint'), 'eastAsia')


def set_spacing(para, sb=0, sa=0):
    para.paragraph_format.space_before = Pt(sb)
    para.paragraph_format.space_after = Pt(sa)


def set_columns(section, n):
    sp = section._sectPr
    for c in sp.findall(qn('w:cols')):
        sp.remove(c)
    c = OxmlElement('w:cols')
    c.set(qn('w:num'), str(n))
    if n > 1:
        c.set(qn('w:space'), '425')
        c.set(qn('w:equalWidth'), '1')
    sp.append(c)


def set_docgrid(section, lp=312):
    sp = section._sectPr
    for d in sp.findall(qn('w:docGrid')):
        sp.remove(d)
    d = OxmlElement('w:docGrid')
    d.set(qn('w:type'), 'lines')
    d.set(qn('w:linePitch'), str(lp))
    sp.append(d)


def set_endnote_pr(section):
    """节级尾注属性：尾注排在本节末尾（参考文献位于正文节末、英文部分之前）+ 十进制编号。"""
    sp = section._sectPr
    for ep in sp.findall(qn('w:endnotePr')):
        sp.remove(ep)
    ep = OxmlElement('w:endnotePr')
    p = OxmlElement('w:pos'); p.set(qn('w:val'), 'sectEnd'); ep.append(p)
    n = OxmlElement('w:numFmt'); n.set(qn('w:val'), 'decimal'); ep.append(n)
    sp.append(ep)


def setup_section(section, cols=1):
    pg = TPL['page']
    section.page_width = Cm(pg['width']); section.page_height = Cm(pg['height'])
    section.top_margin = Cm(pg['top']); section.bottom_margin = Cm(pg['bottom'])
    section.left_margin = Cm(pg['left']); section.right_margin = Cm(pg['right'])
    set_docgrid(section, pg['grid'])
    set_columns(section, cols)
    set_endnote_pr(section)


def add_center(doc, text, cn, en, sz, bold=False, sb=0, sa=0):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    set_run_font(r, cn, en, sz, bold)
    set_spacing(p, sb, sa)
    return p


def add_body(doc, text, style_key='body', indent=True, sb=0, sa=0):
    st = TPL[style_key]
    p = doc.add_paragraph()
    r = p.add_run(text)
    set_run_font(r, st['cn'], st['en'], st['sz'], st.get('bold', False))
    if indent:
        p.paragraph_format.first_line_indent = Pt(st['sz'] * 2)
    set_spacing(p, sb, sa)
    return p


def add_mixed(doc, parts, sb=0, sa=0):
    p = doc.add_paragraph()
    for part in parts:
        text, style_key = part[0], part[1]
        st = TPL[style_key]
        cn = st.get('cn', '宋体')
        r = p.add_run(text)
        set_run_font(r, cn, st['en'], st['sz'], st.get('bold', False))
    set_spacing(p, sb, sa)
    return p


# ==================== 脚注（作者简介，第一页） ====================

FOOTNOTES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:footnote w:type="separator" w:id="-1">'
    '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    '<w:r><w:separator/></w:r></w:p></w:footnote>'
    '<w:footnote w:type="continuationSeparator" w:id="0">'
    '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    '<w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
    '</w:footnotes>'
)


def ensure_footnotes_part(doc):
    for rel in doc.part.rels.values():
        if 'footnotes' in str(rel.reltype):
            return rel.target_part
    ct = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml'
    pt = Part(PackURI('/word/footnotes.xml'), ct, FOOTNOTES_XML.encode('utf-8'), doc.part.package)
    doc.part.relate_to(pt, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes')
    return pt


def add_author_fn_content(doc, fn_id, lines, sz=7.5):
    """第一页脚注内容：作者简介/基金项目/收稿日期，每项一段。"""
    fn_part = ensure_footnotes_part(doc)
    root = etree.fromstring(fn_part.blob)
    fn = etree.SubElement(root, qn('w:footnote'))
    fn.set(qn('w:id'), str(fn_id))
    for line in lines:
        p = etree.SubElement(fn, qn('w:p'))
        pp = etree.SubElement(p, qn('w:pPr'))
        sp = etree.SubElement(pp, qn('w:spacing'))
        sp.set(qn('w:after'), '0'); sp.set(qn('w:line'), '240'); sp.set(qn('w:lineRule'), 'auto')
        r = etree.SubElement(p, qn('w:r'))
        rPr = etree.SubElement(r, qn('w:rPr'))
        rf = etree.SubElement(rPr, qn('w:rFonts'))
        rf.set(qn('w:ascii'), 'Times New Roman'); rf.set(qn('w:hAnsi'), 'Times New Roman')
        rf.set(qn('w:eastAsia'), '宋体')
        sz_e = etree.SubElement(rPr, qn('w:sz')); sz_e.set(qn('w:val'), str(int(sz * 2)))
        szc = etree.SubElement(rPr, qn('w:szCs')); szc.set(qn('w:val'), str(int(sz * 2)))
        t = etree.SubElement(r, qn('w:t')); t.text = line; t.set(XML_SPACE, 'preserve')
    fn_part._blob = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


def add_fn_ref(paragraph):
    """在作者名后加脚注引用（自动编号，显示为 1 的上标）。"""
    r = paragraph.add_run()
    rp = OxmlElement('w:rPr')
    va = OxmlElement('w:vertAlign'); va.set(qn('w:val'), 'superscript'); rp.append(va)
    rf = OxmlElement('w:rFonts')
    rf.set(qn('w:ascii'), '楷体'); rf.set(qn('w:hAnsi'), '楷体'); rf.set(qn('w:eastAsia'), '楷体')
    rp.append(rf)
    r._element.append(rp)
    fr = OxmlElement('w:footnoteReference')
    fr.set(qn('w:id'), '1')
    r._element.append(fr)


# ==================== 尾注（顺序编码制） ====================

ENDNOTES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:endnote w:type="separator" w:id="-1">'
    '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    '<w:r><w:separator/></w:r></w:p></w:endnote>'
    '<w:endnote w:type="continuationSeparator" w:id="0">'
    '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    '<w:r><w:continuationSeparator/></w:r></w:p></w:endnote>'
    '</w:endnotes>'
)


def ensure_endnotes_part(doc):
    for rel in doc.part.rels.values():
        if 'endnotes' in str(rel.reltype):
            return rel.target_part
    ct = 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml'
    pt = Part(PackURI('/word/endnotes.xml'), ct, ENDNOTES_XML.encode('utf-8'), doc.part.package)
    doc.part.relate_to(pt, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes')
    return pt


def _set_part_blob(part, root):
    part._blob = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


def add_ref_heading_endnote(doc, en_id):
    """"参考文献"标题做成一条不被文中引用的尾注：
    渲染时排在尾注区最前、不编号、永远紧贴尾注列表，不会与列表分离跑页。"""
    en_part = ensure_endnotes_part(doc)
    root = etree.fromstring(en_part.blob)
    en = etree.SubElement(root, qn('w:endnote'))
    en.set(qn('w:id'), str(en_id))
    p = etree.SubElement(en, qn('w:p'))
    pp = etree.SubElement(p, qn('w:pPr'))
    sp = etree.SubElement(pp, qn('w:spacing'))
    sp.set(qn('w:before'), '120'); sp.set(qn('w:after'), '60')
    r = etree.SubElement(p, qn('w:r'))
    rPr = etree.SubElement(r, qn('w:rPr'))
    rf = etree.SubElement(rPr, qn('w:rFonts'))
    rf.set(qn('w:ascii'), 'Times New Roman'); rf.set(qn('w:hAnsi'), 'Times New Roman')
    rf.set(qn('w:eastAsia'), TPL['ref_heading']['cn'])
    etree.SubElement(rPr, qn('w:b'))
    sz = etree.SubElement(rPr, qn('w:sz'))
    sz.set(qn('w:val'), str(int(TPL['ref_heading']['sz'] * 2)))
    t = etree.SubElement(r, qn('w:t')); t.text = '参考文献'
    _set_part_blob(en_part, root)


def add_endnote_entry(doc, en_id, ref_text):
    """一条尾注文献条目：[ endnoteRef ] 文献内容，小五号，单倍行距。"""
    en_part = ensure_endnotes_part(doc)
    root = etree.fromstring(en_part.blob)
    st = TPL['ref_entry']
    sz_half = str(int(st['sz'] * 2))

    en = etree.SubElement(root, qn('w:endnote'))
    en.set(qn('w:id'), str(en_id))
    p = etree.SubElement(en, qn('w:p'))
    pp = etree.SubElement(p, qn('w:pPr'))
    sp = etree.SubElement(pp, qn('w:spacing'))
    sp.set(qn('w:after'), '0'); sp.set(qn('w:line'), '240'); sp.set(qn('w:lineRule'), 'auto')

    # [
    rl = etree.SubElement(p, qn('w:r'))
    rpl = etree.SubElement(rl, qn('w:rPr'))
    szl = etree.SubElement(rpl, qn('w:sz')); szl.set(qn('w:val'), sz_half)
    szlc = etree.SubElement(rpl, qn('w:szCs')); szlc.set(qn('w:val'), sz_half)
    tl = etree.SubElement(rl, qn('w:t')); tl.text = '['; tl.set(XML_SPACE, 'preserve')

    # endnoteRef（自动编号）
    re_run = etree.SubElement(p, qn('w:r'))
    rpe = etree.SubElement(re_run, qn('w:rPr'))
    rs = etree.SubElement(rpe, qn('w:rStyle')); rs.set(qn('w:val'), 'EndnoteReference')
    sze = etree.SubElement(rpe, qn('w:sz')); sze.set(qn('w:val'), sz_half)
    szec = etree.SubElement(rpe, qn('w:szCs')); szec.set(qn('w:val'), sz_half)
    etree.SubElement(re_run, qn('w:endnoteRef'))

    # ] + 空格 + 文献内容
    rr = etree.SubElement(p, qn('w:r'))
    rpr = etree.SubElement(rr, qn('w:rPr'))
    rf = etree.SubElement(rpr, qn('w:rFonts'))
    rf.set(qn('w:ascii'), st['en']); rf.set(qn('w:hAnsi'), st['en'])
    rf.set(qn('w:eastAsia'), st['cn'])
    szr = etree.SubElement(rpr, qn('w:sz')); szr.set(qn('w:val'), sz_half)
    szrc = etree.SubElement(rpr, qn('w:szCs')); szrc.set(qn('w:val'), sz_half)
    tr = etree.SubElement(rr, qn('w:t')); tr.text = '] ' + ref_text; tr.set(XML_SPACE, 'preserve')

    _set_part_blob(en_part, root)


def make_superscript_bracket(text):
    """构造上标括号 run 元素（[ 或 ]）。"""
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rf = OxmlElement('w:rFonts')
    rf.set(qn('w:ascii'), 'Times New Roman'); rf.set(qn('w:hAnsi'), 'Times New Roman')
    rf.set(qn('w:eastAsia'), '宋体')
    rPr.append(rf)
    va = OxmlElement('w:vertAlign'); va.set(qn('w:val'), 'superscript'); rPr.append(va)
    sz_half = str(int(TPL['cite']['sz'] * 2))
    sz = OxmlElement('w:sz'); sz.set(qn('w:val'), sz_half); rPr.append(sz)
    szc = OxmlElement('w:szCs'); szc.set(qn('w:val'), sz_half); rPr.append(szc)
    r.append(rPr)
    t = OxmlElement('w:t'); t.text = text; t.set(XML_SPACE, 'preserve')
    r.append(t)
    return r


def make_endnote_ref_run(en_id):
    """构造上标 endnoteReference run。"""
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rs = OxmlElement('w:rStyle'); rs.set(qn('w:val'), 'EndnoteReference'); rPr.append(rs)
    va = OxmlElement('w:vertAlign'); va.set(qn('w:val'), 'superscript'); rPr.append(va)
    sz_half = str(int(TPL['cite']['sz'] * 2))
    sz = OxmlElement('w:sz'); sz.set(qn('w:val'), sz_half); rPr.append(sz)
    szc = OxmlElement('w:szCs'); szc.set(qn('w:val'), sz_half); rPr.append(szc)
    r.append(rPr)
    er = OxmlElement('w:endnoteReference'); er.set(qn('w:id'), str(en_id))
    r.append(er)
    return r


def postprocess_endnotes(doc, refs_by_new_num, n_refs, report):
    """扫描全文 [N] → 按首次出现顺序重编号（一文一引）→ 替换为上标尾注引用 → 建尾注条目。

    refs_by_new_num 已经是按出现顺序重排好的参考文献（1..n_refs）。
    重复引用同一编号：保留首次、删掉后面的（一文一引原则）。
    """
    # Phase 1: 收集所有 [N]
    cite_records = []
    for para in doc.paragraphs:
        for run in para.runs:
            text = run.text or ''
            for m in re.finditer(r'\[(\d+)\]', text):
                cite_records.append((para, run, m.start(), m.end(), int(m.group(1))))

    # Phase 2: 旧编号 → 新编号（按首次出现顺序），重复引用只保留首次
    old_to_new = {}
    removed_dupes = []  # (para, run, ms, me, old_n)
    seen_count = {}
    kept_cites = []
    for para, run, ms, me, old_n in cite_records:
        seen_count[old_n] = seen_count.get(old_n, 0) + 1
        if old_n not in old_to_new:
            old_to_new[old_n] = len(old_to_new) + 1
            kept_cites.append((para, run, ms, me, old_n, old_to_new[old_n]))
        else:
            removed_dupes.append((para, run, ms, me, old_n))

    for old_n, cnt in sorted(seen_count.items()):
        if cnt > 1:
            report.append(f'⚠ 一文一引：原[{old_n}]在文中出现 {cnt} 次，保留第 1 次，删掉后 {cnt - 1} 处标注')
    if any(old_to_new.get(k, k) != k for k in old_to_new):
        mapping = ', '.join(f'[{k}]→[{v}]' for k, v in sorted(old_to_new.items()) if k != v)
        report.append(f'ℹ 编号已按文中首次出现顺序重排（顺序编码制）：{mapping}')

    # 编号连续性检查：文中标注与参考文献表必须对得上
    cited_new = set(old_to_new.values())
    for i in range(1, n_refs + 1):
        if i not in cited_new:
            report.append(f'⚠ 参考文献第 {i} 条未在正文中被引用，已删除该条（顺序编码制只列被引用的）')
    missing = sorted(n for n in cited_new if n > n_refs)
    if missing:
        raise SystemExit(f'错误：文中标注 {missing} 没有对应的参考文献条目，请补齐后重跑')

    # 重排参考文献表（按新编号）
    refs_reordered = {}
    for old_n, new_n in old_to_new.items():
        if old_n in refs_by_new_num:
            refs_reordered[new_n] = refs_by_new_num[old_n]
    # 被删条目：原编号未被引用过的 refs 不进表（已由上面警告）
    final_refs = {i: refs_reordered[i] for i in sorted(refs_reordered)}

    # Phase 3: 建尾注条目（id=1 留给"参考文献"标题，内容条目从 id=2 开始）
    ensure_endnotes_part(doc)
    add_ref_heading_endnote(doc, 1)
    for new_n in range(1, max(final_refs) + 1 if final_refs else 1):
        add_endnote_entry(doc, new_n + 1, final_refs.get(new_n, ''))

    # Phase 4: 按 run 分组替换（同一 run 多个 [N] 一次性替换；重复标注直接删除文本）
    from collections import OrderedDict
    run_groups = OrderedDict()
    for para, run, ms, me, old_n, new_n in kept_cites:
        rk = id(run)
        if rk not in run_groups:
            run_groups[rk] = (para, run, [])
        run_groups[rk][2].append((ms, me, new_n, True))
    for para, run, ms, me, old_n in removed_dupes:
        rk = id(run)
        if rk not in run_groups:
            run_groups[rk] = (para, run, [])
        run_groups[rk][2].append((ms, me, None, False))

    replaced = 0
    for rk, (para, run, cites) in run_groups.items():
        parent = run._element.getparent()
        if parent is None:
            continue
        idx = list(parent).index(run._element)
        text = run.text or ''
        orig_rPr = run._element.find(qn('w:rPr'))
        cites.sort(key=lambda x: x[0])
        new_elems = []
        last = 0
        for ms, me, new_n, is_kept in cites:
            if ms > last:
                rb = OxmlElement('w:r')
                if orig_rPr is not None:
                    rb.append(copy.deepcopy(orig_rPr))
                tb = OxmlElement('w:t'); tb.text = text[last:ms]; tb.set(XML_SPACE, 'preserve')
                rb.append(tb)
                new_elems.append(rb)
            if is_kept:
                new_elems.append(make_superscript_bracket('['))
                new_elems.append(make_endnote_ref_run(new_n + 1))
                new_elems.append(make_superscript_bracket(']'))
                replaced += 1
            # 重复标注：不写入任何元素，等效于从文中删除
            last = me
        if last < len(text):
            ra = OxmlElement('w:r')
            if orig_rPr is not None:
                ra.append(copy.deepcopy(orig_rPr))
            ta = OxmlElement('w:t'); ta.text = text[last:]; ta.set(XML_SPACE, 'preserve')
            ra.append(ta)
            new_elems.append(ra)
        for j, elem in enumerate(new_elems):
            parent.insert(idx + j, elem)
        parent.remove(run._element)

    report.append(f'✅ 文中标注 {replaced} 处 → 真尾注 {len(final_refs)} 条（一文一引，编号 1-{len(final_refs)}）')
    return final_refs


def patch_settings_endnote_numfmt(doc):
    """settings.xml 全局尾注编号设为十进制（Word 默认是小写罗马数字）。"""
    settings_part = doc.settings.element
    ep = settings_part.find(qn('w:endnotePr'))
    if ep is None:
        ep = OxmlElement('w:endnotePr')
        settings_part.insert(0, ep)
    for n in ep.findall(qn('w:numFmt')):
        ep.remove(n)
    n = OxmlElement('w:numFmt'); n.set(qn('w:val'), 'decimal')
    ep.insert(0, n)


# ==================== markdown 解析 ====================

def parse_md(path):
    text = open(path, encoding='utf-8').read()
    lines = text.split('\n')

    data = {
        'title': None, 'abstract': None, 'keywords': None,
        'body_blocks': [],   # ('h1'|'h2'|'h3'|'p'|'intro_p', text)
        'refs_raw': [],      # [(old_n, text)]
    }
    in_refs = False
    para_buf = []

    def flush_para():
        if para_buf:
            joined = ' '.join(para_buf).strip()
            if joined:
                data['body_blocks'].append(('p', joined))
            para_buf.clear()

    def clean(t):
        # 去掉 markdown 强调符号，保留纯文本（[N] 保留）
        return re.sub(r'\*\*', '', t).strip()

    for raw in lines:
        line = raw.strip()
        if not line or line == '---':
            flush_para()
            continue
        if in_refs:
            m = re.match(r'^\[(\d+)\]\s*(.+)$', line)
            if m:
                data['refs_raw'].append((int(m.group(1)), clean(m.group(2))))
            continue
        if line.startswith('# ') and data['title'] is None:
            flush_para()
            data['title'] = clean(line[2:])
            continue
        c = clean(line)
        if re.match(r'^摘\s*要：', c):
            flush_para()
            data['abstract'] = re.sub(r'^摘\s*要：', '', c).strip()
            continue
        if re.match(r'^关键词：', c):
            flush_para()
            data['keywords'] = re.sub(r'^关键词：', '', c).strip()
            continue
        if re.match(r'^参考文献：?$', c):
            flush_para()
            in_refs = True
            continue
        m2 = re.match(r'^##\s+(.*)$', line)
        m3 = re.match(r'^###\s+(.*)$', line)
        m4 = re.match(r'^####\s+(.*)$', line)
        if m4:
            flush_para(); data['body_blocks'].append(('h3', clean(m4.group(1)))); continue
        if m3:
            flush_para(); data['body_blocks'].append(('h2', clean(m3.group(1)))); continue
        if m2:
            flush_para()
            h = clean(m2.group(1))
            if h == '引言':
                # 内置模板：引言不排标题，其下段落直接作正文开头
                data['body_blocks'].append(('intro_mark', ''))
            else:
                data['body_blocks'].append(('h1', h))
            continue
        para_buf.append(line)
    flush_para()

    errs = []
    if not data['title']:
        errs.append('缺标题（# 标题）')
    if not data['abstract']:
        errs.append('缺摘要（**摘要：** ……）')
    if not data['keywords']:
        errs.append('缺关键词（**关键词：** ……）')
    if not data['refs_raw']:
        errs.append('缺参考文献（**参考文献：** + [N] 条目）')
    if errs:
        raise SystemExit('markdown 解析失败：' + '；'.join(errs))
    return data


# ==================== 验证 ====================

def verify(outpath, n_refs, report):
    """重开生成的 docx，逐项自检。"""
    z = zipfile.ZipFile(outpath)
    doc_xml = z.read('word/document.xml').decode('utf-8')
    en_xml = z.read('word/endnotes.xml').decode('utf-8')
    fn_xml = z.read('word/footnotes.xml').decode('utf-8')

    ok = True
    # 1. 文中 endnoteReference 数量
    n_inbody = len(re.findall(r'<w:endnoteReference', doc_xml))
    # 2. 尾注条目（不含 separator/continuationSeparator/标题尾注）
    entries = re.findall(r'<w:endnote w:id="(\d+)">(.*?)</w:endnote>', en_xml, re.S)
    content_entries = [(i, b) for i, b in entries if 'w:type=' not in b]
    heading_ok = any('参考文献' in b and '<w:endnoteRef' not in b.replace('<w:endnoteRef/>', '')
                     for _, b in entries)
    # 3. 无残留 [N] 纯文本
    texts = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', doc_xml))
    leftover = re.findall(r'\[\d+\]', texts)
    # 4. 上标数量
    n_sup = len(re.findall(r'w:val="superscript"', doc_xml))
    # 5. 脚注引用
    n_fn = len(re.findall(r'<w:footnoteReference', doc_xml))

    checks = [
        (n_inbody == n_refs, f'文中真尾注引用 {n_inbody} 个（期望 {n_refs}）'),
        (len(content_entries) == n_refs + 1, f'尾注条目 {len(content_entries)} 条（期望 {n_refs} 条文献 + 1 条标题）'),
        (heading_ok, '"参考文献"标题在尾注区内（紧贴列表不分离）'),
        (all(''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', b)).strip('[] 0123456789') for _, b in content_entries if '参考文献' not in b),
         '所有尾注条目内容非空'),
        (not leftover, f'文中无残留 [N] 纯文本' + (f'（残留：{leftover}）' if leftover else '')),
        (n_sup >= n_refs * 3, f'上标标记 {n_sup} 个（≥{n_refs * 3}，每处引用 3 个 run）'),
        (n_fn == 1, f'作者简介脚注引用 {n_fn} 个（期望 1）'),
        ('w:endnoteRef' in fn_xml or '作者简介' in fn_xml, '脚注内容已写入'),
    ]
    for passed, msg in checks:
        report.append(('✅ ' if passed else '❌ ') + msg)
        ok = ok and passed
    return ok


# ==================== 主流程 ====================

def build(md_path, meta_path, out_path):
    meta = json.load(open(meta_path, encoding='utf-8'))
    data = parse_md(md_path)
    report = []

    refs_by_old_num = {n: t for n, t in data['refs_raw']}
    dup = [n for n in sorted(set(n for n, _ in data['refs_raw']))
           if sum(1 for x, _ in data['refs_raw'] if x == n) > 1]
    if dup:
        raise SystemExit(f'错误：参考文献表中编号 {dup} 重复列出，请去重后重跑')

    doc = Document()
    normal = doc.styles['Normal']
    normal.font.name = 'Times New Roman'
    normal.font.size = Pt(10.5)
    rpr_n = normal.element.find(qn('w:rPr'))
    if rpr_n is None:
        rpr_n = OxmlElement('w:rPr'); normal.element.append(rpr_n)
    rf_n = rpr_n.find(qn('w:rFonts'))
    if rf_n is None:
        rf_n = OxmlElement('w:rFonts'); rpr_n.append(rf_n)
    rf_n.set(qn('w:eastAsia'), '宋体')
    rf_n.set(qn('w:ascii'), 'Times New Roman')
    rf_n.set(qn('w:hAnsi'), 'Times New Roman')

    # ====== 第一节：单栏 标题区 ======
    sec1 = doc.sections[0]
    setup_section(sec1, cols=1)

    t = TPL['title']
    add_center(doc, data['title'], t['cn'], t['en'], t['sz'], t['bold'], sa=6)
    author_p = add_center(doc, meta['author'], TPL['author']['cn'], TPL['author']['en'],
                          TPL['author']['sz'], sa=3)
    add_fn_ref(author_p)
    fn_lines = []
    if meta.get('fund'):
        fn_lines.append(meta['fund'])
    if meta.get('received'):
        fn_lines.append('收稿日期：' + meta['received'])
    fn_lines.append('作者简介：' + meta['author_bio'])
    add_author_fn_content(doc, 1, fn_lines, sz=TPL['fn_author']['sz'])
    a = TPL['affiliation']
    add_center(doc, meta['affiliation'], a['cn'], a['en'], a['sz'], sa=6)
    add_center(doc, '', '宋体', 'Times New Roman', 10.5)

    add_mixed(doc, [('摘 要：', 'abs_label'), (data['abstract'], 'abs_content')])
    add_mixed(doc, [('关键词：', 'abs_label'), (data['keywords'], 'abs_content')])
    clc = meta.get('clc', 'G641')
    doc_code = meta.get('doc_code', 'A')
    add_mixed(doc, [('中图分类号：', 'clc_label'), (clc, 'clc_content'),
                    ('    ', 'clc_content'),
                    ('文献标识码：', 'clc_label'), (doc_code, 'clc_content')])

    # ====== 第二节：双栏 正文 ======
    sec2 = doc.add_section()
    sec2.start_type = WD_SECTION.CONTINUOUS
    setup_section(sec2, cols=2)

    for kind, text in data['body_blocks']:
        if kind == 'h1':
            add_body(doc, text, 'h1', indent=False, sb=3, sa=1)
        elif kind == 'h2':
            add_body(doc, text, 'h2', indent=False, sb=2, sa=1)
        elif kind == 'h3':
            add_body(doc, text, 'h3', indent=False, sb=2, sa=1)
        elif kind in ('p', 'intro_mark'):
            if kind == 'p':
                add_body(doc, text, 'body')
            # intro_mark 无内容，仅标记

    # ====== 尾注处理（必须在正文写完、英文节之前） ======
    final_refs = postprocess_endnotes(doc, refs_by_old_num, len(refs_by_old_num), report)

    # ====== 第三节：单栏 英文部分 + 文末作者简介 ======
    sec3 = doc.add_section()
    sec3.start_type = WD_SECTION.CONTINUOUS
    setup_section(sec3, cols=1)

    en = meta.get('english')
    if en:
        et = TPL['en_title']
        add_center(doc, en['title'], '宋体', et['en'], et['sz'], et['bold'], sb=6, sa=6)
        ea = TPL['en_author']
        add_center(doc, en['author'], '宋体', ea['en'], ea['sz'], sa=3)
        ef = TPL['en_affil']
        add_center(doc, en['affiliation'], '宋体', ef['en'], ef['sz'], sa=6)
        eab = TPL['en_abs']
        # Abstract / Key words
        p_abs = doc.add_paragraph()
        r1 = p_abs.add_run('Abstract: ')
        set_run_font(r1, '宋体', eab['en'], eab['sz'], True)
        r2 = p_abs.add_run(en['abstract'])
        set_run_font(r2, '宋体', eab['en'], eab['sz'], False)
        set_spacing(p_abs, sb=3, sa=3)
        p_kw = doc.add_paragraph()
        r3 = p_kw.add_run('Key words: ')
        set_run_font(r3, '宋体', eab['en'], eab['sz'], True)
        r4 = p_kw.add_run(en['keywords'])
        set_run_font(r4, '宋体', eab['en'], eab['sz'], False)
        set_spacing(p_kw, sb=3, sa=6)

    # 文末作者简介（含联系方式）
    bio = '作者简介：' + meta['author_bio']
    if meta.get('contact'):
        bio += meta['contact']
    b = TPL['bio_end']
    add_body(doc, bio, 'bio_end', indent=False, sa=3)

    patch_settings_endnote_numfmt(doc)
    doc.save(out_path)

    ok = verify(out_path, len(final_refs), report)
    print('\n'.join(report))
    print()
    print('输出：' + out_path)
    if not ok:
        raise SystemExit('❌ 自检未全部通过，请检查后再交付')
    print('✅ 自检全部通过')


def main():
    ap = argparse.ArgumentParser(description='马理论科期刊投稿式排版（内置双栏期刊模板）')
    ap.add_argument('--md', required=True, help='正文 markdown')
    ap.add_argument('--meta', required=True, help='元信息 JSON（作者/单位/作者简介等）')
    ap.add_argument('--out', required=True, help='输出 docx 路径')
    args = ap.parse_args()
    build(args.md, args.meta, args.out)


if __name__ == '__main__':
    main()
