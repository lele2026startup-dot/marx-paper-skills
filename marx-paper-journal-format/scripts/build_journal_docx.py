#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
build_journal_docx.py —— 马理论科期刊投稿式排版（内置双栏期刊模板）

把结构化 markdown 草稿 + meta.json 元信息，生成符合内置双栏期刊模板的 Word 文档：
- 顺序编码制真尾注（endnoteReference 链接 endnotes.xml，可点击跳转、自动编号）
- 文中上标 [N] 方括号（vertAlign=superscript，三个 run：[ + endnoteReference + ]）
- 每处引用保留：首次建立真尾注；重复来源用带链接的 NOTEREF 域，文献条目只列一次
- 自动重编号：按文中首次出现顺序重排编号，参考文献表同步重排
- "参考文献"标题放在第一条有效尾注首段；孤立未引用尾注会被Word忽略
- 作者简介做成第一页真脚注（自动编号），文末另附作者简介+联系方式段
- 默认不覆盖已有输出；只有临时文件验证通过后才发布，授权覆盖用 --overwrite

用法：
    python build_journal_docx.py --md 论文正文.md --meta meta.json --out 投稿版.docx
    可选：--no-english 不出英文部分；--overwrite 授权备份后覆盖

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
    clc          中图分类号，未提供时不猜填
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
import os
import tempfile
from pathlib import Path

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


def add_endnote_entry(doc, en_id, ref_text):
    """一条尾注文献条目：[ endnoteRef ] 文献内容，小五号，单倍行距。"""
    en_part = ensure_endnotes_part(doc)
    root = etree.fromstring(en_part.blob)
    st = TPL['ref_entry']
    sz_half = str(int(st['sz'] * 2))

    en = etree.SubElement(root, qn('w:endnote'))
    en.set(qn('w:id'), str(en_id))
    if en_id == 2:
        heading = etree.SubElement(en, qn('w:p'))
        heading_pr = etree.SubElement(heading, qn('w:pPr'))
        etree.SubElement(heading_pr, qn('w:keepNext'))
        heading_run = etree.SubElement(heading, qn('w:r'))
        heading_rpr = etree.SubElement(heading_run, qn('w:rPr'))
        etree.SubElement(heading_rpr, qn('w:b'))
        fonts = etree.SubElement(heading_rpr, qn('w:rFonts'))
        fonts.set(qn('w:eastAsia'), TPL['ref_heading']['cn'])
        size = etree.SubElement(heading_rpr, qn('w:sz'))
        size.set(qn('w:val'), str(int(TPL['ref_heading']['sz'] * 2)))
        etree.SubElement(heading_run, qn('w:t')).text = '参考文献'
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


def postprocess_endnotes(doc, refs_by_old_num, n_refs, report):
    """One true endnote per source; every repeated occurrence is a linked NOTEREF field."""
    records = []
    for para in doc.paragraphs:
        for run in para.runs:
            for match in re.finditer(r'\[(\d+)\]', run.text or ''):
                records.append((run, match.start(), match.end(), int(match.group(1))))
    cited = {item[3] for item in records}
    missing = sorted(cited - set(refs_by_old_num))
    if missing:
        raise ValueError(f'文中原编号 {missing} 没有对应参考文献条目，未生成输出')
    if not records:
        raise ValueError('正文没有可识别的引用；请检查 [N] 或 Pandoc 脚注定义')
    old_to_new, source_to_new, final_refs = {}, {}, {}
    for _, _, _, old in records:
        if old not in old_to_new:
            # Exact normalized bibliography equality only; never merge versions/pages heuristically.
            key = re.sub(r'\s+', ' ', refs_by_old_num[old]).strip()
            if key not in source_to_new:
                new = len(source_to_new) + 1
                source_to_new[key] = new
                final_refs[new] = refs_by_old_num[old]
            old_to_new[old] = source_to_new[key]
    unused = sorted(set(refs_by_old_num) - cited)
    if unused:
        report.append(f'ℹ 未引条目原编号 {unused} 不进入引文表')
    ensure_endnotes_part(doc)
    for number, text in final_refs.items():
        add_endnote_entry(doc, number + 1, text)
    grouped = {}
    for run, start, end, old in records:
        grouped.setdefault(run._element, (run, []))[1].append((start, end, old_to_new[old]))
    seen = set()
    repeats = 0
    for element, (run, occurrences) in grouped.items():
        parent = element.getparent()
        index = list(parent).index(element)
        text = run.text or ''
        original_pr = element.find(qn('w:rPr'))
        output = []
        last = 0
        def literal(value):
            newrun = OxmlElement('w:r')
            if original_pr is not None:
                newrun.append(copy.deepcopy(original_pr))
            node = OxmlElement('w:t'); node.text = value; node.set(XML_SPACE, 'preserve')
            newrun.append(node)
            return newrun
        for start, end, number in occurrences:
            if start > last:
                output.append(literal(text[last:start]))
            output.append(make_superscript_bracket('['))
            if number not in seen:
                bookmark = OxmlElement('w:bookmarkStart')
                bookmark.set(qn('w:id'), str(number)); bookmark.set(qn('w:name'), f'cite_{number}')
                output.extend([bookmark, make_endnote_ref_run(number + 1)])
                bookmark_end = OxmlElement('w:bookmarkEnd'); bookmark_end.set(qn('w:id'), str(number))
                output.append(bookmark_end)
                seen.add(number)
            else:
                field = OxmlElement('w:fldSimple')
                field.set(qn('w:instr'), f' NOTEREF cite_{number} \\h ')
                field.append(make_superscript_bracket(str(number)))
                output.append(field)
                repeats += 1
            output.append(make_superscript_bracket(']'))
            last = end
        if last < len(text):
            output.append(literal(text[last:]))
        for offset, node in enumerate(output):
            parent.insert(index + offset, node)
        parent.remove(element)
    doc._citation_count = len(records)
    report.append(f'✅ 保留全部 {len(records)} 处引用：{len(final_refs)} 条真尾注、{repeats} 处 NOTEREF 交叉引用；无引用位置删除')
    if repeats:
        report.append('ℹ 重复引用使用 Word NOTEREF 域；交付前在目标 Word/WPS 检查域更新与跳转')
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
    for position in ep.findall(qn('w:pos')):
        ep.remove(position)
    position = OxmlElement('w:pos'); position.set(qn('w:val'), 'sectEnd')
    ep.insert(0, position)
    compat = settings_part.find(qn('w:compat'))
    if compat is None:
        compat = OxmlElement('w:compat'); settings_part.append(compat)
    if compat.find(qn('w:footnoteLayoutLikeWW8')) is None:
        compat.append(OxmlElement('w:footnoteLayoutLikeWW8'))


# ==================== markdown 解析 ====================

def parse_md(path):
    text = Path(path).read_text(encoding='utf-8-sig')
    # Minimal safe metadata support: YAML title is scalar text, not arbitrary YAML execution.
    yaml_title = None
    if text.startswith('---\n'):
        end = text.find('\n---', 4)
        if end < 0:
            raise ValueError('YAML frontmatter 缺结束标记')
        front = text[4:end]
        title = re.search(r'^title:\s*(.+)$', front, re.M)
        if title:
            yaml_title = title.group(1).strip().strip('\'\"')
        text = text[end + 4:].lstrip('\n')
    lines = text.splitlines()
    notes, cleaned = {}, []
    i = 0
    while i < len(lines):
        match = re.match(r'^\[\^([^\]]+)\]:\s*(.*)$', lines[i])
        if not match:
            cleaned.append(lines[i]); i += 1; continue
        label, value = match.groups()
        if label in notes:
            raise ValueError(f'脚注定义重复：{label}')
        values = [value]
        i += 1
        while i < len(lines) and (lines[i].startswith('    ') or lines[i].startswith('\t')):
            values.append(lines[i].strip()); i += 1
        notes[label] = ' '.join(values).strip()
        if not notes[label]:
            raise ValueError(f'脚注定义为空：{label}')
    data = {'title': yaml_title, 'abstract': None, 'keywords': None, 'body_blocks': [], 'refs_raw': []}
    in_refs = False
    paragraphs = []
    def clean(value):
        return value.replace('**', '').strip()
    def flush():
        if paragraphs:
            data['body_blocks'].append(('p', ' '.join(paragraphs).strip()))
            paragraphs.clear()
    for raw in cleaned:
        line = raw.strip()
        value = clean(line)
        heading = re.sub(r'^#{1,6}\s+', '', value)
        if re.fullmatch(r'参考文献\s*[:：]?', heading):
            flush(); in_refs = True; continue
        if not line or line == '---':
            flush(); continue
        if in_refs:
            match = re.fullmatch(r'\[(\d+)\]\s*(.+)', value)
            if not match:
                raise ValueError(f'无法识别参考文献条目：{line[:80]}')
            data['refs_raw'].append((int(match.group(1)), match.group(2))); continue
        if line.startswith('# '):
            flush()
            heading_title = clean(line[2:])
            if data['title'] and data['title'] != heading_title:
                raise ValueError('YAML title 与一级标题冲突')
            data['title'] = heading_title; continue
        match = re.match(r'^摘\s*要\s*[:：]\s*(.*)', value)
        if match:
            flush(); data['abstract'] = match.group(1); continue
        match = re.match(r'^关键词\s*[:：]\s*(.*)', value)
        if match:
            flush(); data['keywords'] = match.group(1); continue
        match = re.match(r'^(#{2,4})\s+(.*)', line)
        if match:
            flush()
            kind = {2: 'h1', 3: 'h2', 4: 'h3'}[len(match.group(1))]
            name = clean(match.group(2))
            data['body_blocks'].append(('intro_mark' if name == '引言' else kind, '' if name == '引言' else name))
            continue
        paragraphs.append(clean(line))
    flush()
    refs = dict(data['refs_raw'])
    if len(refs) != len(data['refs_raw']):
        raise ValueError('参考文献编号重复')
    numeric_cites = {int(match) for _, value in data['body_blocks'] for match in re.findall(r'\[(\d+)\]', value)}
    missing_numeric = sorted(numeric_cites - set(refs))
    if missing_numeric:
        raise ValueError(f'文中原编号 {missing_numeric} 没有对应参考文献条目，不能用新转换的脚注编号补缺')
    next_number = max(refs, default=0) + 1
    note_map = {}
    def note(match):
        nonlocal next_number
        label = match.group(1)
        if label not in notes:
            raise ValueError(f'脚注 [{label}] 没有定义')
        if label not in note_map:
            matching = next((n for n, value in refs.items() if value.strip() == notes[label]), None)
            number = matching if matching is not None else next_number
            if matching is None:
                refs[number] = notes[label]; next_number += 1
            note_map[label] = number
        return f'[{note_map[label]}]'
    converted = []
    for kind, value in data['body_blocks']:
        if re.search(r'\[\d+\s*[,，;；\-–]\s*\d+', value):
            raise ValueError('暂不支持合并/区间引用标注，请展开为 [1][2] 等独立标注')
        converted.append((kind, re.sub(r'\[\^([^\]]+)\]', note, value)))
    data['body_blocks'] = converted
    data['refs_raw'] = sorted(refs.items())
    for field in ('title', 'abstract', 'keywords'):
        if not data[field]:
            raise ValueError(f'缺少 {field}，请补齐标题、摘要和关键词')
    if not refs:
        raise ValueError('缺参考文献或脚注定义')
    return data


def verify(outpath, n_refs, report, expected_citations=None):
    """Validate references, targets, fields and visible content in the completed package."""
    ns = {'w': W_NS}
    with zipfile.ZipFile(outpath) as archive:
        root = etree.fromstring(archive.read('word/document.xml'))
        en = etree.fromstring(archive.read('word/endnotes.xml'))
        fn = etree.fromstring(archive.read('word/footnotes.xml'))
    endrefs = root.findall('.//w:endnoteReference', ns)
    entries = {int(node.get(qn('w:id'))): node for node in en.findall('w:endnote', ns) if int(node.get(qn('w:id'))) >= 2}
    fields = [f for f in root.findall('.//w:fldSimple', ns) if f.get(qn('w:instr'), '').strip().startswith('NOTEREF ')]
    bookmarks = {node.get(qn('w:name')): node for node in root.findall('.//w:bookmarkStart', ns)}
    end_bookmarks = {node.get(qn('w:id')) for node in root.findall('.//w:bookmarkEnd', ns)}
    texts = ''.join(node.text or '' for node in root.xpath('.//w:t[not(ancestor::w:fldSimple)]', namespaces=ns))
    targets_ok = True
    for field in fields:
        instruction = field.get(qn('w:instr'), '').split()
        target = instruction[1] if len(instruction) > 1 else ''
        bookmark = bookmarks.get(target)
        if bookmark is None or bookmark.get(qn('w:id')) not in end_bookmarks:
            targets_ok = False
        match = re.fullmatch(r'cite_(\d+)', target)
        if not match or ''.join(field.itertext()).strip() != match.group(1):
            targets_ok = False
    expected_citations = n_refs if expected_citations is None else expected_citations
    checks = [
        (len(endrefs) == n_refs and {int(r.get(qn('w:id'))) for r in endrefs} == set(entries), '真尾注与文献条目一一对应'),
        (len(endrefs) + len(fields) == expected_citations, f'全部 {expected_citations} 处引用均保留为尾注或交叉引用域'),
        (targets_ok, '重复引用域的书签目标与缓存编号有效'),
        (all(''.join(node.itertext()).strip() for node in entries.values()), '尾注条目非空'),
        (not re.search(r'\[\^|\[\d+\]', texts), '无未转换的原始引用语法'),
        (len(root.findall('.//w:footnoteReference', ns)) == 1 and '作者简介' in ''.join(fn.itertext()), '作者简介脚注在位'),
        (any('参考文献' in ''.join(n.itertext()) for n in en.findall('w:endnote', ns)), '参考文献标题在尾注区'),
        (all(ref.getparent().find('w:rPr/w:vertAlign', ns) is not None for ref in endrefs), '首次引用均有上标格式'),
    ]
    for passed, message in checks:
        report.append(('✅ ' if passed else '❌ ') + message)
    return all(passed for passed, _ in checks)



def publish_file(temp_path, out_path, overwrite=False):
    """Publish only a completed temporary file; protect existing results by default."""
    from datetime import datetime
    import os
    import shutil
    target = Path(out_path)
    temp = Path(temp_path)
    if overwrite:
        if target.exists():
            stamp = datetime.now().strftime('%Y%m%d-%H%M%S-%f')
            backup = target.with_name(target.name + '.bak.' + stamp)
            shutil.copy2(target, backup)
        os.replace(temp, target)
    else:
        # Same directory/filesystem: exclusive link fails if another writer created target.
        if os.name == 'nt':
            os.rename(temp, target)  # Windows rename atomically refuses an existing destination.
        else:
            os.link(temp, target)
            temp.unlink()


def build(md_path, meta_path, out_path, overwrite=False, no_english=False):
    target = Path(out_path)
    if target.exists() and not overwrite:
        raise FileExistsError(f"输出已存在，未覆盖：{target}；另存或获授权后使用 --overwrite")
    meta = json.loads(Path(meta_path).read_text(encoding='utf-8-sig'))
    for required in ('author', 'affiliation', 'author_bio'):
        if not isinstance(meta.get(required), str) or not meta[required].strip():
            raise ValueError(f'元信息缺少 {required}')
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
    clc = meta.get('clc')
    doc_code = meta.get('doc_code', 'A')
    if clc:
        add_mixed(doc, [('中图分类号：', 'clc_label'), (clc, 'clc_content'),
                        ('    ', 'clc_content'),
                        ('文献标识码：', 'clc_label'), (doc_code, 'clc_content')])
    else:
        report.append('ℹ 未提供中图分类号，未自动猜填')

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

    en = None if no_english else meta.get('english')
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
    target.parent.mkdir(parents=True, exist_ok=True)
    fd, temp = tempfile.mkstemp(prefix='.' + target.stem + '-', suffix='.docx', dir=target.parent)
    os.close(fd)
    try:
        doc.save(temp)
        ok = verify(temp, len(final_refs), report, doc._citation_count)
        if not ok:
            raise ValueError('自检失败，未发布输出：' + '\n'.join(report))
        publish_file(temp, target, overwrite)
    finally:
        Path(temp).unlink(missing_ok=True)
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
    ap.add_argument("--overwrite", action="store_true", help="已获授权覆盖，保存旧文件备份")
    ap.add_argument("--no-english", action="store_true", help="不生成英文部分")
    args = ap.parse_args()
    build(args.md, args.meta, args.out, overwrite=args.overwrite, no_english=args.no_english)


if __name__ == '__main__':
    main()
