"""Reusable numbered undergraduate preset; private identity data stays in task-local metadata."""
from pathlib import Path
import argparse
import copy
import importlib.util
import json
import re
import tempfile
import zipfile
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from lxml import etree

HERE = Path(__file__).resolve().parent
NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
SPECIAL = {'前言', '绪论', '结论', '致谢', '参考文献', '附录', '外文翻译'}

def special(title):
    return title in SPECIAL or title.startswith(('附录', '附件'))

def helper(name):
    spec = importlib.util.spec_from_file_location('ug_' + name, HERE / (name + '.py'))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module

def load_profile(meta):
    profile = json.loads((HERE.parent / 'presets/undergraduate-01.json').read_text(encoding='utf-8'))
    unknown = set(meta.get('format_overrides', {})) - {'page', 'heading', 'keywords', 'bibliography'}
    if unknown:
        raise ValueError('未知格式覆盖分组：' + ', '.join(sorted(unknown)))
    for group in ('page', 'heading', 'keywords', 'bibliography'):
        if group in meta.get('format_overrides', {}):
            if not isinstance(meta['format_overrides'][group], dict):
                raise ValueError('格式覆盖项必须是对象')
            for key, value in meta['format_overrides'][group].items():
                if key not in profile[group]:
                    raise ValueError(f'未知格式参数：{group}.{key}')
                if isinstance(profile[group][key], bool):
                    if not isinstance(value, bool):
                        raise ValueError(f'格式参数必须为布尔值：{group}.{key}')
                elif isinstance(profile[group][key], (int, float)) and (isinstance(value, bool) or not isinstance(value, (int, float)) or value <= 0):
                    raise ValueError(f'格式参数必须为正数：{group}.{key}')
                profile[group][key] = value
    return profile

def font(run, cn='宋体', size=12, bold=False, latin='Times New Roman'):
    run.font.name = latin
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0)
    rf = run._element.get_or_add_rPr().get_or_add_rFonts()
    for key, value in (('ascii', latin), ('hAnsi', latin), ('eastAsia', cn)):
        rf.set(qn('w:' + key), value)

def snap(paragraph, enabled=False):
    pr = paragraph._p.get_or_add_pPr()
    node = pr.find(qn('w:snapToGrid'))
    if node is None:
        node = OxmlElement('w:snapToGrid'); pr.append(node)
    node.set(qn('w:val'), '1' if enabled else '0')

def paragraph(doc, text='', *, size=12, cn='宋体', align=None, bold=False, indent=0, spacing=1.25, after=0, before=0, style=None):
    p = doc.add_paragraph(style=style)
    p.paragraph_format.line_spacing = spacing
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.first_line_indent = Pt(indent * size)
    if align is not None:
        p.alignment = align
    snap(p)
    font(p.add_run(text), cn, size, bold)
    return p

def heading(doc, text, profile, english=False, toc=True):
    h = profile['heading']
    return paragraph(doc, text, size=h['size_pt'], cn='Times New Roman' if english else h['font_cn'],
                     align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, spacing=h['line_spacing'],
                     after=h['space_after_pt'], style='Heading 1' if toc else 'Title')

def field(paragraph, instruction, cached=''):
    node = OxmlElement('w:fldSimple'); node.set(qn('w:instr'), instruction)
    r = OxmlElement('w:r'); t = OxmlElement('w:t'); t.text = cached; r.append(t); node.append(r)
    paragraph._p.append(node)
    return node

def set_headers(doc, kinds, profile):
    first_body = True
    for section, kind in zip(doc.sections, kinds):
        page = profile['page']
        section.page_width, section.page_height = Cm(page['width_cm']), Cm(page['height_cm'])
        section.top_margin, section.bottom_margin = Cm(page['top_cm']), Cm(page['bottom_cm'])
        section.left_margin, section.right_margin = Cm(page['left_cm']), Cm(page['right_cm'])
        section.header.is_linked_to_previous = False
        section.footer.is_linked_to_previous = False
        section.different_first_page_header_footer = False
        for container in (section.header, section.footer):
            for p in container.paragraphs:
                p.clear()
        for old in section._sectPr.findall(qn('w:pgNumType')):
            section._sectPr.remove(old)
        if kind != 'cover':
            number = OxmlElement('w:pgNumType')
            number.set(qn('w:fmt'), profile['page_number']['front'] if kind == 'front' else 'decimal')
            if kind == 'front' or first_body:
                number.set(qn('w:start'), '1')
            if kind == 'body':
                first_body = False
            section._sectPr.append(number)
            p = section.header.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            f = field(p, ' PAGE ', '1')
            from docx.text.run import Run
            for node in f.findall(qn('w:r')):
                font(Run(node, p), 'Times New Roman', profile['page_number']['size_pt'])

def move_note_after_punctuation(doc):
    punctuation = re.compile(r'^[”’」』）》】]*[，。；：！？,.!?;:]+')
    for p in doc.paragraphs:
        children = list(p._p)
        for child in children:
            if child.find('.//' + qn('w:footnoteReference')) is None:
                continue
            following = child.getnext()
            while following is not None and following.find('.//' + qn('w:footnoteReference')) is not None:
                following = following.getnext()
            if following is None:
                continue
            texts = following.findall('.//' + qn('w:t'))
            if not texts:
                continue
            match = punctuation.match(texts[0].text or '')
            if match:
                mark = OxmlElement('w:r'); t = OxmlElement('w:t'); t.text = match.group(0); mark.append(t)
                child.addprevious(mark)
                texts[0].text = texts[0].text[match.end():]

def format_notes(doc):
    helper('render_docx').set_footnote_per_page_restart(doc)
    used_ids = {ref.get(qn('w:id')) for ref in doc._element.findall('.//w:footnoteReference', NS)}
    for rel in doc.part.rels.values():
        if not str(rel.reltype).endswith('/footnotes'):
            continue
        root = etree.fromstring(rel.target_part.blob)
        for note in root.findall('w:footnote', NS):
            if int(note.get(qn('w:id'))) <= 0:
                continue
            if note.get(qn('w:id')) not in used_ids:
                root.remove(note)
                continue
            paragraphs = note.findall('w:p', NS)
            for index, p in enumerate(paragraphs):
                pr = p.find(qn('w:pPr'))
                if pr is None:
                    pr = OxmlElement('w:pPr'); p.insert(0, pr)
                if pr.find(qn('w:keepLines')) is None:
                    pr.append(OxmlElement('w:keepLines'))
                if index < len(paragraphs) - 1 and pr.find(qn('w:keepNext')) is None:
                    pr.append(OxmlElement('w:keepNext'))
        rel.target_part._blob = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

def identity_terms(meta):
    return [str(meta[k]).strip() for k in ('author', 'advisor', 'student_id', 'institution', 'faculty') if meta.get(k)]

def inspect_blind(path, meta):
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if name.endswith('.xml') and name.startswith(('word/', 'docProps/')):
                root = etree.fromstring(archive.read(name))
                text = ''.join(root.itertext())
                if any(term in text for term in identity_terms(meta)):
                    raise ValueError('盲审稿仍含本地身份字段，请在源稿核对；未发布输出')

def clear_identity_properties(path):
    # Word may reinsert its configured user name on save; clear it after the final save.
    with zipfile.ZipFile(path) as archive:
        entries = [(item, archive.read(item.filename)) for item in archive.infolist()]
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as archive:
        for item, content in entries:
            if item.filename == 'docProps/core.xml':
                root = etree.fromstring(content)
                for node in root:
                    if etree.QName(node).localname in ('creator', 'lastModifiedBy', 'keywords', 'description', 'category'):
                        node.text = ''
                content = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
            archive.writestr(item, content)

def verify_final_circles(path, pages):
    with zipfile.ZipFile(path) as archive:
        root = etree.fromstring(archive.read('word/document.xml'))
    refs = root.findall('.//w:footnoteReference', NS)
    if len(refs) != len(pages):
        raise ValueError('最终脚注数量与Word分页记录不一致，未发布')
    counters = {}
    for ref, page in zip(refs, pages):
        counters[page] = counters.get(page, 0) + 1
        if counters[page] > 20:
            raise ValueError('本科格式一同页脚注超过20个，请调整内容或版式后重新验收')
        actual = ''.join(ref.getparent().itertext()).strip()
        if actual != chr(0x2460 + counters[page] - 1):
            raise ValueError('Word最终分页与圈码不一致，未发布；请调整版式后重新生成')

def make_document(raw_path, meta, profile, blind=False, has_cover=False):
    doc = Document(raw_path)
    body = doc._element.body
    originals = [copy.deepcopy(node) for node in body if node.tag != qn('w:sectPr')]
    for node in list(body):
        if node.tag != qn('w:sectPr'):
            body.remove(node)
    normal = doc.styles['Normal']
    normal.font.name, normal.font.size = 'Times New Roman', Pt(12)
    kinds = ['cover']
    if has_cover and not blind:
        p = paragraph(doc, '本地封面待载入')
        start = OxmlElement('w:bookmarkStart'); start.set(qn('w:id'), '9000'); start.set(qn('w:name'), 'LocalCoverSlot')
        end = OxmlElement('w:bookmarkEnd'); end.set(qn('w:id'), '9000')
        p._p.insert(1 if p._p.find(qn('w:pPr')) is not None else 0, start)
        p._p.append(end)
    else:
        paragraph(doc, '大学本科毕业论文' if not blind else '本科毕业论文匿名评审稿', size=22, cn='黑体', bold=True,
                  align=WD_ALIGN_PARAGRAPH.CENTER, before=80, after=50, style='Title')
        paragraph(doc, meta['title'], size=18, cn='黑体', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, after=40)
        if not blind:
            labels = {'institution': '学校', 'faculty': '院系', 'major': '专业', 'author': '作者', 'student_id': '学号', 'advisor': '指导教师', 'submission_date': '提交日期'}
            for key, label in labels.items():
                if meta.get(key):
                    paragraph(doc, f'{label}：{meta[key]}', align=WD_ALIGN_PARAGRAPH.CENTER, after=12)
    doc.add_section(WD_SECTION.NEW_PAGE); kinds.append('front')
    paragraph(doc, meta['title'], size=15, cn='黑体', bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, after=11)
    heading(doc, '摘要', profile)
    paragraph(doc, meta['abstract_zh'], indent=2)
    paragraph(doc, '关键词：' + '；'.join(meta['keywords_zh']), cn='黑体', bold=True,
              before=15 * profile['keywords']['gap_lines'])
    doc.add_page_break()
    paragraph(doc, meta['english_title'], size=15, cn='Times New Roman', align=WD_ALIGN_PARAGRAPH.CENTER)
    paragraph(doc, '', cn='Times New Roman')
    heading(doc, 'ABSTRACT', profile, english=True)
    paragraph(doc, meta['abstract_en'], cn='Times New Roman', indent=2)
    paragraph(doc, 'Key words: ' + '; '.join(meta['keywords_en']), cn='Times New Roman', bold=True,
              before=15 * profile['keywords']['gap_lines'])
    doc.add_page_break()
    heading(doc, '目录', profile, toc=False)
    p = paragraph(doc)
    field(p, ' TOC \\o "1-3" \\h \\z \\u ', '目录在 Word 中更新')
    for level in (1, 2, 3):
        name = f'TOC {level}'
        style = doc.styles[name] if name in doc.styles else doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        style.font.name, style.font.size = 'Times New Roman', Pt(12)
        style.element.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'), '宋体')
    chapter = 0
    subsections = [0, 0, 0, 0, 0]
    current = None
    skip = False
    titles = []
    reference_count = 0
    body_chars = 0
    from docx.text.paragraph import Paragraph
    for node in originals:
        p = Paragraph(node, doc._body) if node.tag == qn('w:p') else None
        text = p.text.strip() if p else ''
        level = int(re.search(r'(\d+)$', p.style.style_id).group(1)) if p and p.style.style_id.startswith('Heading') and re.search(r'\d+$', p.style.style_id) else None
        if level and level > 5:
            raise ValueError('该预置只定义到第五级标题，请先整理层级')
        if level and node.findall('.//w:footnoteReference', NS):
            raise ValueError('标题中有脚注，请先将引注移到相应正文，避免重编号时丢失引用')
        if level == 1:
            clean = re.sub(r'^(?:第[一二三四五六七八九十百\d]+章\s*|[一二三四五六七八九十]+、\s*|\d+[、.．]?\s+)', '', text).strip()
            current = clean
            skip = blind and clean == '致谢'
            if skip:
                continue
            doc.add_section(WD_SECTION.NEW_PAGE); kinds.append('body')
            titles.append(clean)
            if not special(clean):
                chapter += 1; subsections = [0, 0, 0, 0, 0]
                p.text = f'{chapter} {clean}'
            else:
                p.text = clean
        if skip:
            continue
        if current not in {'致谢', '参考文献', '附录', '外文翻译'} and not (current or '').startswith('附录') and not level:
            body_chars += len(re.sub(r'\s+', '', text))
        if current is None and text:
            raise ValueError('正文须从一级前言/绪论标题开始；标题和摘要放入本地meta.json')
        if p:
            pf = p.paragraph_format
            pf.space_before, pf.space_after = Pt(0), Pt(0)
            pf.line_spacing = profile['body']['line_spacing']
            snap(p, False)
            if level:
                if level > 1 and special(current):
                    raise ValueError('前言、结论、致谢、参考文献等只列一级标题；其下请用正文段落')
                pf.keep_with_next = True
                pf.first_line_indent = Pt(0)
                pf.line_spacing = profile['heading']['line_spacing']
                pf.space_after = Pt(profile['heading']['space_after_pt'])
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if level == 1 else WD_ALIGN_PARAGRAPH.LEFT
                if level in (2, 3) and chapter and not special(current):
                    index = level - 1
                    subsections[index] += 1
                    for lower in range(index + 1, len(subsections)):
                        subsections[lower] = 0
                    clean = re.sub(r'^(?:\d+(?:\.\d+)*[、.．]?\s+|[（(][一二三四五六七八九十\d]+[）)]\s*)', '', text)
                    if level == 3 and not subsections[1]:
                        raise ValueError('三级标题前缺少二级标题')
                    p.text = '.'.join(map(str, [chapter] + subsections[1:level])) + ' ' + clean
                if level in (4, 5):
                    index = level - 1
                    if not subsections[index - 1]:
                        raise ValueError('下级标题前缺少上一级标题')
                    subsections[index] += 1
                    if subsections[index] > 26:
                        raise ValueError('字母层级标题超过26项，请调整结构')
                    for lower in range(index + 1, len(subsections)):
                        subsections[lower] = 0
                    clean = re.sub(r'^[A-Za-z][.．]\s*', '', text)
                    letter = chr((65 if level == 4 else 97) + subsections[index] - 1)
                    p.text = f'{letter}. {clean}'
                    pf.first_line_indent = Pt(24)
                for run in p.runs:
                    font(run, profile['heading']['font_cn'], profile['heading']['size_pt'] if level == 1 else profile['heading']['subordinate_size_pt'], True)
                    run.font.italic = False
            elif current == '参考文献':
                b = profile['bibliography']
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                pf.left_indent = Pt(b['hanging_chars'] * b['size_pt'])
                pf.first_line_indent = -pf.left_indent
                pf.line_spacing = b['line_spacing']; snap(p, b['snap_to_grid'])
                for run in p.runs:
                    font(run, '宋体', b['size_pt'])
                reference_count += bool(re.match(r'^\s*[\[［]\d+[\]］]', text))
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                pf.first_line_indent = Pt(24)
                for run in p.runs:
                    font(run)
        body.insert(len(body) - 1, node)
    if not any(name in titles for name in ('前言', '绪论')) or '结论' not in titles or '参考文献' not in titles:
        raise ValueError('正文必须包含前言/绪论、结论、参考文献一级标题')
    if not blind and '致谢' not in titles:
        raise ValueError('正式稿缺少致谢部分；盲审稿使用 --blind')
    if len(set(titles)) != len(titles) or titles[0] not in ('前言', '绪论') or chapter < 1:
        raise ValueError('须从前言/绪论开始，包含本论章节，一级标题不能重复')
    conclusion, refs = titles.index('结论'), titles.index('参考文献')
    if any(not special(t) for t in titles[conclusion + 1:]) or conclusion >= refs:
        raise ValueError('本论章节须在结论前，参考文献须在结论后')
    if not blind and not conclusion < titles.index('致谢') < refs:
        raise ValueError('正式稿顺序须为结论、致谢、参考文献')
    set_headers(doc, kinds, profile)
    move_note_after_punctuation(doc)
    format_notes(doc)
    doc.core_properties.author = ''
    doc.core_properties.last_modified_by = ''
    doc.core_properties.title = meta['title']
    doc.core_properties.subject = profile['name']
    return doc, {'sections': kinds, 'chapter_count': chapter, 'reference_count': reference_count,
                 'body_chars_estimate': body_chars, 'blind': blind}

def finalize_word(path, cover_docx=None, meta=None, export_pdf=None):
    import win32com.client
    app = win32com.client.DispatchEx('Word.Application')
    app.Visible = False; app.DisplayAlerts = 0
    doc = None
    try:
        doc = app.Documents.Open(str(Path(path).resolve()), ReadOnly=False, AddToRecentFiles=False)
        if cover_docx:
            cover = app.Documents.Open(str(Path(cover_docx).resolve()), ReadOnly=True, AddToRecentFiles=False)
            try:
                cover.Repaginate()
                if cover.Sections.Count != 1 or cover.ComputeStatistics(2) != 1:
                    raise ValueError('本地封面须为单节单页Word；多页前置材料需先在任务内整理')
                source = cover.Range(0, cover.Content.End - 1)
                doc.Bookmarks.Item('LocalCoverSlot').Range.FormattedText = source.FormattedText
            finally:
                cover.Close(False)
        doc.Fields.Update()
        for toc in doc.TablesOfContents:
            toc.Update()
        doc.Repaginate()
        for toc in doc.TablesOfContents:
            toc.UpdatePageNumbers()
            toc.Range.Font.Name = 'Times New Roman'
            toc.Range.Font.NameFarEast = '宋体'
            toc.Range.Font.Size = 12
            toc.Range.Font.Bold = 0
            toc.Range.Font.Italic = 0
            toc.Range.Font.Color = 0
        pages = []
        for note in doc.Footnotes:
            start, end = note.Range.Duplicate, note.Range.Duplicate
            start.Collapse(1); end.Collapse(0)
            reference_page = note.Reference.Information(3)
            if start.Information(3) != end.Information(3) or start.Information(3) != reference_page:
                raise ValueError('脚注跨页或与引用不在同页，未发布；请缩短该条注释或调整本地版式')
            pages.append(reference_page)
        if doc.TablesOfContents.Count != 1:
            raise ValueError('自动目录未形成，未发布')
        doc.BuiltInDocumentProperties(3).Value = ''
        doc.Save()
        if export_pdf:
            doc.ExportAsFixedFormat(str(Path(export_pdf).resolve()), 17)
        return {'word_version': app.Version, 'pages': doc.ComputeStatistics(2), 'footnote_pages': pages, 'toc_count': doc.TablesOfContents.Count}
    finally:
        if doc is not None:
            doc.Close(False)
        app.Quit()

def build(md_path, meta_path, out_path, *, blind=False, overwrite=False, draft=False, cover_docx=None, pdf_path=None):
    target = Path(out_path).resolve()
    report_path = target.with_suffix('.report.json')
    pdf_target = Path(pdf_path).resolve() if pdf_path else None
    destinations = [target, report_path] + ([pdf_target] if pdf_target else [])
    if len(set(destinations)) != len(destinations):
        raise ValueError('Word、PDF和验收记录不能是同一路径')
    if any(p.exists() for p in destinations) and not overwrite:
        raise FileExistsError('输出或验收记录已存在；另存或在获授权后使用 --overwrite')
    if Path(md_path).suffix.lower() != '.md':
        raise ValueError('正文输入须为.md；已有Word用文档编辑流程')
    if draft and pdf_path:
        raise ValueError('--draft不能导出已验收PDF')
    meta = json.loads(Path(meta_path).read_text(encoding='utf-8-sig'))
    for key in ('title', 'english_title', 'abstract_zh', 'abstract_en'):
        if not isinstance(meta.get(key), str) or not meta[key].strip():
            raise ValueError(f'缺少本地元信息：{key}')
    for key in ('keywords_zh', 'keywords_en'):
        if not isinstance(meta.get(key), list) or not 3 <= len(meta[key]) <= 5 or not all(isinstance(x, str) and x.strip() for x in meta[key]):
            raise ValueError(f'{key} 应为3—5个非空关键词')
    if len(meta['keywords_zh']) != len(meta['keywords_en']):
        raise ValueError('中英文关键词数量须相同')
    if not draft:
        try:
            import win32com.client
        except ImportError as exc:
            raise RuntimeError('自动目录、分页与圈码验收需已安装pywin32和Word的Python；也可显式 --draft 生成待验收稿') from exc
    if cover_docx and draft:
        raise ValueError('本地封面载入需要Word；--draft不载入本地封面')
    profile = load_profile(meta)
    target.parent.mkdir(parents=True, exist_ok=True)
    report = {'preset': profile['id'], 'name': profile['name'], 'needs_confirmation': profile['needs_confirmation'], 'word_verified': not draft}
    base = helper('build_docx')
    with tempfile.TemporaryDirectory(prefix='.undergraduate-', dir=target.parent) as directory:
        raw, ready = Path(directory) / 'raw.docx', Path(directory) / 'ready.docx'
        base.run_pandoc(Path(md_path), raw)
        doc, checks = make_document(raw, meta, profile, blind, bool(cover_docx))
        doc.save(ready)
        if blind:
            inspect_blind(ready, meta)
        if not draft:
            if doc._element.findall('.//' + qn('w:footnoteReference')):
                helper('footnote_circle_marks').process(ready, ready)
            finalize_word(ready, None if blind else cover_docx, meta)
            if doc._element.findall('.//' + qn('w:footnoteReference')):
                helper('renumber_footnotes_per_page').renumber_docx(ready, ready)
            # Export only after document verification, to a temporary PDF first.
            temporary_pdf = Path(directory) / 'verified.pdf' if pdf_path else None
            report['word'] = finalize_word(ready, export_pdf=temporary_pdf)
            verify_final_circles(ready, report['word']['footnote_pages'])
            clear_identity_properties(ready)
            if blind:
                inspect_blind(ready, meta)
        Document(ready)
        report.update(checks)
        report['content_checks'] = {'title_within_20_chars': len(meta['title']) <= profile['content_checks']['title_max_chars'],
                                    'body_min_chars_met': checks['body_chars_estimate'] >= profile['content_checks']['body_min_chars'],
                                    'reference_min_count_met': checks['reference_count'] >= profile['content_checks']['reference_min_count']}
        # These are task quality checks, not excuses to invent or pad research content.
        if pdf_path and not draft:
            pdf_target.parent.mkdir(parents=True, exist_ok=True)
            # Publish paired PDF only after all validation succeeds.
            import shutil
            with tempfile.TemporaryDirectory(prefix='.pdf-publish-', dir=pdf_target.parent) as pdf_temp:
                sibling = Path(pdf_temp) / pdf_target.name
                shutil.copy2(temporary_pdf, sibling)
                base.publish_file(sibling, pdf_target, overwrite)
        report_temp = Path(directory) / 'report.json'
        report_temp.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
        base.publish_file(ready, target, overwrite)
        base.publish_file(report_temp, report_path, overwrite)
    print(json.dumps(report, ensure_ascii=False))
    return report

def main():
    parser = argparse.ArgumentParser(description='大学本科毕业论文格式一')
    parser.add_argument('--md', required=True)
    parser.add_argument('--meta', required=True)
    parser.add_argument('--out', required=True)
    parser.add_argument('--blind', action='store_true')
    parser.add_argument('--overwrite', action='store_true')
    parser.add_argument('--draft', action='store_true', help='仅生成待Word验收稿，不声称目录和分页已核实')
    parser.add_argument('--cover-docx', help='任务本地单节单页封面，绝不收入公共仓库')
    parser.add_argument('--pdf', help='同时导出已验收PDF')
    args = parser.parse_args()
    build(args.md, args.meta, args.out, blind=args.blind, overwrite=args.overwrite, draft=args.draft, cover_docx=args.cover_docx, pdf_path=args.pdf)

if __name__ == '__main__':
    main()
