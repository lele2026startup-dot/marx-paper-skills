"""Preset separation, real document structure, blind review and safe failure behavior."""
from contextlib import redirect_stdout
from pathlib import Path
import importlib.util
import io
import json
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch
import zipfile
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
TEMP_ROOT = ROOT / '_test-output'
TEMP_ROOT.mkdir(exist_ok=True)
SCRIPTS = ROOT / 'marx-paper-thesis-format/scripts'
spec = importlib.util.spec_from_file_location('undergraduate', SCRIPTS / 'build_undergraduate_docx.py')
UG = importlib.util.module_from_spec(spec)
spec.loader.exec_module(UG)
NS = UG.NS

META = {
    'title': '虚构排版验证', 'english_title': 'Fictional Format Verification',
    'abstract_zh': '这是仅验证格式的虚构摘要。', 'abstract_en': 'This fictional abstract verifies formatting.',
    'keywords_zh': ['格式', '目录', '脚注'], 'keywords_en': ['format', 'contents', 'footnotes'],
    'author': '虚构测试学生甲', 'advisor': '虚构测试导师乙', 'student_id': 'SAMPLE0001',
    'institution': '虚构测试单位', 'faculty': '虚构测试院系'
}
BODY = '''# 前言

这是前言的陈述[^a]。

# 第一章 研究起点

## 第一个问题

这是正文陈述[^b]。

### 三级问题

正文内容。

#### 四级问题

正文内容。

##### 五级问题

正文内容。

# 第二章 研究展开

## 后续问题

这是第二章内容。

# 结论

这是结论。

# 致谢

感谢虚构测试导师乙。

# 参考文献

[1] 虚构作者.虚构文献[M].仅供格式测试.

# 附录A

附录不计入正文。

[^a]: 虚构测试出处甲，仅供版面检查，不作学术引用。
[^b]: 虚构测试出处乙，仅供版面检查，不作学术引用。
'''

class Undergraduate(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='preset-test-', dir=TEMP_ROOT)
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.md, self.meta, self.raw = self.root / 'body.md', self.root / 'meta.json', self.root / 'raw.docx'
        self.md.write_text(BODY, encoding='utf-8')
        self.meta.write_text(json.dumps(META, ensure_ascii=False), encoding='utf-8')
        with redirect_stdout(io.StringIO()):
            UG.helper('build_docx').run_pandoc(self.md, self.raw)

    def document(self, blind=False, meta=None):
        data = meta or META
        return UG.make_document(self.raw, data, UG.load_profile(data), blind)

    def test_catalogue_unique_and_resolvable(self):
        entries = json.loads((ROOT / 'marx-paper/references/format-presets.json').read_text(encoding='utf-8'))['presets']
        self.assertEqual(len(entries), 3)
        self.assertEqual(len({x['id'] for x in entries}), len(entries))
        for entry in entries:
            self.assertTrue((ROOT / entry['skill'] / 'SKILL.md').is_file())
            self.assertTrue((ROOT / entry['skill'] / entry['guide']).is_file())
        self.assertEqual(entries[-1]['name'], '大学本科毕业论文格式一')

    def test_sections_headers_and_continuous_body_pages(self):
        doc, report = self.document()
        self.assertEqual(len(doc.sections), len(report['sections']))
        self.assertEqual(report['sections'][:2], ['cover', 'front'])
        for index, section in enumerate(doc.sections):
            n = section._sectPr.find(qn('w:pgNumType'))
            if index == 0:
                self.assertIsNone(n)
                self.assertNotIn('PAGE', section.header._element.xml)
            else:
                self.assertIn('PAGE', section.header._element.xml)
                self.assertEqual(n.get(qn('w:fmt')), 'upperRoman' if index == 1 else 'decimal')
                self.assertEqual(n.get(qn('w:start')), '1' if index < 3 else None)
                self.assertEqual(section.header.paragraphs[0].alignment, UG.WD_ALIGN_PARAGRAPH.RIGHT)
        self.assertEqual(report['chapter_count'], 2)

    def test_heading_numbering_and_body_not_generic(self):
        doc, _ = self.document()
        texts = [p.text for p in doc.paragraphs]
        for expected in ('1 研究起点', '1.1 第一个问题', '1.1.1 三级问题', 'A. 四级问题', 'a. 五级问题', '2 研究展开', '2.1 后续问题'):
            self.assertIn(expected, texts)
        p = next(p for p in doc.paragraphs if p.text == '正文内容。')
        self.assertEqual(p.paragraph_format.line_spacing, 1.25)
        self.assertEqual(p.paragraph_format.first_line_indent.pt, 24)
        self.assertEqual(p._p.find('w:pPr/w:snapToGrid', NS).get(qn('w:val')), '0')
        self.assertEqual(p.runs[0].font.size.pt, 12)
        self.assertIn('TOC', doc._element.xml)

    def test_true_notes_keep_relations_and_move_outside_punctuation(self):
        doc, _ = self.document()
        refs = doc._element.findall('.//w:footnoteReference', NS)
        self.assertEqual(len(refs), 2)
        for ref in refs:
            run = ref.getparent()
            self.assertTrue(run.getprevious().find(qn('w:t')).text.endswith('。'))
        output = self.root / 'notes.docx'
        doc.save(output)
        with zipfile.ZipFile(output) as archive:
            notes = etree.fromstring(archive.read('word/footnotes.xml'))
            for note in notes.findall('w:footnote', NS):
                if int(note.get(qn('w:id'))) > 0:
                    self.assertIsNotNone(note.find('w:p/w:pPr/w:keepLines', NS))
            settings = etree.fromstring(archive.read('word/settings.xml'))
            self.assertEqual(settings.find('w:footnotePr/w:numRestart', NS).get(qn('w:val')), 'eachPage')

    def test_blind_omits_thanks_and_local_identity(self):
        doc, report = self.document(blind=True)
        output = self.root / 'blind.docx'
        doc.save(output)
        UG.inspect_blind(output, META)
        self.assertTrue(report['blind'])
        self.assertNotIn('致谢', [p.text for p in doc.paragraphs])

    def test_blind_leak_rejected_instead_of_deleting_text(self):
        self.md.write_text(BODY.replace('这是结论。', '虚构测试学生甲的结论。'), encoding='utf-8')
        with redirect_stdout(io.StringIO()):
            UG.helper('build_docx').run_pandoc(self.md, self.raw)
        doc, _ = self.document(blind=True)
        output = self.root / 'leak.docx'; doc.save(output)
        with self.assertRaisesRegex(ValueError, '身份字段'):
            UG.inspect_blind(output, META)
        self.assertIn('虚构测试学生甲的结论。', [p.text for p in doc.paragraphs])

    def test_blind_removed_thanks_has_no_orphan_note_identity(self):
        self.md.write_text(BODY.replace('感谢虚构测试导师乙。', '感谢虚构测试导师乙[^thanks]。') + '\n[^thanks]: 虚构测试导师乙的私人身份说明。\n', encoding='utf-8')
        with redirect_stdout(io.StringIO()):
            UG.helper('build_docx').run_pandoc(self.md, self.raw)
        doc, _ = self.document(blind=True)
        output = self.root / 'blind-no-orphan.docx'; doc.save(output)
        UG.inspect_blind(output, META)
        with zipfile.ZipFile(output) as archive:
            notes = archive.read('word/footnotes.xml').decode('utf-8')
            self.assertNotIn('私人身份说明', notes)

    def test_heading_note_is_rejected_before_renumbering(self):
        self.md.write_text(BODY.replace('# 第一章 研究起点', '# 第一章 研究起点[^head]') + '\n[^head]: 标题引注不得丢失。\n', encoding='utf-8')
        with redirect_stdout(io.StringIO()):
            UG.helper('build_docx').run_pandoc(self.md, self.raw)
        with self.assertRaisesRegex(ValueError, '标题中有脚注'):
            self.document()

    def test_invalid_order_fails(self):
        self.md.write_text(BODY.replace('# 结论', '# TEMP').replace('# 致谢', '# 结论').replace('# TEMP', '# 致谢'), encoding='utf-8')
        with redirect_stdout(io.StringIO()):
            UG.helper('build_docx').run_pandoc(self.md, self.raw)
        with self.assertRaisesRegex(ValueError, '顺序'):
            self.document()

    def test_local_override_does_not_modify_preset(self):
        default = UG.load_profile({})
        adjusted = UG.load_profile({'format_overrides': {'page': {'left_cm': 3.5}, 'keywords': {'gap_lines': 2}}})
        self.assertEqual(adjusted['page']['left_cm'], 3.5)
        self.assertEqual(UG.load_profile({}), default)
        with self.assertRaisesRegex(ValueError, '未知'):
            UG.load_profile({'format_overrides': {'unknown': {}}})
        with self.assertRaisesRegex(ValueError, '布尔值'):
            UG.load_profile({'format_overrides': {'bibliography': {'snap_to_grid': 1}}})

    def test_appendix_not_counted_as_body(self):
        _, before = self.document()
        self.md.write_text(BODY.replace('附录不计入正文。', '附录' * 6000), encoding='utf-8')
        with redirect_stdout(io.StringIO()):
            UG.helper('build_docx').run_pandoc(self.md, self.raw)
        _, after = self.document()
        self.assertEqual(after['body_chars_estimate'], before['body_chars_estimate'])
        self.assertEqual(after['reference_count'], 1)

    def test_cli_dispatch_draft_and_records(self):
        output = self.root / 'draft.docx'
        result = subprocess.run([sys.executable, str(SCRIPTS / 'build_docx.py'), str(self.md), str(output),
                                 '--preset', 'undergraduate-01', '--meta', str(self.meta), '--blind', '--draft'], capture_output=True)
        self.assertEqual(result.returncode, 0, result.stderr.decode('utf-8', errors='replace'))
        report = json.loads(output.with_suffix('.report.json').read_text(encoding='utf-8'))
        self.assertEqual(report['preset'], 'undergraduate-01')
        self.assertFalse(report['word_verified'])
        self.assertFalse(report['content_checks']['body_min_chars_met'])
        UG.inspect_blind(output, META)

    def test_existing_sidecar_blocks_publish(self):
        output = self.root / 'protected.docx'
        record = output.with_suffix('.report.json'); record.write_text('preserve', encoding='utf-8')
        with self.assertRaises(FileExistsError):
            UG.build(self.md, self.meta, output, blind=True, draft=True)
        self.assertEqual(record.read_text(encoding='utf-8'), 'preserve')
        self.assertFalse(output.exists())

    def test_word_saved_attribute_order_and_boolean_renumber(self):
        doc, _ = self.document()
        output = self.root / 'circles.docx'; doc.save(output)
        with redirect_stdout(io.StringIO()):
            UG.helper('footnote_circle_marks').process(output, output)
        with zipfile.ZipFile(output) as archive:
            entries = [(item, archive.read(item.filename)) for item in archive.infolist()]
        with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as archive:
            for item, content in entries:
                if item.filename == 'word/document.xml':
                    xml = etree.fromstring(content)
                    for ref in xml.findall('.//w:footnoteReference', NS):
                        identifier = ref.attrib.pop(qn('w:id'))
                        ref.set(qn('w:customMarkFollows'), '1')
                        ref.set(qn('w:id'), identifier)
                    content = etree.tostring(xml, xml_declaration=True, encoding='UTF-8')
                archive.writestr(item, content)
        renumber = UG.helper('renumber_footnotes_per_page')
        with patch.object(renumber, 'get_footnote_pages_via_word', return_value=[5, 6]), redirect_stdout(io.StringIO()):
            renumber.renumber_docx(output, output)
        UG.verify_final_circles(output, [5, 6])
        with self.assertRaisesRegex(ValueError, '圈码不一致'):
            UG.verify_final_circles(output, [5, 5])

    def test_bad_metadata_and_draft_pdf_do_not_publish(self):
        output = self.root / 'bad.docx'
        data = dict(META); data['keywords_en'] = ['only one']
        self.meta.write_text(json.dumps(data), encoding='utf-8')
        with self.assertRaisesRegex(ValueError, '关键词'):
            UG.build(self.md, self.meta, output, draft=True)
        self.assertFalse(output.exists())
        with self.assertRaisesRegex(ValueError, 'draft'):
            UG.build(self.md, self.meta, output, draft=True, pdf_path=self.root / 'bad.pdf')

if __name__ == '__main__':
    unittest.main()
