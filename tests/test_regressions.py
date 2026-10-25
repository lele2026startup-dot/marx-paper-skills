"""Behavioral regressions for encoding, citation preservation, output safety and installation."""
from pathlib import Path
from contextlib import redirect_stdout
import importlib.util
import io
import json
import os
import shutil
import tempfile
import unittest
from unittest.mock import patch
from email.message import Message
import zipfile
from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
TEMP_ROOT = ROOT / '_test-output'
TEMP_ROOT.mkdir(exist_ok=True)

def module(rel, name):
    spec = importlib.util.spec_from_file_location(name, ROOT / rel)
    result = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(result)
    return result

FETCH = module('marx-paper-material/scripts/fetch_text.py', 'fetch')
FORMAT = module('marx-paper-material/scripts/format_docx.py', 'material_format')
JOURNAL = module('marx-paper-journal-format/scripts/build_journal_docx.py', 'journal')
SYNC = module('scripts/sync_skills.py', 'sync')
THESIS = module('marx-paper-thesis-format/scripts/build_docx.py', 'thesis')
CIRCLES = module('marx-paper-thesis-format/scripts/footnote_circle_marks.py', 'circles')
RENUMBER = module('marx-paper-thesis-format/scripts/renumber_footnotes_per_page.py', 'renumber')
NS = {'w': JOURNAL.W_NS}

class Encodings(unittest.TestCase):
    def test_utf8_without_declaration(self):
        original = '<html><p>中文</p></html>'
        text, encoding, bad = FETCH.pick_encoding(original.encode('utf-8'), '')
        self.assertEqual(text, original)
        self.assertEqual(encoding, 'utf-8')
        self.assertEqual(bad, 0)

    def test_legacy_html_meta(self):
        original = '<html><meta charset="gb2312"><p>中文——测试</p></html>'
        self.assertEqual(FETCH.pick_encoding(original.encode('gb18030'), '')[0], original)

    def test_bom_and_forced_encoding(self):
        self.assertEqual(FETCH.pick_encoding('中文'.encode('utf-8-sig'))[0], '中文')
        with self.assertRaises(UnicodeDecodeError):
            FETCH.pick_encoding('中文'.encode('utf-8'), forced='ascii')

    def test_fetch_protects_existing_output_before_network(self):
        with tempfile.TemporaryDirectory(dir=TEMP_ROOT) as directory:
            out = Path(directory) / 'source.txt'
            out.write_text('original', encoding='utf-8')
            with self.assertRaises(FileExistsError):
                FETCH.fetch_one('test', 'https://invalid.invalid', str(out))
            self.assertEqual(out.read_text(), 'original')

    def test_fetch_writes_utf8_and_backs_up_authorized_overwrite(self):
        class Response:
            headers = Message()
            def __enter__(self): return self
            def __exit__(self, *args): pass
            def read(self): return '<html><body><h1>测试标题</h1><p>中文正文。</p></body></html>'.encode('utf-8')
            def geturl(self): return 'https://example.test/article'
        with tempfile.TemporaryDirectory(dir=TEMP_ROOT) as directory:
            target = Path(directory) / 'text.txt'
            target.write_text('original', encoding='utf-8')
            with patch.object(FETCH, 'urlopen', return_value=Response()):
                result = FETCH.fetch_one('test', 'https://example.test/article', str(target), overwrite=True)
            self.assertIn('中文正文', target.read_text(encoding='utf-8'))
            self.assertEqual(result['encoding'], 'utf-8')
            self.assertEqual(next(target.parent.glob('text.txt.bak.*')).read_text(), 'original')

class Journal(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(dir=TEMP_ROOT)
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.meta = self.root / 'meta.json'
        self.meta.write_text(json.dumps({'author': '测试作者', 'affiliation': '测试单位', 'author_bio': '虚构测试'}, ensure_ascii=False), encoding='utf-8')

    def build(self, body, refs, heading='## 参考文献', title='# 测试论文', **kwargs):
        md = self.root / 'input.md'
        md.write_text(title + '\n\n**摘要：** 测试摘要\n\n**关键词：** 测试\n\n## 一、正文\n\n' + body + '\n\n' + heading + '\n' + refs + '\n', encoding='utf-8')
        out = self.root / 'out.docx'
        with redirect_stdout(io.StringIO()):
            JOURNAL.build(str(md), str(self.meta), str(out), **kwargs)
        with zipfile.ZipFile(out) as archive:
            doc = etree.fromstring(archive.read('word/document.xml'))
            notes = etree.fromstring(archive.read('word/endnotes.xml'))
        return out, doc, notes

    def test_repeated_reference_keeps_all_occurrences(self):
        _, doc, notes = self.build('第一处[1]。\n\n第二处[1]。\n\n第三处[1]。', '[1] 作者.虚构题名[M].某地:出版社,2020.')
        self.assertEqual(len(doc.findall('.//w:endnoteReference', NS)), 1)
        fields = doc.findall('.//w:fldSimple', NS)
        self.assertEqual(len(fields), 2)
        self.assertTrue(all('NOTEREF cite_1' in f.get(JOURNAL.qn('w:instr')) for f in fields))
        self.assertEqual(len([n for n in notes if int(n.get(JOURNAL.qn('w:id'))) >= 2]), 1)
        self.assertIn('第二处', ''.join(doc.itertext()))

    def test_noncontinuous_and_out_of_order_numbers(self):
        _, doc, notes = self.build('先[7]后[3]再[7]。', '[3] 甲.文献甲[M].2020.\n[7] 乙.文献乙[M].2021.')
        self.assertEqual([r.get(JOURNAL.qn('w:id')) for r in doc.findall('.//w:endnoteReference', NS)], ['2', '3'])
        content = [n for n in notes if int(n.get(JOURNAL.qn('w:id'))) >= 2]
        self.assertIn('乙.文献乙', ''.join(content[0].itertext()))
        self.assertIn('甲.文献甲', ''.join(content[1].itertext()))

    def test_missing_source_does_not_publish(self):
        with self.assertRaisesRegex(ValueError, '没有对应'):
            self.build('引文[9]。', '[1] 作者.虚构书名[M].2020.')
        self.assertFalse((self.root / 'out.docx').exists())
        self.assertFalse(list(self.root.glob('.*.docx')))

    def test_footnotes_are_converted_with_all_positions(self):
        _, doc, notes = self.build('第一处[^source]。\n\n第二处[^source]。\n\n[^source]: 作者.虚构书名[M].2020.', '', heading='')
        self.assertEqual(len(doc.findall('.//w:endnoteReference', NS)), 1)
        self.assertEqual(len(doc.findall('.//w:fldSimple', NS)), 1)
        self.assertNotIn('[^', ''.join(doc.itertext()))

    def test_bibliography_heading_variants(self):
        for index, heading in enumerate(('## 参考文献', '**参考文献：**', '参考文献:')):
            with self.subTest(heading=heading):
                output = self.root / 'out.docx'
                if output.exists():
                    output.unlink()
                self.build('引文[1]。', '[1] 测试文献.', heading=heading)

    def test_missing_footnote_is_rejected(self):
        with self.assertRaisesRegex(ValueError, '没有定义'):
            self.build('引文[^absent]。', '[1] 测试文献.')
        self.assertFalse((self.root / 'out.docx').exists())

    def test_mixed_missing_number_cannot_alias_converted_note(self):
        with self.assertRaisesRegex(ValueError, '没有对应'):
            self.build('未知[1]，脚注[^source]。\n\n[^source]: 测试文献.', '', heading='')
        self.assertFalse((self.root / 'out.docx').exists())

    def test_grouped_citation_is_rejected(self):
        with self.assertRaisesRegex(ValueError, '暂不支持'):
            self.build('引文[1,2]。', '[1] 甲.\n[2] 乙.')

    def test_existing_output_survives_invalid_rebuild(self):
        out, _, _ = self.build('引文[1]。', '[1] 测试文献.')
        before = out.read_bytes()
        with self.assertRaises(FileExistsError):
            self.build('引文[1]。', '[1] 测试文献.')
        with self.assertRaises(ValueError):
            self.build('引文[9]。', '[1] 测试文献.', overwrite=True)
        self.assertEqual(out.read_bytes(), before)

    def test_authorized_overwrite_retains_backup(self):
        out, _, _ = self.build('引文[1]。', '[1] 测试文献.')
        before = out.read_bytes()
        self.build('新引文[1]。', '[1] 另一文献.', overwrite=True)
        backups = list(self.root.glob('out.docx.bak.*'))
        self.assertEqual(len(backups), 1)
        self.assertEqual(backups[0].read_bytes(), before)

    def test_no_english_flag(self):
        meta = json.loads(self.meta.read_text(encoding='utf-8'))
        meta['english'] = {'title': 'UNIQUE_ENGLISH_TITLE', 'author': 'Test', 'affiliation': 'Test', 'abstract': 'Test', 'keywords': 'Test'}
        self.meta.write_text(json.dumps(meta), encoding='utf-8')
        _, doc, _ = self.build('引文[1]。', '[1] 测试文献.', no_english=True)
        self.assertNotIn('UNIQUE_ENGLISH_TITLE', ''.join(doc.itertext()))

    def test_yaml_title(self):
        self.build('引文[1]。', '[1] 测试文献.', title='---\ntitle: 测试论文\n---')

    def test_duplicate_source_entries_keep_positions(self):
        _, doc, _ = self.build('引文[1]和[8]。', '[1] 同一文献.\n[8] 同一文献.')
        self.assertEqual(len(doc.findall('.//w:endnoteReference', NS)), 1)
        self.assertEqual(len(doc.findall('.//w:fldSimple', NS)), 1)

    def test_reference_heading_belongs_to_a_referenced_note(self):
        _, doc, notes = self.build('引文[1]。', '[1] 测试文献.')
        identifier = doc.find('.//w:endnoteReference', NS).get(JOURNAL.qn('w:id'))
        note = notes.find(f'w:endnote[@w:id="{identifier}"]', NS)
        self.assertEqual(''.join(note.find('w:p', NS).itertext()), '参考文献')

class OutputAndInstall(unittest.TestCase):
    def test_footnote_scratch_directory_is_not_deleted(self):
        with tempfile.TemporaryDirectory(dir=TEMP_ROOT) as directory:
            root = Path(directory)
            scratch = root / '_fn_tmp'
            scratch.mkdir()
            sentinel = scratch / 'user-file.txt'
            sentinel.write_text('retain', encoding='utf-8')
            invalid = root / 'invalid.docx'
            invalid.write_bytes(b'not-a-zip')
            previous = Path.cwd()
            try:
                os.chdir(root)
                with self.assertRaises(zipfile.BadZipFile):
                    CIRCLES.process(invalid, root / 'output.docx')
            finally:
                os.chdir(previous)
            self.assertEqual(sentinel.read_text(), 'retain')
            self.assertFalse(list(root.glob('.footnote-*')))

    def test_overflow_circles_are_readable(self):
        for helper in (CIRCLES, RENUMBER):
            self.assertEqual(helper.get_circle(20), '⑳')
            self.assertEqual(helper.get_circle(21), '(21)')

    def test_material_output_is_protected_and_backed_up(self):
        with tempfile.TemporaryDirectory(dir=TEMP_ROOT) as directory:
            root = Path(directory)
            text = root / 'material.txt'
            out = root / 'material.docx'
            text.write_text('虚构标题\n正文内容。', encoding='utf-8')
            with redirect_stdout(io.StringIO()):
                FORMAT.render(text, out)
            before = out.read_bytes()
            with self.assertRaises(FileExistsError):
                FORMAT.render(text, out)
            with redirect_stdout(io.StringIO()):
                FORMAT.render(text, out, overwrite=True)
            self.assertEqual(next(root.glob('material.docx.bak.*')).read_bytes(), before)

    def test_thesis_refuses_existing_result_before_conversion(self):
        with tempfile.TemporaryDirectory(dir=TEMP_ROOT) as directory:
            root = Path(directory)
            md, out = root / 'input.md', root / 'out.docx'
            md.write_text('# 标题', encoding='utf-8')
            out.write_bytes(b'original')
            with self.assertRaises(FileExistsError):
                THESIS.build(md, out)
            self.assertEqual(out.read_bytes(), b'original')

    def test_install_restore_and_detect_drift(self):
        with tempfile.TemporaryDirectory(dir=TEMP_ROOT) as directory:
            root = Path(directory)
            source = root / 'source'
            source.mkdir()
            shutil.copy2(ROOT / 'skills-manifest.json', source / 'skills-manifest.json')
            for skill in json.loads((ROOT / 'skills-manifest.json').read_text(encoding='utf-8'))['skills']:
                shutil.copytree(ROOT / skill['name'], source / skill['name'], ignore=shutil.ignore_patterns('__pycache__', '*.pyc'))
            target, backup = root / 'skills', root / 'backups'
            result = SYNC.sync(source, target, apply=True, backup_dir=backup)
            self.assertTrue(result['consistent'])
            self.assertTrue(SYNC.sync(source, target)['consistent'])
            topic = target / 'marx-paper-topic/SKILL.md'
            topic.write_text('truncated local edit', encoding='utf-8')
            self.assertFalse(SYNC.sync(source, target)['consistent'])
            with self.assertRaisesRegex(ValueError, '本地改动'):
                SYNC.sync(source, target, apply=True, backup_dir=backup)
            result = SYNC.sync(source, target, apply=True, backup_dir=backup, overwrite_installed=True)
            saved = Path(result['backup']) / 'marx-paper-topic/SKILL.md'
            self.assertEqual(saved.read_text(), 'truncated local edit')
            self.assertEqual(topic.read_bytes(), (ROOT / 'marx-paper-topic/SKILL.md').read_bytes())
            (source / 'marx-paper-topic/SKILL.md').write_text('truncated source', encoding='utf-8')
            with self.assertRaisesRegex(ValueError, '不完整'):
                SYNC.sync(source, target, apply=True, backup_dir=backup)

if __name__ == '__main__':
    unittest.main()
