# -*- coding: utf-8 -*-
"""
renumber_footnotes_per_page.py — 让 customMarkFollows 圈码脚注每页重新编号。

用 Word/WPS 的 COM 接口拿真实页码（而不是 LibreOffice，因为两者分页不同）。

前置条件：docx 已经被 footnote_circle_marks.py 处理过（圈码已写入，连续编号）。
本脚本做的事：
  1. 用 Word COM 打开 docx，查每个脚注引用的真实页码
  2. 按页分组，每页内按出现顺序重新赋 ①②③
  3. 改写 docx 的 document.xml 和 footnotes.xml，替换圈码字符

依赖：pywin32 (Word 或 WPS 任一，COM 接口都行)
"""
import os
import re
import sys
import shutil
import zipfile
from pathlib import Path

# 圈码字符 ①-⑳ (0x2460-0x2473), ㉑-㉟ (0x3251-0x3255)
CIRCLE_NUMS = [chr(0x2460 + i) for i in range(20)]
CIRCLE_NUMS_EXT = [chr(0x3251 + i) for i in range(15)]
ALL_CIRCLES = CIRCLE_NUMS + CIRCLE_NUMS_EXT


def get_circle(n):
    if 1 <= n <= 20:
        return CIRCLE_NUMS[n - 1]
    if 21 <= n <= 35:
        return CIRCLE_NUMS_EXT[n - 21]
    return str(n)


def get_footnote_pages_via_word(docx_path):
    """用 Word COM 打开文档，返回 [页码] 列表，按脚注顺序（1起）。
    页码也是 Word 报告的真实页码。
    """
    import win32com.client
    abs_path = os.path.abspath(str(docx_path))
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        doc = word.Documents.Open(abs_path)
        pages = []
        for i in range(1, doc.Footnotes.Count + 1):
            fn = doc.Footnotes(i)
            # wdActiveEndPageNumber = 3
            page = fn.Reference.Information(3)
            pages.append(page)
        doc.Close(False)
        return pages
    finally:
        word.Quit()


def scan_ordered_circles(doc_xml):
    """按 document.xml 出现顺序，提取 customMarkFollows run 里的圈码字符。"""
    pattern = re.compile(
        r'<w:footnoteReference[^>]*w:customMarkFollows="true"[^>]*/>'
        r'\s*<w:t[^>]*>([^<]+)</w:t>',
        re.DOTALL
    )
    return [m.group(1) for m in pattern.finditer(doc_xml)]


def renumber_docx(input_docx, output_docx):
    # 1. 用 Word COM 查真实页码
    print("步骤1: 用 Word COM 查询脚注真实页码...")
    pages = get_footnote_pages_via_word(input_docx)
    print(f"  共 {len(pages)} 个脚注，页码分布: {pages}")

    # 2. 解压 docx
    tmp = Path("_renumber_tmp")
    if tmp.exists():
        shutil.rmtree(tmp)
    tmp.mkdir()
    with zipfile.ZipFile(input_docx) as z:
        z.extractall(tmp)

    doc_path = tmp / "word" / "document.xml"
    fn_path = tmp / "word" / "footnotes.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")
    fn_xml = fn_path.read_text(encoding="utf-8")

    # 3. 按出现顺序扫描旧圈码
    ordered = scan_ordered_circles(doc_xml)
    print(f"步骤2: document.xml 中找到 {len(ordered)} 个圈码引用")

    if len(ordered) != len(pages):
        print(f"警告: 圈码数({len(ordered)}) 与 Word 脚注数({len(pages)}) 不一致!")

    # 4. 构建重编号映射：按页分组，每页内按顺序赋 ①②③
    remap = {}
    pages_dict = {}
    for idx, old_circle in enumerate(ordered):
        p = pages[idx] if idx < len(pages) else max(pages_dict.keys(), default=0) + 1
        pages_dict.setdefault(p, []).append(old_circle)
    for p in sorted(pages_dict.keys()):
        for n, old_c in enumerate(pages_dict[p], start=1):
            remap[old_c] = get_circle(n)
    print(f"步骤3: 重编号映射构建完成 ({len(remap)} 项)")

    # 5. 替换 document.xml 里 customMarkFollows run 的圈码
    def repl_doc(m):
        ref_tag = m.group(1)
        old_circle = m.group(2)
        new_circle = remap.get(old_circle, old_circle)
        return f'{ref_tag}<w:t xml:space="preserve">{new_circle}</w:t>'

    doc_pattern = re.compile(
        r'(<w:footnoteReference[^>]*w:customMarkFollows="true"[^>]*/>)\s*'
        r'<w:t[^>]*>([^<]+)</w:t>',
        re.DOTALL
    )
    new_doc_xml = doc_pattern.sub(repl_doc, doc_xml)
    doc_path.write_text(new_doc_xml, encoding="utf-8")

    # 6. 替换 footnotes.xml 里的圈码
    def repl_fn(m):
        old_circle = m.group(1)
        new_circle = remap.get(old_circle, old_circle)
        return f'<w:r><w:t xml:space="preserve">{new_circle}</w:t></w:r>'

    fn_circle_pattern = re.compile(
        r'<w:r><w:t xml:space="preserve">([' + re.escape("".join(ALL_CIRCLES)) + r'])</w:t></w:r>'
    )
    new_fn_xml = fn_circle_pattern.sub(repl_fn, fn_xml)
    fn_path.write_text(new_fn_xml, encoding="utf-8")

    # 7. 重新打包
    with zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as z:
        for f in tmp.rglob("*"):
            if f.is_file():
                z.write(f, f.relative_to(tmp).as_posix())
    shutil.rmtree(tmp)
    print(f"\n完成: {output_docx}")


if __name__ == "__main__":
    renumber_docx(sys.argv[1], sys.argv[2])
