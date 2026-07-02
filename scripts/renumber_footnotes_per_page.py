# -*- coding: utf-8 -*-
"""
renumber_footnotes_per_page.py — 让 customMarkFollows 圈码脚注每页重新编号。

前置条件：docx 已经被 footnote_circle_marks.py 处理过（圈码已写入）。
本脚本做的事：
  1. 用 LibreOffice headless 把 docx 转 PDF
  2. 用 pdfplumber 找出每个圈码出现在哪一页
  3. 按页分组，每页内按出现顺序重新赋 ①②③
  4. 改写 docx 的 document.xml 和 footnotes.xml，替换圈码字符

依赖：LibreOffice (soffice)、pdfplumber
"""
import os
import re
import sys
import shutil
import zipfile
import subprocess
from pathlib import Path

# 圈码字符 ①-⑳ (0x2460-0x2473), ㉑-㉟ (0x3251-0x3255)
CIRCLE_NUMS = [chr(0x2460 + i) for i in range(20)]  # ①-⑳
CIRCLE_NUMS_EXT = [chr(0x3251 + i) for i in range(15)]  # ㉑-㉟
ALL_CIRCLES = CIRCLE_NUMS + CIRCLE_NUMS_EXT

SOFFICE = r"C:\Program Files\LibreOffice\program\soffice.exe"


def get_circle(n):
    if 1 <= n <= 20:
        return CIRCLE_NUMS[n - 1]
    if 21 <= n <= 35:
        return CIRCLE_NUMS_EXT[n - 21]
    return str(n)


def docx_to_pdf(docx_path, outdir):
    """LibreOffice headless 转 PDF。"""
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    cmd = [
        SOFFICE, "--headless", "--convert-to", "pdf",
        "--outdir", str(outdir),
        "-env:UserInstallation=file:///C:/tmp/lo_profile",
        str(docx_path)
    ]
    subprocess.run(cmd, capture_output=True, timeout=120, check=True)
    pdf_name = Path(docx_path).stem + ".pdf"
    return outdir / pdf_name


def find_circle_pages(pdf_path):
    """扫描 PDF，返回 {圈码字符: 页码(从1起)}。
    一个圈码在一页出现就记录该页；脚注引用和底部内容同页，取首次出现页即可。
    """
    import pdfplumber
    circle_to_page = {}
    with pdfplumber.open(str(pdf_path)) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            for c in ALL_CIRCLES:
                if c in text and c not in circle_to_page:
                    circle_to_page[c] = i + 1  # 页码从 1 起
    return circle_to_page


def build_renumber_map(circle_to_page, ordered_circles):
    """构建 {旧圈码: 新圈码} 映射。
    ordered_circles: 按文档出现顺序的旧圈码列表（来自 document.xml 扫描顺序）。
    按页分组后，每页内按出现顺序赋 ①②③。
    """
    # 按页分组
    pages = {}
    for c in ordered_circles:
        p = circle_to_page.get(c)
        if p is None:
            # 找不到页码（不该发生），归到最后一页之后
            p = max(pages.keys(), default=0) + 1
        pages.setdefault(p, []).append(c)

    # 每页内按出现顺序赋新圈码
    remap = {}
    for p in sorted(pages.keys()):
        circles_in_page = pages[p]
        for idx, old_c in enumerate(circles_in_page, start=1):
            remap[old_c] = get_circle(idx)
    return remap


def scan_ordered_circles(doc_xml):
    """按 document.xml 中出现顺序，提取 customMarkFollows run 里的圈码字符。"""
    # 匹配 customMarkFollows 后面紧跟的 <w:t>圈码</w:t>
    pattern = re.compile(
        r'<w:footnoteReference[^>]*w:customMarkFollows="true"[^>]*/>'
        r'\s*<w:t[^>]*>([^<]+)</w:t>',
        re.DOTALL
    )
    return [m.group(1) for m in pattern.finditer(doc_xml)]


def renumber_docx(input_docx, output_docx):
    # 1. 转 PDF
    print("步骤1: LibreOffice 转 PDF...")
    pdf_path = docx_to_pdf(input_docx, "_renumber_tmp_pdf")
    print(f"  PDF: {pdf_path}")

    # 2. 找圈码页码
    print("步骤2: pdfplumber 扫描圈码页码...")
    circle_to_page = find_circle_pages(pdf_path)
    print(f"  找到 {len(circle_to_page)} 个圈码的页码")

    # 3. 解压 docx
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

    # 4. 按出现顺序扫描旧圈码
    ordered = scan_ordered_circles(doc_xml)
    print(f"步骤3: document.xml 中找到 {len(ordered)} 个圈码引用")

    # 5. 构建重编号映射
    remap = build_renumber_map(circle_to_page, ordered)
    print(f"步骤4: 重编号映射 ({len(remap)} 项)")

    # 6. 替换 document.xml 里 customMarkFollows run 的圈码
    def repl_doc(m):
        ref_tag = m.group(1)
        old_circle = m.group(2)
        new_circle = remap.get(old_circle, old_circle)
        # 找到这个引用对应的 footnote id，后面也要同步改底部
        return f'{ref_tag}<w:t xml:space="preserve">{new_circle}</w:t>'

    doc_pattern = re.compile(
        r'(<w:footnoteReference[^>]*w:customMarkFollows="true"[^>]*/>)\s*'
        r'<w:t[^>]*>([^<]+)</w:t>',
        re.DOTALL
    )
    new_doc_xml = doc_pattern.sub(repl_doc, doc_xml)
    doc_path.write_text(new_doc_xml, encoding="utf-8")

    # 7. 替换 footnotes.xml 里的圈码
    # 底部脚注的圈码是 <w:r><w:t>①</w:t></w:r>（footnote_circle_marks 写入的，无 rStyle）
    def repl_fn(m):
        old_circle = m.group(1)
        new_circle = remap.get(old_circle, old_circle)
        return f'<w:r><w:t xml:space="preserve">{new_circle}</w:t></w:r>'

    fn_pattern = re.compile(
        r'<w:r><w:t xml:space="preserve">([^<①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯]+)</w:t></w:r>'
    )
    # 上面的排除写法不对，直接匹配所有圈码 run
    fn_circle_pattern = re.compile(
        r'<w:r><w:t xml:space="preserve">([' + re.escape("".join(ALL_CIRCLES)) + r'])</w:t></w:r>'
    )
    new_fn_xml = fn_circle_pattern.sub(repl_fn, fn_xml)
    fn_path.write_text(new_fn_xml, encoding="utf-8")

    # 8. 重新打包
    with zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as z:
        for f in tmp.rglob("*"):
            if f.is_file():
                z.write(f, f.relative_to(tmp).as_posix())
    shutil.rmtree(tmp)
    shutil.rmtree("_renumber_tmp_pdf")
    print(f"\n完成: {output_docx}")


if __name__ == "__main__":
    renumber_docx(sys.argv[1], sys.argv[2])
