# -*- coding: utf-8 -*-
"""
renumber_footnotes_per_page.py — 让 customMarkFollows 圈码脚注每页重新编号。

用 Word/WPS 的 COM 接口拿真实页码（而不是 LibreOffice，因为两者分页不同）。

前置条件：docx 已经被 footnote_circle_marks.py 处理过（圈码已写入，连续编号）。
本脚本做的事：
  1. 用 Word COM 打开 docx，查每个脚注引用的真实页码
  2. 按页分组，每页内按出现顺序重新赋 ①②③
  3. 改写 docx 的 document.xml 和 footnotes.xml，替换圈码字符

关键设计：用 footnote 的 w:id 作重编号映射的 key（不用旧圈码字符）。
因为同一文献被重复引用时，footnote_circle_marks 会把多个引用点写成同一个圈码字符，
如果用圈码字符作 key，remap 会被反复覆盖导致错位。w:id 是每个引用点唯一的，不会冲突。

超过 20 个的圈码：Unicode ①-⑳ 只到 20，㉑-㉟ 有字符但常见字体（宋体等）多无字形，
所以超过 20 自动降级成 (21)(22)... 形式，并在结束时提示用户该页脚注过多。

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
    """1-20 用圈码，21-35 用扩展圈码（字体可能不支持，但先给），超过则降级成 (n)"""
    if 1 <= n <= 20:
        return CIRCLE_NUMS[n - 1]
    if 21 <= n <= 35:
        return CIRCLE_NUMS_EXT[n - 21]
    return f'({n})'


def get_footnote_pages_via_word(docx_path):
    """用 Word COM 打开文档，返回 [页码] 列表，按脚注顺序（1起）。"""
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

    # 3. 按 document.xml 出现顺序，提取每个 footnoteReference 的 w:id 和它后面的圈码
    # 真实结构：<w:footnoteReference w:id="X" w:customMarkFollows="true"/><w:t ...>圈码</w:t>
    pat = re.compile(
        r'<w:footnoteReference\s+w:id="(\d+)"\s+w:customMarkFollows="true"\s*/>'
        r'\s*(<w:r[^>]*>)?\s*<w:t[^>]*>([^<]+)</w:t>'
    )
    matches = pat.findall(doc_xml)
    if not matches:
        # 兜底：放宽匹配
        refs_pat = re.compile(
            r'<w:footnoteReference\s+w:id="(\d+)"[^>]*/>'
            r'[\s\S]{0,200}?<w:t[^>]*>([^<]+)</w:t>'
        )
        matches = [(m.group(1), '', m.group(2)) for m in refs_pat.finditer(doc_xml)]
    print(f"步骤2: 匹配 {len(matches)} 个 footnote-圈码对")

    id_list = [m[0] for m in matches]
    if len(id_list) != len(pages):
        print(f"  警告: 数量不一致 {len(id_list)} vs {len(pages)}")

    # 4. 按页分组重编号 —— 用 w:id 作 key（每个引用点唯一，不会冲突）
    id_to_circle = {}
    page_counter = {}
    page_groups = {}
    overflow_pages = []
    for idx, fid in enumerate(id_list):
        p = pages[idx] if idx < len(pages) else 0
        page_counter[p] = page_counter.get(p, 0) + 1
        new_c = get_circle(page_counter[p])
        id_to_circle[fid] = new_c
        page_groups.setdefault(p, []).append(new_c)
        if page_counter[p] > 20:
            overflow_pages.append(p)

    print("步骤3: 每页圈码分配:")
    for p in sorted(page_groups):
        print(f"  页 {p}: {page_groups[p]}")

    if overflow_pages:
        print(f"\n⚠ 提示: 以下页脚注超过 20 个，已降级为 (n) 形式: {sorted(set(overflow_pages))}")
        print("  建议：该页脚注过多，可考虑合并部分引用或拆分段落。")

    # 5. 替换 document.xml 里每个 footnoteReference 后的圈码
    def repl_doc(m):
        full = m.group(0)
        fid = m.group(1)
        old_c = m.group(3)
        new_c = id_to_circle.get(fid, old_c)
        return full.replace(f'>{old_c}<', f'>{new_c}<', 1)
    new_doc_xml = pat.sub(repl_doc, doc_xml)
    doc_path.write_text(new_doc_xml, encoding="utf-8")

    # 6. 替换 footnotes.xml：按 <w:footnote w:id="X"> 分块，每块换第一个圈码
    parts = re.split(r'(<w:footnote\s+w:id="\d+"[^>]*>)', fn_xml)
    for i in range(1, len(parts), 2):
        head = parts[i]
        m_id = re.search(r'w:id="(\d+)"', head)
        if not m_id:
            continue
        fid = m_id.group(1)
        if i + 1 < len(parts):
            content = parts[i + 1]
            replaced = False
            for old_c in ALL_CIRCLES:
                if old_c in content:
                    new_c = id_to_circle.get(fid, old_c)
                    content = content.replace(old_c, new_c, 1)
                    replaced = True
                    break
            if not replaced:
                # 数字 fallback（如 36/37/38）
                m = re.search(r'<w:t[^>]*>(\d+)</w:t>', content)
                if m:
                    new_c = id_to_circle.get(fid, m.group(1))
                    content = content[:m.start()] + re.sub(
                        r'(<w:t[^>]*>)[^<]*(</w:t>)',
                        lambda mm: f'{mm.group(1)}{new_c}{mm.group(2)}',
                        content[m.start():], count=1
                    )
            parts[i + 1] = content
    new_fn_xml = ''.join(parts)
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
