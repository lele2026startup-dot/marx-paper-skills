# -*- coding: utf-8 -*-
"""
footnote_circle_marks.py — 原型：把脚注引用和底部标记改成圈码 ①②③。

方向 B 验证：用 customMarkFollows 让脚注引用显示自定义字符（圈码），
同时把底部 footnoteRef 也替换成圈码字符。w:id 关联保留，仍是真脚注。

第一版：全文连续编号（① 到 ⑯），先验证 customMarkFollows 思路。
每页重编号在思路验证成功后再加（需要分页判定）。
"""
import sys
import zipfile
import shutil
import re
from pathlib import Path

# 圈码字符 ①②③...㉑㉒...（Unicode 0x2460 起，到 0x2473 是 ①-⑳，0x3251 起是 ㉑-㉟）
CIRCLE_NUMS = [chr(0x2460 + i) for i in range(20)]  # ①-⑳


def get_circle(n):
    """1->①, 2->②, ..., >20 用 (21) 等括号数字"""
    if 1 <= n <= 20:
        return CIRCLE_NUMS[n - 1]
    # 超过 20 用带圈括号数字 ㉑㉒㉓... (0x3251 起)
    if 21 <= n <= 35:
        return chr(0x3251 + (n - 21))
    return str(n)  # 兜底


def process(input_docx, output_docx):
    # 解压
    tmp = Path('_fn_tmp')
    if tmp.exists():
        shutil.rmtree(tmp)
    tmp.mkdir()
    with zipfile.ZipFile(input_docx) as z:
        z.extractall(tmp)

    # 1. 读 document.xml，找所有 footnoteReference，按出现顺序分配圈码
    doc_path = tmp / 'word' / 'document.xml'
    doc_xml = doc_path.read_text(encoding='utf-8')

    # 按出现顺序找 footnoteReference，记录 id -> 圈码 映射
    id_to_circle = {}
    counter = 0
    # 匹配 <w:r>...<w:footnoteReference w:id="N"/></w:r>
    ref_pattern = re.compile(
        r'(<w:r\b[^>]*>)(<w:rPr>.*?</w:rPr>)?(<w:footnoteReference\s+w:id="(\d+)"\s*/>)\s*(</w:r>)',
        re.DOTALL
    )

    def replace_ref(m):
        nonlocal counter
        counter += 1
        r_open, rpr, ref_tag, fid, r_close = m.groups()
        circle = get_circle(counter)
        id_to_circle[fid] = circle
        # 加 customMarkFollows="true" 到 footnoteReference，并在 run 里加圈码文本
        new_ref = ref_tag.replace('/>', ' w:customMarkFollows="true"/>')
        # 在 rPr 里确保有 FootnoteReference 样式（superscript 上标）
        # 用 <w:t>圈码</w:t> 作为自定义标记
        return f'{r_open}{rpr}{new_ref}<w:t xml:space="preserve">{circle}</w:t>{r_close}'

    new_doc_xml = ref_pattern.sub(replace_ref, doc_xml)
    doc_path.write_text(new_doc_xml, encoding='utf-8')
    print(f'正文：替换了 {counter} 个脚注引用为圈码')

    # 2. 读 footnotes.xml，把每个 footnote 里的 <w:footnoteRef/> 替换成圈码
    fn_path = tmp / 'word' / 'footnotes.xml'
    fn_xml = fn_path.read_text(encoding='utf-8')

    # 找每个 <w:footnote w:id="N">...<w:footnoteRef/>...</w:footnote>
    footnote_pattern = re.compile(
        r'(<w:footnote\s+w:id="(\d+)"[^>]*>)(.*?)(</w:footnote>)',
        re.DOTALL
    )

    def replace_fn(m):
        open_tag, fid, inner, close_tag = m.groups()
        if fid not in id_to_circle:
            return m.group(0)  # 分隔符等保留
        circle = id_to_circle[fid]
        new_run = (
            f'<w:r><w:t xml:space="preserve">{circle}</w:t></w:r>'
        )
        # 原始结构跨行带缩进：<w:r>\r\n  <w:rPr>...</w:rPr>\r\n  <w:footnoteRef />\r\n</w:r>
        # 用宽松正则匹配含 footnoteRef 的整个 run
        run_pattern = re.compile(
            r'<w:r\b[^>]*>\s*(?:<w:rPr>.*?</w:rPr>)?\s*'
            r'<w:footnoteRef\s*/>\s*</w:r>',
            re.DOTALL
        )
        new_inner, n = run_pattern.subn(new_run, inner, count=1)
        if n == 0:
            # 兜底：单独替换 footnoteRef
            new_inner = re.sub(r'<w:footnoteRef\s*/>', new_run, inner, count=1)
        return f'{open_tag}{new_inner}{close_tag}'

    new_fn_xml = footnote_pattern.sub(replace_fn, fn_xml)
    fn_path.write_text(new_fn_xml, encoding='utf-8')
    print(f'底部：替换了 {len(id_to_circle)} 个 footnoteRef 为圈码')

    # 3. 重新打包
    with zipfile.ZipFile(output_docx, 'w', zipfile.ZIP_DEFLATED) as z:
        for f in tmp.rglob('*'):
            if f.is_file():
                z.write(f, f.relative_to(tmp).as_posix())
    shutil.rmtree(tmp)
    print(f'已生成: {output_docx}')


if __name__ == '__main__':
    process(sys.argv[1], sys.argv[2])
