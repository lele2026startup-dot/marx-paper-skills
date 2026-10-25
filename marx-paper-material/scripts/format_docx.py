# -*- coding: utf-8 -*-
"""
format_docx.py — 资料库 TXT 批量渲染为规范排版 Word（马理论材料入库用）

用法：
  python format_docx.py <文件或文件夹> [<文件或文件夹> ...]
  - 传 .txt 文件 → 渲染同目录同名 .docx
  - 传文件夹 → 递归找里面所有 .txt 逐个渲染（资料库整目录批量）

排版规范（标题黑体居中、正文宋体小四、1.5 倍行距、首行缩进）：
  主标题（第一非空行）  黑体 16pt 三号  居中加粗
  日期行（含年月、短行）宋体 12pt 小四  居中
  * 编者注              楷体 10.5pt 五号  左缩进
  ## 小节标题           黑体 14pt 四号  加粗
  正文段落              宋体 12pt 小四  1.5 倍行距、首行缩进两字符、段后 6pt
  页边距                上下 2.54cm 左右 3.17cm

安全约定：只写 .docx，绝不改动 .txt；整篇乱码的文件跳过不动（打印警告）。
默认拒绝覆盖已有 docx；--overwrite 显式授权后备份旧文件。
"""
import os, re, sys, io, argparse, tempfile
from pathlib import Path
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn

DATE_LINE = re.compile(r'[一〇零九八七六五四三二○Ｏ0-9]{2,}.{0,40}(年|月)')


def garble_count(text):
    return text.count("锟斤拷") + text.count("\ufffd")


def set_run(run, font, size, bold=False):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = rPr.makeelement(qn('w:rFonts'), {})
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), font)
    rFonts.set(qn('w:ascii'), font if font != "宋体" else "Times New Roman")
    rFonts.set(qn('w:hAnsi'), font if font != "宋体" else "Times New Roman")


def add_para(doc, text, font, size, bold=False, align=None, indent_chars=0,
             space_after=6, line=None):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    if align is not None:
        pf.alignment = align
    pf.space_after = Pt(space_after)
    if line:
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        pf.line_spacing = line
    if indent_chars:
        # 首行缩进：按字符数设置
        pPr = p._element.get_or_add_pPr()
        ind = pPr.find(qn('w:ind'))
        if ind is None:
            ind = pPr.makeelement(qn('w:ind'), {})
            pPr.append(ind)
        ind.set(qn('w:firstLineChars'), str(indent_chars * 100))
        ind.set(qn('w:firstLine'), str(int(indent_chars * size * 20)))
    r = p.add_run(text)
    set_run(r, font, size, bold)
    return p



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
            os.rename(temp, target)
        else:
            os.link(temp, target)
            temp.unlink()


def render(txt_path, docx_path, overwrite=False):
    if Path(docx_path).exists() and not overwrite:
        raise FileExistsError(f"输出已存在，未覆盖：{docx_path}；另存或获授权后使用 --overwrite")
    text = Path(txt_path).read_text(encoding="utf-8")
    bad = garble_count(text)
    if bad > 20:
        print(f"[SKIP 乱码{bad}处] {os.path.basename(txt_path)}")
        return False

    doc = Document()
    # 页边距
    for sec in doc.sections:
        sec.top_margin = Cm(2.54)
        sec.bottom_margin = Cm(2.54)
        sec.left_margin = Cm(3.17)
        sec.right_margin = Cm(3.17)

    lines = [ln.rstrip() for ln in text.split("\n")]
    i, n = 0, len(lines)
    titled = False

    while i < n:
        s = lines[i].strip()
        if not s:
            i += 1
            continue
        if not titled:
            # 第一非空行 = 主标题
            add_para(doc, s, "黑体", 16, bold=True,
                     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)
            titled = True
            i += 1
            continue
        if s.startswith("## "):
            add_para(doc, s[3:].strip(), "黑体", 14, bold=True, space_after=6)
        elif s.startswith("* ") or s.startswith("*"):
            add_para(doc, s.lstrip("*").strip(), "楷体", 10.5, space_after=8)
        elif DATE_LINE.search(s) and len(s) < 60 and not s.startswith("　") and i <= 5:
            # 日期行（总在标题附近）：居中
            add_para(doc, s, "宋体", 12,
                     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)
        else:
            add_para(doc, s, "宋体", 12, indent_chars=2, line=1.5)
        i += 1

    target = Path(docx_path)
    target.parent.mkdir(parents=True, exist_ok=True)
    fd, temp = tempfile.mkstemp(prefix='.' + target.stem + '-', suffix='.docx', dir=target.parent)
    os.close(fd)
    try:
        doc.save(temp)
        Document(temp)  # Reopen before publication.
        publish_file(temp, target, overwrite)
    finally:
        Path(temp).unlink(missing_ok=True)
    warn = f"（注意：{bad} 处疑似乱码字符）" if bad else ""
    print(f"[OK] {os.path.basename(docx_path)}  {len(text)}字符{warn}")
    return True


def collect_txts(paths):
    txts = []
    for p in paths:
        if os.path.isfile(p) and p.lower().endswith(".txt"):
            txts.append(p)
        elif os.path.isdir(p):
            for root, _, files in os.walk(p):
                for f in files:
                    if f.lower().endswith(".txt"):
                        txts.append(os.path.join(root, f))
    return txts


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("paths", nargs="+")
    ap.add_argument("--overwrite", action="store_true", help="已获授权覆盖；先保存原文件备份")
    args = ap.parse_args()

    txts = collect_txts(args.paths)
    if not txts:
        print("没找到 .txt 文件")
        sys.exit(1)
    ok, skip = 0, 0
    for t in txts:
        d = os.path.splitext(t)[0] + ".docx"
        if render(t, d, overwrite=args.overwrite):
            ok += 1
        else:
            skip += 1
    print(f"\n完成：{ok} 份已排版，{skip} 份跳过")


if __name__ == "__main__":
    main()
