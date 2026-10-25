# -*- coding: utf-8 -*-
"""
build_docx.py — 一键 pipeline：markdown → 规范 Word 文档。

把 format skill 的四步串成一条命令，避免漏跑或跳步：
  1. pandoc 把 markdown 转成 docx（脚注用 pandoc 的 [^n] 语法）
  2. render_docx.py 设置字体/行距/缩进/排版
  3. footnote_circle_marks.py 把脚注编号改成圈码 ①②③
  4. renumber_footnotes_per_page.py 让圈码每页重新编号

用法：
  python build_docx.py <输入.md> <输出.docx> [--no-circle]

--no-circle：跳过第3、4步，只用 pandoc 默认的数字编号（1,2,3...从头到尾），
             用户可在 Word 里手动调圈码和每页重编号。
             适用于模型能力较弱或脚注过多（单页超20）的场景。

为什么要一键脚本：format skill 要求四个步骤依次跑完才能产出规范文档。
分步手动跑时，agent 可能自作主张跳过某步（如信任 render_docx 自带功能而不跑
renumber），导致圈码/每页重编号失败。一键脚本从机制上避免跳步。

依赖：pandoc, python-docx, pywin32 (Word 或 WPS)
"""
import os
import sys
import subprocess
import argparse
import tempfile
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent.resolve()


def find_pandoc():
    """找 pandoc 可执行文件"""
    # 常见安装路径
    candidates = [
        os.environ.get("PANDOC", ""),
        shutil.which("pandoc") if shutil else "",
        str(Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft/WinGet/Packages/JohnMacFarlane.Pandoc_Microsoft.Winget.Source_8wekyb3d8bbwe/pandoc-3.10/pandoc.exe"),
    ]
    for c in candidates:
        if c and Path(c).exists():
            return c
    return None

import shutil  # noqa: E402  (find_pandoc 用到)


def run_pandoc(md_path, tmp_docx):
    """第1步：pandoc markdown → docx。关闭 smart 扩展避免中文引号配对错误。"""
    pandoc = find_pandoc()
    if not pandoc:
        raise RuntimeError("找不到 pandoc，请安装或设置 PANDOC 环境变量")
    cmd = [pandoc, "--from=markdown-smart", "--to=docx", "-o", str(tmp_docx), str(md_path)]
    print(f"步骤1: pandoc 转换 {md_path.name} → {tmp_docx.name}")
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"pandoc 转换失败: {result.stderr[:300]}")
    if not tmp_docx.exists():
        raise RuntimeError(f"pandoc 未生成 {tmp_docx}")
    print("  ✓ 完成")
    return tmp_docx


def run_render(tmp_docx, out_docx):
    """第2步：render_docx.py 设置字体/行距/排版"""
    print("步骤2: render_docx.py 设置字体/行距/排版")
    script = SCRIPT_DIR / "render_docx.py"
    result = subprocess.run(
        [sys.executable, str(script), str(tmp_docx), str(out_docx)],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        print(f"  错误: {result.stderr[:500]}")
        raise RuntimeError("render_docx.py 失败")
    print("  ✓ 完成")


def run_circle_marks(docx_path):
    """第3步：footnote_circle_marks.py 把脚注编号改成圈码"""
    print("步骤3: footnote_circle_marks.py 圈码化")
    script = SCRIPT_DIR / "footnote_circle_marks.py"
    result = subprocess.run(
        [sys.executable, str(script), str(docx_path), str(docx_path)],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        print(f"  错误: {result.stderr[:500]}")
        raise RuntimeError("footnote_circle_marks.py 失败")
    print("  ✓ 完成")


def run_renumber(docx_path):
    """第4步：renumber_footnotes_per_page.py 每页重新编号"""
    print("步骤4: renumber_footnotes_per_page.py 每页重新编号")
    script = SCRIPT_DIR / "renumber_footnotes_per_page.py"
    result = subprocess.run(
        [sys.executable, str(script), str(docx_path), str(docx_path)],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        print(f"  错误: {result.stderr[:500]}")
        raise RuntimeError("renumber_footnotes_per_page.py 失败")
    # 打印脚本的输出（含每页分配信息）
    if result.stdout:
        for line in result.stdout.strip().split('\n'):
            print(f"  {line}")
    print("  ✓ 完成")


def check_ref_duplicates(docx_path):
    """参考文献去重校验：打开 docx，把“参考文献”标题后各条按“责任者+题名+版本”归一比对，
    打印重复项供人工核对（不自动删，只报告）。GB/T 7714 要求同一文献只列一次，
    人眼容易漏（如同一书不同位置被列为 [1] 和 [5]）。脚本跑一遍把它标出来更稳。"""
    try:
        from docx import Document
        import re
        doc = Document(str(docx_path))
    except Exception as e:
        print("  参考文献去重校验跳过（打不开 docx）")
        return
    # 收集“参考文献”标题后出现的条目段
    entries = []
    in_refs = False
    for p in doc.paragraphs:
        t = p.text.strip()
        if t == '参考文献':
            in_refs = True
            continue
        if not in_refs:
            continue
        # 条目形如 [1] 责任者.题名[...].出版地：出版社，年:页.
        if re.match(r'^\s*(?:\[\d+\]|［\d+］|①|\d+[.、])', t):
            # 归一 key：杂编号、页码，只留“责任者+题名+出版社+年”
            core = re.sub(r'^\s*(?:\[\d+\]|［\d+］|①|\d+[.、])\s*', '', t)
            # 去掉起止页码 [M]..:1-5 之类末尾页码与句末标点
            core = re.sub(r':[0-9\-]+\.?\s*$', '', core)
            core = re.sub(r'[\s。.]+$', '', core)
            core = core.strip()
            entries.append((core, t[:60]))
    if not entries:
        return
    seen = {}
    dups = []
    for core, disp in entries:
        if core in seen:
            dups.append((seen[core], disp))
        else:
            seen[core] = disp
    if dups:
        print("  ⚠ 参考文献·疑似重复（请人工核对是否同一文献）：")
        for a, b in dups:
            print(f"    - {a!r}")
            print(f"      ∟ {b!r}")
    else:
        print("  ✓ 参考文献无重复 (校验通过)")



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


def build(md_path, out_docx, use_circle=True, overwrite=False):
    md_path = Path(md_path)
    out_docx = Path(out_docx)
    if not md_path.exists():
        raise FileNotFoundError(f"找不到 {md_path}")

    if out_docx.exists() and not overwrite:
        raise FileExistsError(f"输出已存在，未覆盖：{out_docx}；另存或获授权后使用 --overwrite")
    if md_path.suffix.lower() != '.md':
        raise ValueError("本生成脚本只接受 .md；已有 docx 请用文档编辑流程")
    out_docx.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='.paper-build-', dir=out_docx.parent) as temp:
        tmp_docx = Path(temp) / 'pandoc.docx'
        ready = Path(temp) / 'verified.docx'
        run_pandoc(md_path, tmp_docx)
        run_render(tmp_docx, ready)
        if use_circle:
            run_circle_marks(ready)
            run_renumber(ready)
        from docx import Document
        Document(str(ready))
        check_ref_duplicates(ready)
        publish_file(ready, out_docx, overwrite)
    print(f"\n✅ 全部完成: {out_docx}")
    if not use_circle:
        print("  简单数字编号模式；圈码与分页尚未验证")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="一键把 markdown 论文转成规范 Word 文档")
    parser.add_argument("md", help="输入 markdown 文件路径")
    parser.add_argument("docx", help="输出 docx 文件路径")
    parser.add_argument("--no-circle", action="store_true",
                        help="不用圈码和每页重编号，pandoc 默认数字编号（用户手动调）")
    parser.add_argument("--overwrite", action="store_true", help="已获授权覆盖，保留旧文件备份")
    args = parser.parse_args()
    build(args.md, args.docx, use_circle=not args.no_circle, overwrite=args.overwrite)
