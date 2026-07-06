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
        print(f"  pandoc 警告/错误: {result.stderr[:300]}")
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


def build(md_path, out_docx, use_circle=True):
    md_path = Path(md_path)
    out_docx = Path(out_docx)
    if not md_path.exists():
        raise FileNotFoundError(f"找不到 {md_path}")

    tmp_docx = out_docx.with_suffix(".tmp.docx")

    # 第1步：pandoc
    run_pandoc(md_path, tmp_docx)

    # 第2步：render_docx
    run_render(tmp_docx, out_docx)
    tmp_docx.unlink(missing_ok=True)

    if use_circle:
        # 第3、4步：圈码化 + 每页重编号
        run_circle_marks(out_docx)
        run_renumber(out_docx)

    print(f"\n✅ 全部完成: {out_docx}")
    if not use_circle:
        print("  （--no-circle 模式：脚注为 pandoc 默认数字编号，请在 Word 里手动调圈码和每页重编号）")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="一键把 markdown 论文转成规范 Word 文档")
    parser.add_argument("md", help="输入 markdown 文件路径")
    parser.add_argument("docx", help="输出 docx 文件路径")
    parser.add_argument("--no-circle", action="store_true",
                        help="不用圈码和每页重编号，pandoc 默认数字编号（用户手动调）")
    args = parser.parse_args()
    build(args.md, args.docx, use_circle=not args.no_circle)
