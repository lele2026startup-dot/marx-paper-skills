# -*- coding: utf-8 -*-
"""
fetch_text.py — 网页全文抓取与清洗（马理论材料入库用）

单篇：
  python fetch_text.py --name "标题" --url "https://..." --out "目录/标题.txt"

批量（jobs.json = [{"name":..., "url":..., "out":...}, ...]）：
  python fetch_text.py --jobs jobs.json

要点（详见 references/sources.md）：
- 编码按显式指定/BOM/有效声明/严格UTF-8/旧编码处理，零乱码计数不替代原页面核对；
  老中文站（marxists.org 等）多为 GB18030，用 gb2312 会破坏"——"等字符。
- marxists.org 页面提取到文末"注释"区之前截止，不收脚注与页尾导航。
- 输出 UTF-8 txt，打印字数、乱码计数、首尾预览供人工核验。
"""
import argparse, json, re, sys, io, html as H
from pathlib import Path
from urllib.request import Request, urlopen
from urllib.parse import urlparse
import codecs
import tempfile
import os

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}


def pick_encoding(content: bytes, declared: str = '', forced: str = None):
    """Honor explicit encoding; otherwise BOM, reliable declaration, strict UTF-8, legacy."""
    if forced:
        text = content.decode(forced, errors='strict')
        return text, forced, text.count('\ufffd') + text.count('锟斤拷')
    if content.startswith(codecs.BOM_UTF8):
        return content.decode('utf-8-sig'), 'utf-8-sig', 0
    if content.startswith((codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE)):
        return content.decode('utf-16'), 'utf-16', 0
    head = content[:8192].decode('ascii', errors='ignore')
    meta = re.search(r'<meta\b[^>]*charset\s*=\s*[\'\"]?\s*([\w.-]+)', head, re.I)
    candidates = []
    for enc in (declared, meta.group(1) if meta else None):
        if enc and enc.strip().lower() not in ('iso-8859-1', 'latin-1', 'ascii'):
            enc = enc.strip().strip('\'\"')
            if enc.lower() in ('gb2312', 'gbk', 'gb_2312-80'):
                enc = 'gb18030'
            candidates.append(enc)
    candidates.extend(['utf-8', 'gb18030', 'big5'])
    seen = set()
    for enc in candidates:
        if enc.lower() in seen:
            continue
        seen.add(enc.lower())
        try:
            text = content.decode(enc, errors='strict')
        except (UnicodeDecodeError, LookupError):
            continue
        bad = text.count('\ufffd') + text.count('锟斤拷')
        if bad == 0:
            return text, enc, bad
    return None


def strip_tags_fragment(html_fragment: str) -> str:
    """把一段 HTML 转成纯文本：标题转 ##，段落换行，实体解码。"""
    t = html_fragment
    t = re.sub(r"<script[\s\S]*?</script>", "", t, flags=re.I)
    t = re.sub(r"<style[\s\S]*?</style>", "", t, flags=re.I)
    t = re.sub(r"<h([1-6])[^>]*>", lambda m: "\n## ", t, flags=re.I)
    t = re.sub(r"</h[1-6]>", "\n", t, flags=re.I)
    t = re.sub(r"<br[^>]*>", "\n", t, flags=re.I)
    t = re.sub(r"<p[^>]*>", "\n", t, flags=re.I)
    t = re.sub(r"<[^>]+>", "", t)
    t = H.unescape(t)
    # 清理行内残留的不完整标签碎片
    t = re.sub(r"<[^>\n]{0,80}$", "", t, flags=re.M)
    t = t.replace("\u3000", "\u3000")  # 全角空格保留（原文缩进）
    lines = [ln.rstrip() for ln in t.split("\n")]
    out, blank = [], 0
    nav_pat = re.compile(r'(上一页|下一页|目\s*录|返回目录|首\s*页|返回)(\s+(上一页|下一页|目\s*录|返回目录|首\s*页|返回))*')
    crumb_pat = re.compile(r'^[（(]?\d{1,4}[）)]$')
    for ln in lines:
        s = ln.strip()
        if s and (nav_pat.fullmatch(s) or crumb_pat.match(s)):
            continue  # 去掉导航行和面包屑残留
        if s:
            out.append(ln)
            blank = 0
        else:
            blank += 1
            if blank == 1:
                out.append("")
    # 首行开头可能粘着面包屑残留（如"035） 标题"）：仅当后面还跟着正文时才剥
    for idx, ln in enumerate(out):
        if ln.strip():
            out[idx] = re.sub(r'^\s*\d{1,4}[）)]\s+(?=\S)', '', ln, count=1)
            break
    return "\n".join(out).strip()


def extract_marxists(page: str) -> str:
    """marxists.org 经典著作页：title1 标题起，文末注释区止。"""
    # 定位正文起点
    m = re.search(r"<p class=['\"]?title1['\"]?[^>]*>", page, flags=re.I)
    if not m:
        m = re.search(r"<h1[^>]*>", page, flags=re.I)
    start = m.start() if m else page.find("<body")
    if start < 0:
        start = 0
    body = page[start:]
    # 定位注释区终点：最后一个 hr 后跟"注释"，或"注释"字样独立出现
    cut = None
    for mm in re.finditer(r"<hr[^>]*>\s*(?:<[^>]+>\s*)*[^<]*注释", body, flags=re.I):
        cut = mm.start()
    if cut is None:
        mm = re.search(r"<(?:p|div|span)[^>]*>\s*(?:<b>)?\s*注\s*释", body, flags=re.I)
        if mm:
            cut = mm.start()
    if cut is not None:
        body = body[:cut]
    return strip_tags_fragment(body)


def extract_generic(page: str) -> str:
    """通用页面：去导航/页脚后取 body。"""
    t = re.sub(r"<!DOCTYPE[^>]*>", "", page, flags=re.I)
    t = re.sub(r"<(nav|header|footer|aside)[\s\S]*?</\1>", "", t, flags=re.I)
    m = re.search(r"<body[\s\S]*?</body>", t, flags=re.I)
    if m:
        t = m.group(0)
    return strip_tags_fragment(t)


def fetch_one(name: str, url: str, out: str, encoding: str = None, overwrite=False) -> dict:
    target = Path(out)
    if target.exists() and not overwrite:
        raise FileExistsError(f'输出已存在，未覆盖：{out}；另存或获授权后使用 --overwrite')
    with urlopen(Request(url, headers=UA), timeout=90) as response:
        content = response.read()
        declared = response.headers.get_content_charset() or ''
        final_url = response.geturl()
    result = pick_encoding(content, declared, forced=encoding)
    if result is None:
        raise RuntimeError(f'{name}: 无法可靠解码，请检查原页面或指定 --encoding')
    text, enc, bad = result
    host = (urlparse(final_url).hostname or '').lower()
    body = extract_marxists(text) if host == 'marxists.org' or host.endswith('.marxists.org') else extract_generic(text)
    final_bad = body.count('\ufffd') + body.count('锟斤拷')
    if not body.strip() or final_bad:
        raise ValueError(f'{name}: 正文为空或有疑似乱码，未写入输出')
    target.parent.mkdir(parents=True, exist_ok=True)
    fd, temp = tempfile.mkstemp(prefix='.' + target.name + '-', suffix='.tmp', dir=target.parent)
    try:
        with os.fdopen(fd, 'w', encoding='utf-8', newline='\n') as stream:
            stream.write(body + '\n')
        publish_file(temp, target, overwrite)
    finally:
        Path(temp).unlink(missing_ok=True)
    return {'name': name, 'url': final_url, 'out': str(target), 'encoding': enc,
            'chars': len(body), 'garbled': final_bad,
            'head': body[:60].replace('\n', ' '), 'tail': body[-60:].replace('\n', ' ')}



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


def report(r: dict):
    flag = "OK" if r["garbled"] == 0 else f"!!乱码{r['garbled']}处"
    print(f"[{flag}] {r['name']}")
    print(f"    编码={r['encoding']} 字数={r['chars']} → {r['out']}")
    print(f"    首: {r['head']}")
    print(f"    尾: {r['tail']}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--name")
    ap.add_argument("--url")
    ap.add_argument("--out")
    ap.add_argument("--encoding", default=None, help="强制指定编码（默认自动探测）")
    ap.add_argument("--jobs", help="批量：JSON 文件，元素为 {name,url,out[,encoding]}")
    ap.add_argument("--overwrite", action="store_true", help="已获授权覆盖；先保存原文件备份")
    args = ap.parse_args()

    if args.jobs:
        jobs = json.load(open(args.jobs, encoding="utf-8"))
    elif args.name and args.url and args.out:
        jobs = [{"name": args.name, "url": args.url, "out": args.out, "encoding": args.encoding}]
    else:
        ap.error("需要 --name --url --out 或 --jobs")

    ok, fail = 0, 0
    for j in jobs:
        try:
            report(fetch_one(j["name"], j["url"], j["out"], j.get("encoding"), overwrite=args.overwrite))
            ok += 1
        except Exception as e:
            print(f"[FAIL] {j['name']}: {e}")
            fail += 1
    print(f"\n完成：{ok} 成功，{fail} 失败")
    if fail:
        sys.exit(1)


if __name__ == "__main__":
    main()
