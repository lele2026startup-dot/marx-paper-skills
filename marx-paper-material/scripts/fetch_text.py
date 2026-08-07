# -*- coding: utf-8 -*-
"""
fetch_text.py — 网页全文抓取与清洗（马理论材料入库用）

单篇：
  python fetch_text.py --name "标题" --url "https://..." --out "目录/标题.txt"

批量（jobs.json = [{"name":..., "url":..., "out":...}, ...]）：
  python fetch_text.py --jobs jobs.json

要点（详见 references/sources.md）：
- 编码按「页面声明 → gb18030 → utf-8」顺序尝试，用乱码计数验证；
  老中文站（marxists.org 等）多为 GB18030，用 gb2312 会破坏"——"等字符。
- marxists.org 页面提取到文末"注释"区之前截止，不收脚注与页尾导航。
- 输出 UTF-8 txt，打印字数、乱码计数、首尾预览供人工核验。
"""
import argparse, json, re, sys, io, html as H
import requests

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}


def pick_encoding(content: bytes, declared: str):
    """按「声明 → gb18030 → utf-8」试解码，返回 (文本, 实际编码, 乱码数)。"""
    candidates = []
    if declared and declared.lower() not in ("iso-8859-1", "latin-1", "ascii"):
        candidates.append(declared)
    candidates += ["gb18030", "utf-8", "big5"]
    best = None
    for enc in candidates:
        try:
            text = content.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
        bad = text.count("\ufffd") + text.count("锟斤拷")
        if best is None or bad < best[2]:
            best = (text, enc, bad)
        if bad == 0:
            break
    return best


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


def fetch_one(name: str, url: str, out: str, encoding: str = None) -> dict:
    r = requests.get(url, headers=UA, timeout=90)
    r.raise_for_status()
    declared = encoding or r.headers.get("Content-Type", "").split("charset=")[-1].strip()
    result = pick_encoding(r.content, declared)
    if result is None:
        raise RuntimeError(f"{name}: 无法解码页面")
    text, enc, bad = result

    if "marxists.org" in url:
        body = extract_marxists(text)
    else:
        body = extract_generic(text)

    import os
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    with open(out, "w", encoding="utf-8") as f:
        f.write(body + "\n")

    final_bad = body.count("\ufffd") + body.count("锟斤拷")
    return {
        "name": name, "url": url, "out": out, "encoding": enc,
        "chars": len(body),
        "garbled": final_bad,
        "head": body[:60].replace("\n", " "),
        "tail": body[-60:].replace("\n", " "),
    }


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
            report(fetch_one(j["name"], j["url"], j["out"], j.get("encoding")))
            ok += 1
        except Exception as e:
            print(f"[FAIL] {j['name']}: {e}")
            fail += 1
    print(f"\n完成：{ok} 成功，{fail} 失败")
    if fail:
        sys.exit(1)


if __name__ == "__main__":
    main()
