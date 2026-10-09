# -*- coding: utf-8 -*-
"""从华为 HDX（未加密 ZIP）导出指定章节目录为可离线阅读的静态站点。

用法：
    python export_chapter.py --chapter "用户手册 > iFC用户手册" --out "iFC用户手册离线版"

    --hdx     HDX 文件路径
    --chapter 目标章节路径，用 " > " 分隔目录标题，逐级精确匹配
    --out     输出目录

设计要点：
    * 目录树来自 resources/navi.xml，正文文件名不写死。
    * ZIP 内路径大小写与 navi 引用常不一致，统一建立大小写映射并改写引用。
    * 页面按声明编码解码，转 UTF-8 并同步修改 charset 声明。
    * 输出镜像 ZIP 内的相对目录结构，使原有相对引用天然成立。
    * 章节外链接改写到 viewer/external.html，保留原目标信息，不伪造正文。
"""
from __future__ import annotations

import argparse
import json
import os
import posixpath
import re
import sys
import time
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from urllib.parse import quote, unquote

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

CHARSET_META_RE = re.compile(r'(charset\s*=\s*["\']?)([A-Za-z0-9_\-]+)', re.I)
CSS_CHARSET_RE = re.compile(rb'@charset\s+"([^"]+)"', re.I)
CSS_URL_RE = re.compile(r'url\(\s*(["\']?)([^"\')]+)\1\s*\)', re.I)
CSS_IMPORT_RE = re.compile(r'@import\s+(?:url\(\s*)?["\']?([^"\')\s;]+)', re.I)
JS_LITERAL_RE = re.compile(
    r'["\']([A-Za-z0-9_./\-]+\.(?:png|jpg|jpeg|gif|svg|css|js|xml|json))["\']', re.I)

ASSET_EXTS = (".png", ".jpg", ".jpeg", ".gif", ".svg", ".bmp", ".ico",
              ".css", ".js", ".woff", ".woff2", ".ttf", ".eot")
ATTACH_EXTS = (".zip", ".rar", ".7z", ".pdf", ".doc", ".docx", ".xls", ".xlsx",
               ".ppt", ".pptx", ".csv", ".tar", ".gz", ".bat", ".sh")
PAGE_EXTS = (".html", ".htm")
SKIP_SCHEMES = ("#", "http://", "https://", "mailto:", "javascript:", "ftp:", "data:", "tel:")
DECODE_FALLBACKS = ("gb18030", "utf-8", "big5", "cp1252", "latin-1")
MARKER = ".hdx_export_marker.json"

STUB_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-cn">
<head>
<meta charset="utf-8">
<title>本次章节样例未包含该内容</title>
<style>
 body{font-family:"Microsoft YaHei","Segoe UI",sans-serif;margin:0;padding:40px;
      background:#f7f8fa;color:#1f2329}
 .card{max-width:760px;margin:0 auto;background:#fff;border:1px solid #e5e6eb;
       border-radius:10px;padding:28px 32px}
 h1{font-size:19px;margin:0 0 6px}
 p{line-height:1.75;font-size:14px;color:#4e5969}
 code{background:#f2f3f5;padding:2px 6px;border-radius:4px;font-size:13px;
      word-break:break-all}
 .tag{display:inline-block;background:#fff1e8;color:#d4380d;border:1px solid #ffd8bf;
      border-radius:4px;padding:2px 8px;font-size:12px;margin-bottom:14px}
</style>
</head>
<body>
<div class="card">
  <div class="tag">章节外引用</div>
  <h1>本次离线样例未包含该内容</h1>
  <p>目标页面不在本次导出的章节范围内，
     因此未随本离线包一并导出。</p>
  <p><strong>原始目标：</strong><br><code id="target"></code></p>
  <p><strong>来源页面：</strong><br><code id="source"></code></p>
  <p style="color:#86909c">若需完整内容，请在原始 HDX
     文档包或 ICS Lite 中查看。</p>
</div>
<script>
var h = decodeURIComponent(location.hash.replace(/^#/, ""));
var parts = h.split("|");
document.getElementById("target").textContent = parts[0] || "(未记录)";
document.getElementById("source").textContent = parts[1] || "(未记录)";

/* 明确告诉阅读器：当前是“章节外说明页”，不是任何正文页。
   阅读器据此清空正文高亮并提供返回入口，不会把它当成清单里的某一项。 */
function reportExternal() {
  try {
    window.parent.postMessage({ __hdx: 1, type: 'external',
      payload: location.hash.replace(/^#/, ""),
      target: parts[0] || '', source: parts[1] || '' }, '*');
  } catch (e) { /* 独立打开时无父窗口 */ }
}
if (document.readyState === 'complete') { reportExternal(); }
window.addEventListener('load', reportExternal);
window.addEventListener('pageshow', reportExternal);
</script>
</body>
</html>
"""


class PackageError(RuntimeError):
    pass


class Package:
    """HDX / ZIP 包，提供大小写不敏感路径解析。"""

    def __init__(self, path):
        self.path = path
        if not os.path.isfile(path):
            raise PackageError(f"找不到 HDX 文件：{path}")
        try:
            self.z = zipfile.ZipFile(path)
        except zipfile.BadZipFile as e:
            raise PackageError(f"HDX 不是可读取的 ZIP：{path}（{e}）") from e
        self.names = [n for n in self.z.namelist() if not n.endswith("/")]
        self._lower = {}
        self.collisions = []
        for n in self.names:
            k = n.lower()
            if k in self._lower and self._lower[k] != n:
                self.collisions.append((self._lower[k], n))
            else:
                self._lower[k] = n
        if self.collisions:
            raise PackageError(
                "ZIP 内存在仅大小写不同的路径，无法唯一解析：\n  "
                + "\n  ".join(f"{a}  <->  {b}" for a, b in self.collisions))

    def resolve(self, rel):
        return self._lower.get(rel.replace("\\", "/").lower())

    def read(self, name):
        return self.z.read(name)


# ---------------------------------------------------------------- navi.xml
def find_by_path(root, path_parts):
    current = [root]
    for part in path_parts:
        nxt = []
        for el in current:
            for c in el.findall("topic"):
                if c.get("txt", "") == part:
                    nxt.append(c)
        current = nxt
        if not current:
            return []
    return current


def collect_subtree(node):
    out = []

    def walk(e, depth):
        for c in e.findall("topic"):
            out.append({"depth": depth, "txt": c.get("txt", ""),
                        "url": c.get("url", ""), "id": c.get("id", "")})
            walk(c, depth + 1)

    out.append({"depth": 0, "txt": node.get("txt", ""),
                "url": node.get("url", ""), "id": node.get("id", "")})
    walk(node, 1)
    return out


# ------------------------------------------------------------------ 编码
def detect_charset(raw):
    m = CHARSET_META_RE.search(raw[:4096].decode("latin-1", "replace"))
    if m:
        return m.group(2).lower()
    m = CSS_CHARSET_RE.search(raw[:4096])
    return m.group(1).decode("ascii", "replace").lower() if m else None


def decode_text(raw, label, notes):
    enc = detect_charset(raw)
    cands = []
    if enc:
        cands.append(enc)
        if enc in ("gb2312", "gbk"):
            cands.append("gb18030")
    cands.extend(c for c in DECODE_FALLBACKS if c not in cands)
    for c in cands:
        try:
            text = raw.decode(c)
            if c != enc:
                notes.append(f"{label}: 声明编码 {enc or '无'}，实际按 {c} 解码")
            return text, c
        except (UnicodeDecodeError, LookupError):
            continue
    notes.append(f"{label}: 所有候选编码均失败，按 latin-1 有损解码")
    return raw.decode("latin-1"), "latin-1"


# -------------------------------------------------------------- 引用改写
def visible_text(html):
    """标签之外的可见文本，用于源/输出一致性比对。"""
    no_script = re.sub(r'(?is)<(script|style)\b.*?</\1>', " ", html)
    return re.sub(r'\s+', ' ', " ".join(re.split(r'<[^>]*>', no_script))).strip()


def rewrite_attrs(text, fn):
    """对 href/src 属性值应用 fn，保留原引号。"""
    for q in ('"', "'"):
        pat = re.compile(r'(\b(?:href|src)\s*=\s*' + q + r')([^' + q + r']*)(' + q + r')')
        text = pat.sub(lambda m: m.group(1) + fn(m.group(2)) + m.group(3), text)
    return text


def split_ref(raw):
    """把引用拆成 (路径, 后缀)，后缀含原有的 ?query 与 #fragment。

    只有在定位包内文件时才需要去掉 query/fragment；改写回 HTML 时要把它们原样带上。
    """
    path, sep_f, frag = raw.partition("#")
    path, sep_q, query = path.partition("?")
    suffix = (sep_q + query if sep_q else "") + (sep_f + frag if sep_f else "")
    return path, suffix


class Rewriter:
    """按规范化后的完整包内路径判断章节归属，改写 href/src 为输出目录中的真实路径。"""

    def __init__(self, pkg, chapter_paths, stub_rel):
        self.pkg = pkg
        # 实际导出的正文路径集合（已解析大小写的完整包内路径），不用 basename 判断
        self.chapter_paths = set(chapter_paths)
        self.stub_rel = stub_rel
        self.stats = defaultdict(int)
        self.assets = set()
        # 三类目标分别记录
        self.internal = defaultdict(set)         # 章节内：原引用 -> 来源页
        self.external_in_pkg = defaultdict(set)  # 章节外且在包内：实际包内路径 -> 来源页
        self.external_absent = defaultdict(set)  # 章节外且包内不存在：原引用 -> 来源页
        self.missing = defaultdict(set)          # 非页面资源未能解析：原引用 -> 来源页

    def _stub(self, src_zip, original):
        payload = quote(original, safe="")
        src = quote(src_zip, safe="")
        return rel_href(src_zip, self.stub_rel) + f"#{payload}|{src}"

    def rewrite_one(self, value, src_zip):
        raw = value.strip()
        if not raw or raw.startswith(SKIP_SCHEMES):
            return value
        path, suffix = split_ref(raw)
        if not path:
            return value                      # 纯 #fragment，交给浏览器，原样保留

        decoded = unquote(path)               # 包内文件名可能是 URL 编码形式
        guess = posixpath.normpath(posixpath.join(posixpath.dirname(src_zip), decoded))
        actual = self.pkg.resolve(guess)      # 解析大小写，得到规范化的完整包内路径

        if decoded.lower().endswith(PAGE_EXTS):
            if actual is not None and actual in self.chapter_paths:
                self.internal[raw].add(src_zip)
                self.stats["internal_links"] += 1
                return quote(rel_href(src_zip, actual)) + suffix
            if actual is not None:
                self.external_in_pkg[actual].add(src_zip)
                self.stats["external_in_pkg"] += 1
            else:
                self.external_absent[raw].add(src_zip)
                self.stats["external_absent_in_pkg"] += 1
            return self._stub(src_zip, raw)

        if actual is not None:
            self.assets.add(actual)
            self.stats["assets"] += 1
            new = rel_href(src_zip, actual)
            if new != path:
                self.stats["asset_case_fixed"] += 1
            return quote(new) + suffix

        self.stats["assets_missing"] += 1
        self.missing[raw].add(src_zip)
        return value

    def anchor_data(self, href, src_zip):
        """为 <a> 生成供阅读器拦截点击用的数据属性（在 href 改写之后调用）。

        章节内链接标记 data-hdx / data-hdx-a；章节外说明页标记 data-hdx-ext，
        并把原始目标带过去，保证说明页能显示原信息。
        """
        raw = href.strip()
        if not raw or raw.startswith(SKIP_SCHEMES) or raw.startswith("#"):
            return ""
        path, suffix = split_ref(raw)
        if not path:
            return ""
        decoded = unquote(path)
        guess = posixpath.normpath(posixpath.join(posixpath.dirname(src_zip), decoded))
        if guess == self.stub_rel:                       # 指向章节外说明页
            frag = suffix.split("#", 1)[1] if "#" in suffix else ""
            return 'data-hdx-ext="%s"' % html_escape(frag)
        actual = self.pkg.resolve(guess)
        if actual is not None and actual in self.chapter_paths:
            frag = suffix.split("#", 1)[1] if "#" in suffix else ""
            query = suffix.split("#", 1)[0]
            query = query[1:] if query.startswith("?") else ""
            s = 'data-hdx="%s"' % html_escape(actual)
            if query:
                # ?query 必须一起带过去：阅读器按 key 重新拼地址时，
                # 只有 key 和 #锚点的话会把 query 丢掉。
                s += ' data-hdx-q="%s"' % html_escape(query)
            if frag:
                s += ' data-hdx-a="%s"' % html_escape(frag)
            return s
        return ""


def rel_href(from_zip, to_zip):
    return posixpath.relpath(to_zip, posixpath.dirname(from_zip))


# --------------------------------------------------- 正文与阅读器的状态同步
# 注入到每个正文页末尾：告诉阅读器“当前显示的是哪一页”。
# file:// 下父页面无法读取 iframe 的 DOM/location，因此用 postMessage 上报；
# 父页面用 event.source===iframe.contentWindow 校验来源，并用页面清单校验 key。
SYNC_JS = """/* 由 export_chapter.py 注入；只做两件事，都不改变正文内容：
   1. 在配套阅读器里打开时，把章节内 / 章节外跳转交给阅读器统一导航；
   2. 上报当前真实地址，供阅读器同步高亮、标题、上一项/下一项与锚点。
   只有与阅读器握手成功后（inReader）才接管点击：单独打开正文页
   ——window.parent === window，例如双击 HTML、拖进浏览器或在新标签页打开——
   时一律不拦截，链接走浏览器原生跳转，路径、?query 与 #锚点都不会丢。
   阅读器用 event.source 校验来源、用页面清单校验 key。 */
(function () {
  var P = window.__HDX_PAGE__;
  var inReader = false;      // 与父窗口握手成功之前，不接管任何点击

  function post(msg) {
    if (window.parent === window) { return; }   // 独立打开：没有父窗口可通信
    try { window.parent.postMessage(msg, '*'); } catch (e) {}
  }

  function reportPage() {
    if (!P || !P.k) { return; }
    post({ __hdx: 1, type: 'page', k: P.k,
           hash: location.hash || '', search: location.search || '' });
  }

  // 章节内链接带 data-hdx，章节外链接带 data-hdx-ext，两者都由导出脚本标注；
  // ?query 与 #fragment 另由 data-hdx-q / data-hdx-a 带过去，避免拼地址时丢失。
  function handOff(a) {
    if (!inReader) { return false; }   // 不在阅读器内：交给浏览器原生导航
    var k = a.getAttribute('data-hdx');
    if (k) {
      post({ __hdx: 1, type: 'link', k: k,
             q: a.getAttribute('data-hdx-q') || '',
             a: a.getAttribute('data-hdx-a') || '' });
      return true;
    }
    if (a.hasAttribute('data-hdx-ext')) {
      post({ __hdx: 1, type: 'extlink',
             payload: a.getAttribute('data-hdx-ext') || '' });
      return true;
    }
    return false;
  }

  document.addEventListener('click', function (ev) {
    if (ev.defaultPrevented || ev.button !== 0) { return; }
    if (ev.metaKey || ev.ctrlKey || ev.shiftKey || ev.altKey) { return; }
    var el = ev.target;
    while (el && el.nodeType !== 1) { el = el.parentNode; }
    while (el && el !== document && el.tagName !== 'A') { el = el.parentNode; }
    if (!el || el.tagName !== 'A') { return; }
    if (handOff(el)) { ev.preventDefault(); }
  }, true);

  // 握手：先询问，父窗口确认自己是阅读器后回 hello，收到才启用消息导航。
  // 只认 window.parent 发来的消息（e.source），其它窗口一律忽略。
  window.addEventListener('message', function (e) {
    if (e.source !== window.parent) { return; }
    var d = e.data;
    if (d && d.__hdx === 1 && d.type === 'hello') {
      inReader = true;
      reportPage();
    }
  });

  function hello() { post({ __hdx: 1, type: 'hi' }); }

  if (document.readyState === 'complete') { reportPage(); }
  window.addEventListener('load', function () { hello(); reportPage(); });
  window.addEventListener('pageshow', function () { hello(); reportPage(); });
  window.addEventListener('hashchange', reportPage); // 页内锚点
  hello();   // 父窗口可能已经就绪；未就绪时上面的 load/pageshow 会再问一次
})();
"""


def html_escape(s):
    return (s.replace("&", "&amp;").replace('"', "&quot;")
            .replace("<", "&lt;").replace(">", "&gt;"))


ANCHOR_TAG_RE = re.compile(r"<a\b[^>]*>", re.I)


def annotate_anchors(text, rw, src_zip):
    """给 <a> 补数据属性，供阅读器拦截点击。只改标签属性，不动正文内容。"""
    def sub(m):
        tag = m.group(0)
        hm = re.search(r'\bhref\s*=\s*(["\'])(.*?)\1', tag, re.S)
        if not hm:
            return tag
        add = rw.anchor_data(hm.group(2), src_zip)
        if not add:
            return tag
        return tag[:-1] + " " + add + ">"

    return ANCHOR_TAG_RE.sub(sub, text)


def inject_sync(html, page_key, sync_rel):
    """在正文页末尾注入同步脚本。只新增 <script>，不动正文结构、表格、图片与锚点。"""
    tag = ('<script>window.__HDX_PAGE__={"k":"%s"};</script>\n'
           '<script src="%s"></script>\n' % (page_key, sync_rel))
    low = html.lower()
    pos = low.rfind("</body>")
    if pos == -1:
        return html + "\n" + tag
    return html[:pos] + tag + html[pos:]


# ------------------------------------------------------------------ 输出
def safe_join(out_dir, rel):
    p = os.path.normpath(os.path.join(out_dir, rel.replace("/", os.sep)))
    root = os.path.normpath(out_dir)
    if not (p == root or p.startswith(root + os.sep)):
        raise PackageError(f"输出路径越界，已阻止：{rel}")
    return p


def write_out(out_dir, rel, data):
    p = safe_join(out_dir, rel)
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "wb") as f:
        f.write(data)
    return p


# ------------------------------------------------------------------ 主流程
def main():
    ap = argparse.ArgumentParser(description="从 HDX 导出指定章节为离线站点")
    ap.add_argument("--hdx",
                    default=r"C:\code\AWcode\华为\HUAWEI IMS IP多媒体子系统解决方案 V500R019C50 产品文档 06\IMS_V500R019C50_06_zh_CZM0530J.hdx",
                    help="HDX 文件路径")
    ap.add_argument("--chapter", default="用户手册 > iFC用户手册",
                    help='章节路径，如 "用户手册 > iFC用户手册"')
    ap.add_argument("--out", default="iFC用户手册离线版", help="输出目录")
    ap.add_argument("--navi", default="resources/navi.xml")
    ap.add_argument("--stub", default="viewer/external.html")
    ap.add_argument("--force", action="store_true",
                    help="允许写入非本工具生成的非空目录")
    args = ap.parse_args()

    # 只覆盖本工具生成的内容，不对无关目录动手
    marker = os.path.join(args.out, MARKER)
    if os.path.isdir(args.out) and os.listdir(args.out) \
            and not os.path.exists(marker) and not args.force:
        raise PackageError(
            f"输出目录已存在且不是本工具生成的，拒绝写入：{os.path.abspath(args.out)}\n"
            f"  确认要继续请加 --force，或换一个 --out 目录。")

    t0 = time.time()
    notes, log = [], []

    def say(s):
        print(s, flush=True)
        log.append(s)

    pkg = Package(args.hdx)
    say(f"HDX 打开成功：{args.hdx}")
    say(f"ZIP 条目数：{len(pkg.names)}（仅大小写冲突：{len(pkg.collisions)}）")

    navi_name = pkg.resolve(args.navi)
    if navi_name is None:
        raise PackageError(f"包内找不到目录文件 {args.navi}")
    root = ET.fromstring(pkg.read(navi_name))

    parts = [p.strip() for p in args.chapter.split(">")]
    matches = find_by_path(root, parts)
    if not matches:
        raise PackageError(f"未找到章节：{args.chapter}")
    if len(matches) > 1:
        raise PackageError(f"章节路径不唯一（{len(matches)} 个），请用更完整的标题路径")
    say(f"章节定位：{' > '.join(parts)} -> 1 个节点")

    nodes = collect_subtree(matches[0])
    for n in nodes:
        actual = pkg.resolve("resources/" + n["url"])
        if actual is None:
            raise PackageError(f"目录节点正文缺失：{n['txt']} -> {n['url']}")
        n["zip"] = actual
    chapter_paths = sorted(set(n["zip"] for n in nodes))
    unique_files = len(chapter_paths)
    say(f"目录节点数（含章节首页）：{len(nodes)}，唯一正文文件数：{unique_files}")
    dup_nodes = len(nodes) - unique_files
    if dup_nodes:
        say(f"提示：{dup_nodes} 个目录节点与其它节点指向同一正文文件（正文只提取一次）")

    rw = Rewriter(pkg, chapter_paths, args.stub)

    # 正文页（同一正文被多个目录节点引用时只提取一次）
    pages, body_info = [], {}
    for n in nodes:
        if n["zip"] not in body_info:
            raw = pkg.read(n["zip"])
            text, enc_used = decode_text(raw, n["url"], notes)
            before = visible_text(text)
            new_text = rewrite_attrs(text, lambda v, z=n["zip"]: rw.rewrite_one(v, z))
            new_text = annotate_anchors(new_text, rw, n["zip"])
            new_text = CHARSET_META_RE.sub(lambda m: m.group(1) + "utf-8", new_text, count=1)
            new_text = inject_sync(new_text, n["zip"], rel_href(n["zip"], "viewer/sync.js"))
            write_out(args.out, n["zip"], new_text.encode("utf-8"))
            body_info[n["zip"]] = {
                "src_bytes": len(raw), "src_encoding": enc_used,
                "text_intact": before == visible_text(new_text),
                "src_text_len": len(before), "out_text_len": len(visible_text(new_text)),
            }
        pages.append({
            "txt": n["txt"], "depth": n["depth"], "id": n["id"],
            "navi_url": n["url"], "zip": n["zip"], "out": n["zip"],
            **body_info[n["zip"]],
        })
    say(f"已导出正文页：{len(pages)}（唯一文件 {unique_files}）")

    # 资源闭包
    processed, css_fixed = set(), 0
    queue = [a for a in rw.assets if a.lower().endswith((".css", ".js"))]
    js_literals_missing = []
    while queue:
        a = queue.pop()
        if a in processed:
            continue
        processed.add(a)
        raw = pkg.read(a)
        if a.lower().endswith(".css"):
            text, _ = decode_text(raw, a, notes)

            def css_sub(m):
                nonlocal css_fixed
                q, val = m.group(1), m.group(2).strip()
                if val.startswith(("data:", "http:", "https:")):
                    return m.group(0)
                tgt = posixpath.normpath(posixpath.join(
                    posixpath.dirname(a), val.split("?")[0].split("#")[0]))
                real = pkg.resolve(tgt)
                if not real:
                    rw.missing[val].add(a)
                    return m.group(0)
                rw.assets.add(real)
                if real.lower().endswith(".css"):
                    queue.append(real)
                newv = rel_href(a, real)
                if newv != val:
                    css_fixed += 1
                return f"url({q}{newv}{q})"

            text = CSS_URL_RE.sub(css_sub, text)

            def imp_sub(m):
                tgt = posixpath.normpath(posixpath.join(
                    posixpath.dirname(a), m.group(1).split("?")[0]))
                real = pkg.resolve(tgt)
                if real:
                    rw.assets.add(real)
                    queue.append(real)
                    return f'@import "{rel_href(a, real)}"'
                return m.group(0)

            text = CSS_IMPORT_RE.sub(imp_sub, text)
            write_out(args.out, a, text.encode("utf-8"))
        else:
            write_out(args.out, a, raw)
            for lit in set(JS_LITERAL_RE.findall(raw.decode("latin-1", "replace"))):
                tgt = posixpath.normpath(posixpath.join(posixpath.dirname(a), lit))
                real = pkg.resolve(tgt)
                if real:
                    rw.assets.add(real)
                else:
                    js_literals_missing.append((a, lit))

    for a in sorted(rw.assets):
        if a in processed:
            continue
        write_out(args.out, a, pkg.read(a))
    say(f"已导出资源：{len(rw.assets)} 个（CSS 内路径修正 {css_fixed} 处）")

    images = [p for p in rw.assets if p.lower().endswith((".png", ".jpg", ".jpeg", ".gif", ".svg", ".bmp", ".ico"))]
    atts = [p for p in rw.assets if p.lower().endswith(ATTACH_EXTS)]

    result = {
        "chapter_path": args.chapter,
        "hdx": args.hdx,
        "generated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "node_count": len(pages),
        "unique_body_files": unique_files,
        "pages": pages,
        "assets": sorted(rw.assets),
        "asset_count": len(rw.assets),
        "image_count": len(images),
        "attachments": sorted(atts),
        "attachment_count": len(atts),
        # 三类目标分别记录：章节内 / 章节外（包内存在）/ 章节外（包内不存在）
        "internal_links": {k: sorted(v) for k, v in sorted(rw.internal.items())},
        "internal_link_targets": len(rw.internal),
        "external_in_pkg": {k: sorted(v) for k, v in sorted(rw.external_in_pkg.items())},
        "external_in_pkg_count": len(rw.external_in_pkg),
        "external_absent": {k: sorted(v) for k, v in sorted(rw.external_absent.items())},
        "external_absent_count": len(rw.external_absent),
        "external_count": len(rw.external_in_pkg) + len(rw.external_absent),
        "missing_refs": {k: sorted(v) for k, v in sorted(rw.missing.items())},
        "js_literal_missing": js_literals_missing,
        "encoding_notes": notes,
        "rewrite_stats": dict(rw.stats),
    }
    write_out(args.out, "reports/manifest.json",
              json.dumps(result, ensure_ascii=False, indent=2).encode("utf-8"))

    write_out(args.out, args.stub, STUB_TEMPLATE.encode("utf-8"))
    build_viewer(args.out, result)
    write_readme(args.out, result)

    write_out(args.out, MARKER, json.dumps({
        "tool": "export_chapter.py", "chapter_path": args.chapter,
        "hdx": args.hdx, "generated_at": result["generated_at"],
    }, ensure_ascii=False, indent=2).encode("utf-8"))

    elapsed = time.time() - t0
    say(f"导出总用时：{elapsed:.2f}s")
    with open(safe_join(args.out, "reports/export_log.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(log) + "\n")
    print(f"\n输出目录：{os.path.abspath(args.out)}")
    return 0


# ------------------------------------------------------------- 阅读入口
VIEWER_CSS = """:root{--bd:#e5e6eb;--fg:#1f2329;--mut:#86909c;--acc:#0a6cff;--bg:#fff;--side:#f7f8fa}
*{box-sizing:border-box}
html,body{margin:0;height:100%;font-family:"Microsoft YaHei","PingFang SC","Segoe UI",sans-serif;color:var(--fg)}
#app{display:flex;height:100vh;overflow:hidden}
#side{width:330px;min-width:220px;max-width:60vw;background:var(--side);border-right:1px solid var(--bd);
      display:flex;flex-direction:column}
.brand{padding:13px 16px;font-weight:700;font-size:15px;border-bottom:1px solid var(--bd);background:#fff;
       display:flex;align-items:center;justify-content:space-between;gap:8px}
.brand small{font-weight:400;color:var(--mut);font-size:12px}
.search{padding:10px 12px;border-bottom:1px solid var(--bd);background:#fff}
.search input{width:100%;padding:7px 10px;border:1px solid var(--bd);border-radius:6px;font-size:13px;outline:none}
.search input:focus{border-color:var(--acc)}
#tree{flex:1;overflow:auto;padding:8px 6px 40px}
.node{display:flex;align-items:flex-start;gap:6px;padding:5px 8px;border-radius:6px;cursor:pointer;
      font-size:13px;line-height:1.5;color:#333}
.node:hover{background:#eceef1}
.node.active{background:#e0ecff;color:#0a4fd0;font-weight:600}
.node .dot{flex:0 0 auto;color:#c9cdd4;font-size:11px;margin-top:2px}
.node.active .dot{color:var(--acc)}
.node .t{flex:1;word-break:break-word}
.node.d0{font-weight:700}.node.d1{font-weight:600}
.hit{background:#fff3b0;border-radius:2px}
main{flex:1;display:flex;flex-direction:column;min-width:0}
#bar{display:flex;align-items:center;gap:10px;padding:8px 14px;border-bottom:1px solid var(--bd);background:#fff}
#bar button{border:1px solid var(--bd);background:#fff;border-radius:6px;padding:5px 12px;cursor:pointer;font-size:13px}
#bar button:hover:not(:disabled){border-color:var(--acc);color:var(--acc)}
#bar button:disabled{opacity:.4;cursor:default}
#cur{flex:1;text-align:center;font-size:13px;color:var(--mut);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
#cur b{color:var(--fg)}
iframe{flex:1;border:0;width:100%;background:#fff}
.empty{padding:24px;color:var(--mut);font-size:13px}
"""

VIEWER_JS = """(function () {
  var N = window.NAV, list = N.items;
  var treeEl = document.getElementById('tree'), frame = document.getElementById('frame');
  var curEl = document.getElementById('cur'), q = document.getElementById('q');
  var prev = document.getElementById('prev'), next = document.getElementById('next');
  var backBtn = document.getElementById('back');
  var current = 0, filtered = null, external = false, lastBody = 0;
  var loadedUrl = list.length ? list[0].u : '';   // iframe 当前实际加载的地址
  var keyIndex = {};
  var i;

  // 页面清单，父页面用它校验来自 iframe 的 key 是否属于本次导出
  for (i = 0; i < list.length; i++) {
    list[i].i = i;
    if (!(list[i].u in keyIndex)) { keyIndex[list[i].u] = i; }
  }

  function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  // 阅读器地址里的状态：正文页 #p=<序号>[&q=<query>][&a=<锚点>]，说明页 #ext=<原目标>
  function parseHash(h) {
    h = String(h || '').replace(/^#/, '');
    if (!h) { return null; }
    var m = /(?:^|&)p=(\\d+)/.exec(h);
    if (m) {
      var am = /(?:^|&)a=([^&]*)/.exec(h);
      var qm = /(?:^|&)q=([^&]*)/.exec(h);
      return { kind: 'p', idx: +m[1],
               q: qm ? decodeURIComponent(qm[1]) : '',
               a: am ? decodeURIComponent(am[1]) : '' };
    }
    var em = /^ext=([\\s\\S]*)$/.exec(h);
    if (em) { return { kind: 'ext', payload: decodeURIComponent(em[1]) }; }
    return null;
  }

  // 正文页要加载的完整地址：路径 + ?query + #锚点，三者都不能丢
  function bodyUrl(st) {
    return list[st.idx].u + (st.q ? '?' + st.q : '') + (st.a ? '#' + st.a : '');
  }

  function render(items, isSearch) {
    treeEl.innerHTML = '';
    if (!items.length) { treeEl.innerHTML = '<div class="empty">没有匹配的目录项</div>'; return; }
    items.forEach(function(it) {
      var d = document.createElement('div');
      // 显示章节外说明页时不点亮任何正文项，避免被误认为正在读某一页
      d.className = 'node d' + Math.min(it.d, 3) +
        (!external && it.i === current ? ' active' : '');
      d.dataset.idx = it.i;
      var t = document.createElement('span'); t.className = 't'; t.textContent = it.t;
      if (isSearch && q.value) {
        var k = it.t.toLowerCase().indexOf(q.value.toLowerCase());
        if (k >= 0) {
          t.innerHTML = '';
          t.appendChild(document.createTextNode(it.t.slice(0, k)));
          var m = document.createElement('span'); m.className = 'hit';
          m.textContent = it.t.slice(k, k + q.value.length);
          t.appendChild(m); t.appendChild(document.createTextNode(it.t.slice(k + q.value.length)));
        }
      }
      var dot = document.createElement('span'); dot.className = 'dot'; dot.textContent = '\\u25CF';
      d.appendChild(dot); d.appendChild(t);
      if (!isSearch) d.style.paddingLeft = (8 + it.d * 14) + 'px';
      d.title = it.t;
      d.onclick = function() { navigate(+this.dataset.idx, '', ''); };
      treeEl.appendChild(d);
    });
    var a = treeEl.querySelector('.node.active');
    if (a) a.scrollIntoView({block: 'nearest'});
  }

  function paint() {
    if (external) {
      curEl.innerHTML = '<b>章节外内容</b> &nbsp;本次样例未包含';
    } else {
      var it = list[current];
      curEl.innerHTML = '<b>' + esc(it.t) + '</b> &nbsp;(' + (current + 1) + '/' + list.length + ')';
    }
    prev.disabled = external ? (lastBody <= 0) : (current === 0);
    next.disabled = external ? (lastBody >= list.length - 1) : (current === list.length - 1);
    if (backBtn) { backBtn.style.display = external ? '' : 'none'; }
    render(filtered || list, !!filtered);
  }

  function setFrame(url) {
    if (loadedUrl === url) { return; }
    loadedUrl = url;
    frame.src = url;
  }

  // 只有这里真正加载页面；所有入口都先改 hash，再由 hashchange 统一加载。
  function load(st) {
    if (!st) { return; }
    if (st.kind === 'p') {
      if (st.idx < 0 || st.idx >= list.length) { return; }
      external = false; current = st.idx; lastBody = st.idx;
      setFrame(bodyUrl(st));
    } else {
      external = true;
      setFrame('viewer/external.html#' + st.payload);
    }
    paint();
  }

  // 阅读器自己的地址始终只描述“现在看的是哪一页”，不新增历史记录：
  // 前进/后退由浏览器的正文加载序列驱动，落定后再反向同步地址。
  function replaceHash(h) {
    try { location.replace(location.href.split('#')[0] + h); }
    catch (e) { location.hash = h; }
  }
  function goHash(h) {
    if (location.hash === h) { load(parseHash(h)); } else { replaceHash(h); }
  }
  function navigate(idx, q, a) {
    goHash(navHash(idx, q, a));
  }
  function navigateExternal(payload) {
    goHash('#ext=' + encodeURIComponent(payload));
  }
  function navHash(idx, q, a) {
    return '#p=' + idx + (q ? '&q=' + encodeURIComponent(q) : '')
                     + (a ? '&a=' + encodeURIComponent(a) : '');
  }
  function greet() {
    try { frame.contentWindow.postMessage({ __hdx: 1, type: 'hello' }, '*'); }
    catch (e) { /* 正文尚未加载时忽略，load 事件后还会再来一次 */ }
  }
  frame.addEventListener('load', greet);

  window.addEventListener('hashchange', function () {
    load(parseHash(location.hash));
  });

  // 正文内的点击由页面发消息上来，父页面统一导航：
  // 这样 iframe 不会自己产生历史记录，浏览器前进/后退只对应父页面的一步。
  window.addEventListener('message', function (e) {
    if (e.source !== frame.contentWindow) { return; }   // 只接受本 iframe 的消息
    var d = e.data;
    if (!d || d.__hdx !== 1) { return; }
    if (d.type === 'hi') {
      greet();                                          // 正文页询问：回 hello 建立通信
    } else if (d.type === 'link') {
      var idx = keyIndex[d.k];
      if (idx === undefined) { return; }                // 不属于导出清单，忽略
      navigate(idx, d.q || '', d.a || '');
    } else if (d.type === 'extlink') {
      navigateExternal(d.payload || '');
    } else if (d.type === 'external') {
      // 说明页上报自身：正常情况下阅读器已经知道，这里只兜底。
      if (!external) { navigateExternal(d.payload || ''); }
    } else if (d.type === 'page') {
      var k = keyIndex[d.k];
      if (k === undefined) { return; }                  // 不属于导出清单，忽略
      // 正文由浏览器前进/后退带动时，靠这条上报把阅读器地址对齐到正文真实地址：
      // ?query 与 #锚点都要一起带上，否则后退/前进回来会丢掉锚点。
      var q = (d.search || '').replace(/^\\?/, '');
      var a = (d.hash || '').replace(/^#/, '');
      loadedUrl = list[k].u + (q ? '?' + q : '') + (a ? '#' + a : '');
      external = false; current = k; lastBody = k;
      // 地址已经一致就什么都不做：不产生新的 hashchange，也就不会重复加载、
      // 不会把正文的滚动位置重置掉（loadedUrl 已对齐，即便触发也不会重载）。
      var want = navHash(k, q, a);
      if (location.hash !== want) { replaceHash(want); }
      paint();
    }
  });

  q.oninput = function () {
    var v = q.value.trim().toLowerCase();
    if (!v) { filtered = null; render(list, false); return; }
    filtered = list.filter(function (it) { return it.t.toLowerCase().indexOf(v) >= 0; });
    render(filtered, true);
  };
  q.onkeydown = function (e) {
    if (e.key === 'Enter' && filtered && filtered.length) { navigate(filtered[0].i, '', ''); }
    if (e.key === 'Escape') { q.value = ''; q.oninput(); }
  };
  prev.onclick = function () { navigate(external ? lastBody - 1 : current - 1, '', ''); };
  next.onclick = function () { navigate(external ? lastBody + 1 : current + 1, '', ''); };
  if (backBtn) { backBtn.onclick = function () { navigate(lastBody, '', ''); }; }
  document.addEventListener('keydown', function (e) {
    if (e.target === q) { return; }
    if (e.key === 'ArrowLeft') { prev.onclick(); }
    if (e.key === 'ArrowRight') { next.onclick(); }
  });

  var st = parseHash(location.hash);
  if (st) {
    load(st);
  } else {
    current = 0; lastBody = 0; paint();
    try { location.replace(location.href.split('#')[0] + '#p=0'); } catch (e) {}
  }
})();
"""


def build_viewer(out_dir, result):
    title = result["chapter_path"].split(">")[-1].strip()
    items = [{"t": p["txt"], "d": p["depth"], "u": p["zip"]} for p in result["pages"]]
    nav_json = json.dumps({"title": title, "items": items}, ensure_ascii=False)
    nav_json = nav_json.replace("</", "<\\/")

    html = f"""<!DOCTYPE html>
<html lang="zh-cn">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{title}（离线阅读）</title>
<link rel="stylesheet" href="viewer/viewer.css">
</head>
<body>
<div id="app">
  <aside id="side">
    <div class="brand"><span>{title}</span><small>{result['node_count']} 项</small></div>
    <div class="search"><input id="q" type="search" placeholder="搜索目录标题（回车跳转 / Esc 清空）"></div>
    <div id="tree"></div>
  </aside>
  <main>
    <header id="bar">
      <button id="prev">&#8249; 上一项</button>
      <button id="back" style="display:none">&#8592; 返回正文</button>
      <span id="cur"></span>
      <button id="next">下一项 &#8250;</button>
    </header>
    <iframe id="frame" src="{items[0]['u']}" title="正文"></iframe>
  </main>
</div>
<script>window.NAV = {nav_json};</script>
<script src="viewer/viewer.js"></script>
</body>
</html>
"""
    write_out(out_dir, "index.html", html.encode("utf-8"))
    write_out(out_dir, "viewer/viewer.css", VIEWER_CSS.encode("utf-8"))
    write_out(out_dir, "viewer/viewer.js", VIEWER_JS.encode("utf-8"))
    write_out(out_dir, "viewer/sync.js", SYNC_JS.encode("utf-8"))


def write_readme(out_dir, result):
    title = result["chapter_path"].split(">")[-1].strip()
    lines = [
        f"{title} —— 离线阅读包",
        "=" * 40,
        "",
        "【怎么用】",
        "  1. 双击本目录下的 index.html，用 Chrome / Edge / Firefox 打开即可。",
        "     不需要 ICS Lite，也不需要本地服务器；目录数据已内嵌在页面里，",
        "     不受 file:// 下 fetch 限制。",
        "  2. 左侧是原章节目录（层级与顺序与 navy.xml 一致），右侧是正文。",
        "  3. 顶部“上一项 / 下一项”按目录顺序翻页；键盘 ← → 同样可用。",
        "  4. 搜索框输入标题关键字可过滤目录，回车跳到第一条。",
        "  5. 正文里的章节内链接可点击，左侧高亮、顶部标题与页码、上一项/下一项",
        "     都会同步；浏览器前进/后退也始终与正文保持一致。",
        "  6. 指向本次未导出章节的链接会打开一个说明页，保留原始目标路径，",
        "     不会伪造成正文；说明页顶部有“← 返回正文”可回到刚才的位置。",
        "  7. resources/ 下的正文页也可以单独打开（双击、新标签页、从阅读器里",
        "     “在新标签页打开”都行）：此时不经过阅读器，链接按浏览器原生方式跳转，",
        "     路径、?query 与 #锚点都保持原样。",
        "",
        "【目录结构】",
        "  index.html        阅读入口（双击这个）",
        "  viewer/           阅读器自身的 css / js 与章节外链接说明页",
        "  resources/        正文 HTML 与图片、样式、脚本等依赖",
        "  reports/          导出清单 manifest.json 与校验报告",
        "  使用说明.txt      本文件",
        "",
        "【本次导出概况】",
        f"  目录节点数：{result['node_count']}",
        f"  唯一正文文件数：{result['unique_body_files']}",
        f"  资源文件数：{result['asset_count']}（其中图片 {result['image_count']}）",
        f"  附件数：{result['attachment_count']}",
        f"  章节外引用：{result['external_count']} 个目标",
        f"  生成时间：{result['generated_at']}",
        "",
        "【说明】",
        "  * 正文已由原 GB2312 转为 UTF-8，并同步更新了 charset 声明。",
        "  * 输出镜像原包内相对目录，整包可整体复制到别处使用，不依赖绝对路径。",
        "  * 章节外引用不代表本包已完整覆盖该内容。",
        "",
    ]
    write_out(out_dir, "使用说明.txt", "\n".join(lines).encode("utf-8"))


if __name__ == "__main__":
    sys.exit(main())
