#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Split an existing export_word DOCX by actual ZIP bytes, preserving source content.

python split_word.py --input "Word导出/用户手册全集.docx" --out-dir "Word导出/用户手册分卷" --prefix "用户手册" --max-mb 10
Only supports DOCX created by export_word.py (HDXGenerated chapter markers).
"""
from __future__ import annotations

import argparse
from copy import deepcopy
import hashlib
from io import BytesIO
import json
import math
from pathlib import Path
import posixpath
import re
import sys
import time
from urllib.parse import quote, unquote, urlsplit
from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree as E

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P = "http://schemas.openxmlformats.org/package/2006/relationships"
C = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"w": W, "r": R}
DOC = "word/document.xml"
RELS = "word/_rels/document.xml.rels"


class SplitError(RuntimeError):
    pass


def element(tag, text=None, **attrs):
    e = E.Element("{" + W + "}" + tag)
    for k, v in attrs.items(): e.set("{" + W + "}" + k, str(v))
    if text is not None: e.text = text
    return e


def paragraph(text="", style=None):
    p = element("p")
    if style:
        pr = element("pPr"); pr.append(element("pStyle", val=style)); p.append(pr)
    if text:
        r = element("r"); r.append(element("t", text)); p.append(r)
    return p


def serialize(root):
    return E.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def texts(nodes):
    return "".join(t.text or "" for n in nodes for t in n.iter("{" + W + "}t"))


def bookmark_names(nodes):
    return [b.get("{" + W + "}name") for n in nodes for b in n.iter("{" + W + "}bookmarkStart")]


def is_chapter(p):
    return bool(p.xpath('.//w:rStyle[@w:val="HDXGenerated"]', namespaces=NS))


def heading_level(p):
    st = p.find("w:pPr/w:pStyle", NS)
    m = re.fullmatch(r"Heading([1-9])", st.get("{" + W + "}val", ""), re.I) if st is not None else None
    return int(m[1]) if m else None


def rels_name(part):
    if not part: return "_rels/.rels"
    return posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")


def rel_target(part, target):
    target = unquote(target).replace("\\", "/")
    name = posixpath.normpath(target.lstrip("/")) if target.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(part), target))
    if name == ".." or name.startswith("../"):
        raise SplitError("DOCX 内部关系路径越界：" + target)
    return name


def checked_path(root, rel):
    root = Path(root).resolve(); result = (root / rel).resolve()
    if not result.is_relative_to(root): raise SplitError("附件路径越界：" + rel)
    return result


class Splitter:
    def __init__(self, source, prefix, data=None, content_start=None):
        self.source = Path(source).resolve()
        self.prefix = prefix
        with ZipFile(BytesIO(data) if data is not None else self.source) as z:
            self.parts = {n: z.read(n) for n in z.namelist() if not n.endswith("/")}
        if DOC not in self.parts or RELS not in self.parts:
            raise SplitError("不是支持的 DOCX 文件。")
        self.root = E.fromstring(self.parts[DOC])
        body = self.root.find("w:body", NS)
        if body is None: raise SplitError("DOCX 正文缺失。")
        nodes = list(body)
        self.sect = nodes.pop() if nodes and nodes[-1].tag == "{" + W + "}sectPr" else None
        start = content_start if content_start is not None else next((i for i, n in enumerate(nodes) if is_chapter(n)), None)
        if start is None:
            raise SplitError("未找到 HDXGenerated 章节标记。本工具仅拆分 export_word.py 生成的 Word。")
        self.content = nodes[start:]
        prefix_nodes = nodes[:start]
        self.before_toc, self.after_toc = self.strip_toc(prefix_nodes)
        self.chapter_cuts = [i for i, n in enumerate(self.content) if i and is_chapter(n)]
        self.heading_cuts = [i for i, n in enumerate(self.content) if i and heading_level(n)]
        self.source_rels = E.fromstring(self.parts[RELS])
        self.attachments = {}
        self.warnings = []
        self.rid_counter = 0

    def strip_toc(self, nodes):
        begin = next((i for i, n in enumerate(nodes)
                      if any(re.match(r"\s*TOC\b", t.text or "") for t in n.iter("{" + W + "}instrText"))), None)
        if begin is None: return nodes, []
        depth, end = 0, None
        for i in range(begin, len(nodes)):
            for f in nodes[i].iter("{" + W + "}fldChar"):
                kind = f.get("{" + W + "}fldCharType")
                if kind == "begin": depth += 1
                if kind == "end": depth -= 1
            if depth == 0:
                end = i; break
        if end is None: raise SplitError("原文目录域不完整，无法安全拆分。")
        title = begin - 1 if begin and texts([nodes[begin - 1]]).strip() == "目录" else begin
        return nodes[:title], nodes[end + 1:]

    def filename(self, i):
        return self.prefix + str(i + 1) + ".docx"

    def owners(self, ranges):
        owners = {}
        for i, (start, end) in enumerate(ranges):
            for name in bookmark_names(self.content[start:end]):
                if name in owners and owners[name] != i:
                    raise SplitError("源文件存在重复书签，不能唯一建立跨卷链接：" + name)
                owners[name] = i
        return owners

    def toc(self, content):
        result = [paragraph("目录", "Title")]
        p = paragraph()
        r = element("r"); r.append(element("fldChar", fldCharType="begin", dirty="true")); p.append(r)
        r = element("r"); t = element("instrText", ' TOC \\o "1-9" \\h \\z \\u ')
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve"); r.append(t); p.append(r)
        r = element("r"); r.append(element("fldChar", fldCharType="separate")); p.append(r)
        result.append(p)
        for item in content:
            level = heading_level(item)
            names = bookmark_names([item])
            if not level or not names: continue
            p = paragraph(); pr = element("pPr"); pr.append(element("ind", left=(level - 1) * 180)); p.append(pr)
            h = element("hyperlink", anchor=names[0], history="1")
            r = element("r"); r.append(element("t", texts([item]))); h.append(r); p.append(h)
            r = element("r"); r.append(element("t", "  ")); p.append(r)
            fld = element("fldSimple", instr=" PAGEREF " + names[0] + " \\h ", dirty="true")
            r = element("r"); r.append(element("t", "—")); fld.append(r); p.append(fld); result.append(p)
        p = paragraph(); r = element("r"); r.append(element("fldChar", fldCharType="end")); p.append(r); result.append(p)
        return result

    def external_files(self, document, relations):
        used = {v for e in document.iter() for k, v in e.attrib.items() if k.startswith("{" + R + "}")}
        for rel in relations:
            if rel.get("Id") not in used or rel.get("TargetMode") != "External": continue
            target = rel.get("Target", "")
            u = urlsplit(target)
            if u.scheme or u.netloc or not u.path: continue
            # Newly generated cross-volume links have no source file to copy.
            if rel.get("Id", "").startswith("splitLink"): continue
            path = unquote(u.path).replace("\\", "/")
            f = checked_path(self.source.parent, path)
            if not f.is_file():
                raise SplitError("原 Word 引用的本地附件不存在：" + str(f))
            self.attachments[path] = f

    def package(self, document, relations):
        used = {v for e in document.iter() for k, v in e.attrib.items() if k.startswith("{" + R + "}")}
        conditional = {"image", "hyperlink", "header", "footer", "oleObject", "package", "aFChunk", "chart", "diagramData", "diagramLayout", "diagramQuickStyle", "diagramColors"}
        for rel in list(relations):
            if rel.get("Type", "").rsplit("/", 1)[-1] in conditional and rel.get("Id") not in used:
                relations.remove(rel)
        overrides = {DOC: serialize(document), RELS: serialize(relations)}
        reachable = set()
        queue = [""]
        while queue:
            part = queue.pop()
            if part:
                if part in reachable: continue
                if part not in self.parts and part not in overrides:
                    raise SplitError("DOCX 内部引用缺失：" + part)
                reachable.add(part)
            relname = rels_name(part)
            data = overrides.get(relname, self.parts.get(relname))
            if data is None: continue
            reachable.add(relname)
            for rel in E.fromstring(data):
                if rel.get("TargetMode") == "External": continue
                queue.append(rel_target(part, rel.get("Target", "")))
        types = E.fromstring(self.parts["[Content_Types].xml"])
        for item in list(types):
            if item.tag == "{" + C + "}Override" and unquote(item.get("PartName", "")).lstrip("/") not in reachable:
                types.remove(item)
        overrides["[Content_Types].xml"] = serialize(types)
        reachable.add("[Content_Types].xml")
        out = BytesIO()
        with ZipFile(out, "w", compression=ZIP_DEFLATED) as z:
            for name in sorted(reachable): z.writestr(name, overrides.get(name, self.parts.get(name)))
        return out.getvalue()

    def render(self, ranges, i, owners):
        start, end = ranges[i]
        content = [deepcopy(n) for n in self.content[start:end]]
        root = E.Element(self.root.tag, attrib=dict(self.root.attrib), nsmap=self.root.nsmap)
        body = E.SubElement(root, "{" + W + "}body")
        for n in self.before_toc: body.append(deepcopy(n))
        body.append(paragraph(f"分卷 {i + 1} / {len(ranges)}", "Subtitle"))
        for n in self.toc(content): body.append(n)
        if self.after_toc:
            for n in self.after_toc: body.append(deepcopy(n))
        else:
            p = paragraph(); r = element("r"); r.append(element("br", type="page")); p.append(r); body.append(p)
        for n in content: body.append(n)
        sect = self.section_at(end)
        if sect is not None: body.append(deepcopy(sect))
        relations = deepcopy(self.source_rels)
        local = set(bookmark_names(list(body)))
        broken, cross = [], []
        relation_ids = {r.get("Id") for r in relations}
        for link in root.iter("{" + W + "}hyperlink"):
            anchor = link.get("{" + W + "}anchor")
            if not anchor or anchor in local: continue
            if link.get("{" + R + "}id"): continue  # Existing external-document bookmark.
            owner = owners.get(anchor)
            if owner is None:
                broken.append(anchor); continue
            self.rid_counter += 1
            rid = "splitLink" + str(self.rid_counter)
            while rid in relation_ids:
                self.rid_counter += 1; rid = "splitLink" + str(self.rid_counter)
            rel = E.SubElement(relations, "{" + P + "}Relationship")
            rel.set("Id", rid); rel.set("Type", R + "/hyperlink")
            rel.set("Target", quote(self.filename(owner), safe="")); rel.set("TargetMode", "External")
            link.set("{" + R + "}id", rid)
            cross.append({"target_file": self.filename(owner), "bookmark": anchor})
        if broken: raise SplitError("存在无法定位的内部书签：" + ", ".join(broken[:5]))
        self.external_files(root, relations)
        if texts(content) != texts(self.content[start:end]):
            raise SplitError("拆分后正文文本发生改变，已停止。")
        # Images are copied as original package parts; only unreferenced parts are pruned.
        data = self.package(root, relations)
        with ZipFile(BytesIO(data)) as z:
            if z.testzip() is not None: raise SplitError("生成的分卷压缩结构损坏。")
            for name in z.namelist():
                if name.startswith("word/media/") and z.read(name) != self.parts[name]:
                    raise SplitError("分卷中的图片字节发生改变。")
        titles = [texts([n]) for n in content if is_chapter(n)]
        report = {"file": self.filename(i), "source_block_start": start, "source_block_end": end,
                  "bytes": len(data), "first_chapter": titles[0] if titles else "前一章节续篇",
                  "last_chapter": titles[-1] if titles else "前一章节续篇",
                  "text_intact": True, "images_intact": True, "cross_volume_links": cross}
        return data, report

    def section_at(self, end):
        # A sectPr describes the preceding content, including a volume ending mid-section.
        for node in self.content[max(0, end - 1):]:
            sec = node.find('w:pPr/w:sectPr', NS)
            if sec is not None: return sec
        return self.sect

    def safe_cuts(self):
        bookmarks = set(); depth = 0; result = set()
        for i, node in enumerate(self.content):
            for e in node.iter():
                if e.tag == '{' + W + '}bookmarkStart': bookmarks.add(e.get('{' + W + '}id'))
                if e.tag == '{' + W + '}bookmarkEnd': bookmarks.discard(e.get('{' + W + '}id'))
                if e.tag == '{' + W + '}fldChar':
                    kind = e.get('{' + W + '}fldCharType')
                    if kind == 'begin': depth += 1
                    elif kind == 'end': depth -= 1
            if not bookmarks and depth == 0: result.add(i + 1)
        return result

    def cut(self, start, end):
        if end - start <= 1:
            raise SplitError("单个段落、图片或完整表格超过体积限制，不能无损切开。请提高上限或单独处理该内容；未输出超限分卷。")
        mid = (start + end) / 2
        safe = self.safe_cuts()
        choices = [v for v in self.chapter_cuts if start < v < end and v in safe]
        if not choices: choices = [v for v in self.heading_cuts if start < v < end and v in safe]
        if not choices: choices = [v for v in safe if start < v < end]
        if not choices: raise SplitError('内容包含跨段书签或域，无法在上限内完整分卷；未截断或丢弃内容。')
        return min(choices, key=lambda x: abs(x - mid))


def split_docx(source, out_dir, prefix="用户手册", max_mb=10, allow_large_attachments=False, reserve_percent=5):
    start_time = time.perf_counter()
    source, out_dir = Path(source).resolve(), Path(out_dir).resolve()
    if not source.is_file(): raise SplitError("找不到输入 Word：" + str(source))
    if not math.isfinite(max_mb) or max_mb <= 0: raise SplitError("--max-mb 必须为正数。")
    if not 0 <= reserve_percent <= 30: raise SplitError("预留比例必须在 0 到 30 之间。")
    if not prefix or re.search(r'[<>:"/\\|?*]', prefix) or prefix.endswith((" ", ".")):
        raise SplitError("文件名前缀包含不支持的字符。")
    if out_dir.exists() and (not out_dir.is_dir() or any(out_dir.iterdir())):
        raise SplitError("输出目录必须不存在或为空，请换一个新文件夹：" + str(out_dir))
    limit = int(max_mb * 1_000_000)  # Decimal MB: stricter than Windows' displayed MiB.
    target = int(limit * (1 - reserve_percent / 100))
    s = Splitter(source, prefix)
    s.external_files(s.root, s.source_rels)
    oversized_assets = [{"file": n, "bytes": p.stat().st_size} for n, p in s.attachments.items() if p.stat().st_size > limit]
    if oversized_assets and not allow_large_attachments:
        details = "；".join(f"{r['file']}（{r['bytes']:,} 字节）" for r in oversized_assets)
        raise SplitError("以下原始附件超过上限：" + details +
                         "。尚未执行分卷；只有确认 10 MB 限制不包括原始附件时，才加 --allow-large-attachments。")
    ranges = [(0, len(s.content))]
    print(f"输入：{source}\n每份上限：{limit:,} 字节；生成目标：{target:,} 字节（预留 {reserve_percent:g}%）", flush=True)
    for iteration in range(30):
        owners = s.owners(ranges)
        outputs = []
        new_ranges = []
        oversized = False
        for i, (start, end) in enumerate(ranges):
            data, report = s.render(ranges, i, owners)
            if len(data) > target:
                cut = s.cut(start, end)
                new_ranges.extend([(start, cut), (cut, end)])
                oversized = True
            else:
                new_ranges.append((start, end)); outputs.append((data, report))
        if not oversized: break
        ranges = new_ranges
        print(f"体积检查后细分为 {len(ranges)} 份，重新核对跨卷链接和实际文件大小……", flush=True)
    else: raise SplitError("分卷次数超过限制，未输出文件。")
    if any(len(data) > limit for data, _ in outputs): raise SplitError("仍有超限文件，未输出。")
    oversized_assets = [{"file": n, "bytes": p.stat().st_size} for n, p in s.attachments.items() if p.stat().st_size > limit]
    if oversized_assets and not allow_large_attachments:
        details = "；".join(f"{r['file']}（{r['bytes']:,} 字节）" for r in oversized_assets)
        raise SplitError("Word 分卷可满足限制，但以下原始附件仍超过上限：" + details +
                         "。尚未写出分卷；只有确认 10 MB 限制不包括原始附件时，才加 --allow-large-attachments。")
    # Check all output destinations before writing, including attachment paths.
    filenames = {report["file"] for _, report in outputs}
    reserved = filenames | {"分卷报告.json", "阅读说明.txt"}
    if any(name.casefold() in {n.casefold() for n in reserved} for name in s.attachments):
        raise SplitError("附件名称与分卷文件冲突。")
    for name in s.attachments: checked_path(out_dir, name)
    out_dir.mkdir(parents=True, exist_ok=True)
    attachment_reports = []
    for name, source_file in s.attachments.items():
        target_file = checked_path(out_dir, name); target_file.parent.mkdir(parents=True, exist_ok=True)
        data = source_file.read_bytes(); target_file.write_bytes(data)
        digest = hashlib.sha256(data).hexdigest()
        if hashlib.sha256(target_file.read_bytes()).hexdigest() != digest: raise SplitError("附件复制校验失败：" + name)
        attachment_reports.append({"file": name, "bytes": len(data), "sha256": digest})
    for data, report in outputs:
        p = out_dir / report["file"]; p.write_bytes(data)
        if p.stat().st_size > limit or p.read_bytes() != data: raise SplitError("分卷写入后的校验失败：" + p.name)
        print(f"{p.name}：{p.stat().st_size / 1_000_000:.2f} MB", flush=True)
    report = {"source": str(source), "limit_bytes": limit, "target_bytes": target,
              "parts": [r for _, r in outputs], "attachments": attachment_reports,
              "oversized_attachments": oversized_assets,
              "all_docx_within_limit": True, "all_delivered_files_within_limit": not oversized_assets,
              "source_body_blocks": len(s.content), "coverage_contiguous": ranges[0][0] == 0 and ranges[-1][1] == len(s.content) and all(a[1] == b[0] for a, b in zip(ranges, ranges[1:])),
              "source_body_text_sha256": hashlib.sha256(texts(s.content).encode()).hexdigest(),
              "elapsed_seconds": round(time.perf_counter() - start_time, 3),
              "visual_review": "保留原正文排版；拆分后的分页和跨文件链接仍需在 Word 中实测"}
    (out_dir / "分卷报告.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    instructions = ("按文件名编号顺序阅读。所有 Word 和附件文件夹应保持在当前目录结构内。\n"
                    "每份 Word 都按实际保存字节检查，10 MB 按 10,000,000 字节计算；默认额外预留 5%。\n"
                    "请在 Word 中更新目录（Ctrl+A → F9），更新保存后的文件大小需重新检查。\n"
                    "目录只包含本分卷的标题，章节编号沿用原手册；章内引用跨卷时链接到相应 Word。\n"
                    "可能需要 Ctrl+单击跟随链接；Word 的安全设置可能询问是否打开另一文件。\n"
                    "正文文字和图片已做自动完整性检查；最终分页及跨文件链接请在 Word 中验收。\n")
    if oversized_assets:
        instructions += "注意：已明确允许原始附件超过上限，不能声称交付中的每个文件都小于 10 MB。\n"
        instructions += "\n".join(f"{v['file']}：{v['bytes'] / 1_000_000:.2f} MB" for v in oversized_assets)
    (out_dir / "阅读说明.txt").write_text(instructions, encoding="utf-8")
    print(f"完成：{len(outputs)} 份 Word；耗时 {report['elapsed_seconds']} 秒。输出：{out_dir}")
    return report


def main():
    if hasattr(sys.stdout, "reconfigure"): sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description="按实际体积无损拆分 HDX 工具生成的 Word，保留图片并重建跨卷链接")
    ap.add_argument("--input", required=True, type=Path)
    ap.add_argument("--out-dir", required=True, type=Path, help="不存在或为空的新目录")
    ap.add_argument("--prefix", default="用户手册")
    ap.add_argument("--max-mb", type=float, default=10, help="单份 Word 上限，按十进制 MB 计算，默认 10")
    ap.add_argument("--reserve-percent", type=float, default=5)
    ap.add_argument("--allow-large-attachments", action="store_true", help="明确允许原始附件超过体积上限")
    args = ap.parse_args()
    try:
        split_docx(args.input, args.out_dir, args.prefix, args.max_mb, args.allow_large_attachments, args.reserve_percent)
        return 0
    except (SplitError, OSError, ValueError, E.XMLSyntaxError) as exc:
        print("错误：" + str(exc), file=sys.stderr); return 2


if __name__ == "__main__":
    raise SystemExit(main())
