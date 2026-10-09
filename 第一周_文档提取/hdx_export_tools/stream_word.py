"""Generate bounded Word volumes as HDX pages are rendered, never a whole-book DOCX.

Native Word attachments keep their own styles/package and start a new volume.
Only small staged volumes are finalized for forward references after generation.
"""
from copy import deepcopy
from io import BytesIO
import hashlib
import json
import math
import os
from pathlib import Path
import re
import subprocess
import tempfile
import time
from zipfile import ZipFile

from lxml import etree as E
from export_word import Package, PackageError, WordBuilder, select_nodes, bookmark_key, local_path, VERSION
from split_word import Splitter, SplitError, W, R, NS, DOC, RELS, texts, bookmark_names, paragraph, element, serialize


def convert_attachment(data, name, work):
    """Lossless package passthrough for DOCX; Word native file conversion for DOC."""
    suffix = Path(name).suffix.lower()
    if suffix == '.docx':
        return data, {'converter': 'DOCX 原始包', 'conversion_required': False}
    if suffix != '.doc':
        raise PackageError('目前正文合并支持 DOC/DOCX；此文件需单独转换：' + name)
    if os.name != 'nt':
        raise PackageError('旧版 DOC 自动转换需要 Windows 上的 Microsoft Word：' + name)
    import winreg
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'Word.Application\CLSID'):
            pass
    except OSError as exc:
        raise PackageError('此电脑未注册 Microsoft Word。请安装 Word，或先将 DOC 另存为 DOCX：' + name) from exc
    source = work / 'original.doc'; dest = work / 'converted.docx'; check = work / 'conversion.json'
    source.write_bytes(data)
    print('正在用 Microsoft Word 转换原始 DOC 并核对正文及对象数量：' + name, flush=True)
    result = subprocess.run(['powershell.exe', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
                             '-File', str(Path(__file__).with_name('convert_doc.ps1')),
                             '-Source', str(source), '-Destination', str(dest), '-Report', str(check)],
                            capture_output=True)
    if result.returncode or not dest.is_file() or not check.is_file():
        detail = result.stderr.decode('utf-8', errors='backslashreplace')[-2000:]
        raise PackageError('原始 DOC 转换失败，未将附件内容标记为已合并：' + name + '\n' + detail)
    report = json.loads(check.read_text(encoding='utf-8-sig'))
    if not report.get('main_story_text_intact') or not report.get('object_counts_intact'):
        raise PackageError('DOC 转换后的完整性检查未通过：' + name)
    return dest.read_bytes(), report


class VolumeWriter:
    def __init__(self, stage, prefix, max_mb):
        self.stage, self.prefix = Path(stage), prefix
        self.limit = int(max_mb * 1_000_000)
        self.target = int(self.limit * .95)
        self.shards = []
        self.pending_merge = []
        self.source_hash = hashlib.sha256()
        self.source_blocks = 0

    def snapshot(self, builder):
        buf = BytesIO(); builder.doc.save(buf)
        return Splitter(self.stage / 'logical.docx', self.prefix, buf.getvalue(), builder.stream_prefix_count)

    def raw(self, s, start, end):
        root = deepcopy(s.root); body = root.find('w:body', NS)
        for n in list(body): body.remove(n)
        for n in s.before_toc + s.after_toc: body.append(deepcopy(n))
        prefix_count = len(body)
        for n in s.content[start:end]: body.append(deepcopy(n))
        sect = s.section_at(end)
        if sect is not None: body.append(deepcopy(sect))
        return s.package(root, deepcopy(s.source_rels)), prefix_count

    def stage_content(self, s):
        ranges = [(0, len(s.content))]
        while ranges:
            a, b = ranges.pop(0)
            if a == b: continue
            data, prefix_count = self.raw(s, a, b)
            # Budget includes the actual local TOC, not merely body/media estimates.
            with ZipFile(BytesIO(data)) as z: root = E.fromstring(z.read(DOC))
            body = root.find('w:body', NS)
            for n in reversed(s.toc(s.content[a:b])): body.insert(prefix_count, n)
            measured = len(s.package(root, deepcopy(s.source_rels)))
            if measured > self.target:
                cut = s.cut(a, b); ranges[0:0] = [(a, cut), (cut, b)]
                continue
            path = self.stage / ('volume-' + str(len(self.shards) + 1) + '.docx')
            path.write_bytes(data)
            self.shards.append({'path': path, 'prefix_count': prefix_count})
            self.source_hash.update(texts(s.content[a:b]).encode())
            self.source_blocks += b - a
            print(f'已暂存分卷 {len(self.shards)}：{measured / 1_000_000:.2f} MB', flush=True)

    def page(self, builder, final=False):
        body = builder.doc.element.body
        count = builder.stream_prefix_count
        if len(body) - 1 <= count: return
        s = self.snapshot(builder)
        data, _ = self.raw(s, 0, len(s.content))
        if not final and len(data) < self.target:
            return
        self.stage_content(s)
        for n in list(body)[count:-1]: body.remove(n)
        # Drop relationships to already-staged media so the live package stays bounded.
        used = {v for e in builder.doc.element.iter() for k, v in e.attrib.items() if k.startswith('{' + R + '}')}
        for rid, rel in list(builder.doc.part.rels.items()):
            if rel.reltype.rsplit('/', 1)[-1] in {'image', 'hyperlink'} and rid not in used:
                del builder.doc.part.rels[rid]

    def attachment(self, builder, page):
        self.page(builder, final=True)
        with tempfile.TemporaryDirectory(dir=self.stage, prefix='convert-') as work:
            original = builder.pkg.read(page['path'])
            data, conversion = convert_attachment(original, page['path'], Path(work))
        s = Splitter(self.stage / 'native.docx', self.prefix, data, 0)
        # Local references to another file would become broken outside the source package.
        # Fail closed; don't discard linked images, include-text fields, or external objects.
        for part_name, payload in s.parts.items():
            if part_name.endswith('.rels'):
                for rel in E.fromstring(payload):
                    if rel.get('TargetMode') == 'External':
                        is_web_link = rel.get('Type', '').endswith('/hyperlink') and re.match(r'^(https?|mailto):', rel.get('Target', ''), re.I)
                        if not is_web_link:
                            raise PackageError('附件含依赖外部文件的关系，需先在 Word 中嵌入后再导出：' + page['path'])
            if part_name.startswith('word/') and part_name.endswith('.xml'):
                root = E.fromstring(payload)
                instructions = [e.text or '' for e in root.iter('{' + W + '}instrText')]
                instructions += [e.get('{' + W + '}instr', '') for e in root.iter('{' + W + '}fldSimple')]
                if any(re.match(r'\s*(INCLUDETEXT|INCLUDEPICTURE|LINK)\b', v) for v in instructions):
                    raise PackageError('附件含外部内容域，请先在 Word 中嵌入为本地内容：' + page['path'])
        original_text = texts(s.content)
        # Materialize inherited header/footer references before a mid-section boundary.
        inherited = {}
        sections = [sec for n in s.content for sec in n.findall('w:pPr/w:sectPr', NS)]
        if s.sect is not None: sections.append(s.sect)
        for sec in sections:
            for tag in ('headerReference', 'footerReference'):
                for e in sec.findall('w:' + tag, NS): inherited[(tag, e.get('{' + W + '}type'))] = deepcopy(e)
            present = {(e.tag.rsplit('}', 1)[-1], e.get('{' + W + '}type')) for e in sec
                       if e.tag in {'{' + W + '}headerReference', '{' + W + '}footerReference'}}
            for key, e in inherited.items():
                if key not in present: sec.insert(0, deepcopy(e))
        # Preserve original native styles, numbering, images, notes, headers and footers.
        # Namespace bookmark names so unrelated attachments cannot collide.
        namespace = 'a' + hashlib.sha256(page['path'].encode()).hexdigest()[:12] + '_'
        names = {n: namespace + str(i) for i, n in enumerate(bookmark_names(s.content))}
        for n in s.content:
            for e in n.iter():
                if e.tag == '{' + W + '}bookmarkStart':
                    e.set('{' + W + '}name', names[e.get('{' + W + '}name')])
                anchor = e.get('{' + W + '}anchor')
                if anchor in names: e.set('{' + W + '}anchor', names[anchor])
                if e.tag == '{' + W + '}instrText' and e.text:
                    for old, new in names.items():
                        e.text = re.sub(r'(?<![\w])' + re.escape(old) + r'(?![\w])', new, e.text)
                instr = e.get('{' + W + '}instr')
                if instr:
                    for old, new in names.items():
                        instr = re.sub(r'(?<![\w])' + re.escape(old) + r'(?![\w])', new, instr)
                    e.set('{' + W + '}instr', instr)
        heading = paragraph(page['number'] + ' ' + page['title'])
        pr = element('pPr'); pr.append(element('outlineLvl', val='0')); heading.insert(0, pr)
        all_ids = [int(e.get('{' + W + '}id')) for n in s.content for e in n.iter('{' + W + '}bookmarkStart')]
        new_id = max(all_ids, default=0) + 1
        heading.insert(1, element('bookmarkStart', id=new_id, name=bookmark_key(page['path'])))
        heading.append(element('bookmarkEnd', id=new_id))
        s.content.insert(0, heading)
        builder.bookmarks.add(bookmark_key(page['path']))
        # Existing attachment TOC is part of its source content; do not strip it.
        s.before_toc = []; s.after_toc = []
        s.chapter_cuts = []; s.heading_cuts = []
        # Native REF/PAGEREF/TOC fields can rely on bookmarks in other volumes.
        # Keep the original cached results locked rather than let Word replace them
        # with 'reference not found'. Ordinary page fields remain updateable.
        field_stack = []
        for n in s.content:
            for e in n.iter():
                if e.tag == '{' + W + '}fldChar':
                    kind = e.get('{' + W + '}fldCharType')
                    if kind == 'begin': field_stack.append(e)
                    if kind == 'end' and field_stack: field_stack.pop()
                if e.tag == '{' + W + '}instrText' and re.match(r'\s*(TOC|REF|PAGEREF)\b', e.text or '') and field_stack:
                    field_stack[-1].set('{' + W + '}fldLock', 'true')
                if e.tag == '{' + W + '}fldSimple' and re.match(r'\s*(TOC|REF|PAGEREF)\b', e.get('{' + W + '}instr', '')):
                    e.set('{' + W + '}fldLock', 'true')
        # Header/footer/notes can also reference a body bookmark. Preserve their
        # cached text and update the names there, without modifying image bytes.
        for part_name, payload in list(s.parts.items()):
            if not part_name.startswith('word/') or not part_name.endswith('.xml') or part_name == DOC:
                continue
            root = E.fromstring(payload); changed = False; stack = []
            for e in root.iter():
                if e.tag == '{' + W + '}fldChar':
                    kind = e.get('{' + W + '}fldCharType')
                    if kind == 'begin': stack.append(e)
                    if kind == 'end' and stack: stack.pop()
                anchor = e.get('{' + W + '}anchor')
                if anchor in names: e.set('{' + W + '}anchor', names[anchor]); changed = True
                if e.tag == '{' + W + '}instrText' and e.text:
                    old_text = e.text
                    for old, new in names.items():
                        e.text = re.sub(r'(?<![\w])' + re.escape(old) + r'(?![\w])', new, e.text)
                    if old_text != e.text:
                        changed = True
                        if stack: stack[-1].set('{' + W + '}fldLock', 'true')
                instr = e.get('{' + W + '}instr')
                if instr:
                    original_instr = instr
                    for old, new in names.items():
                        instr = re.sub(r'(?<![\w])' + re.escape(old) + r'(?![\w])', new, instr)
                    if instr != original_instr:
                        e.set('{' + W + '}instr', instr); e.set('{' + W + '}fldLock', 'true'); changed = True
            if changed: s.parts[part_name] = serialize(root)
        if texts(s.content[1:]) != original_text:
            raise PackageError('附件正文文本核对失败：' + page['path'])
        self.stage_content(s)
        return {k: page[k] for k in ('title', 'path', 'number', 'kind', 'encoding')} | {
            'body_merged': True, 'attachment_sha256': hashlib.sha256(original).hexdigest(),
            'converted_docx_sha256': hashlib.sha256(data).hexdigest(), 'conversion': conversion,
            'converted_body_text_sha256': hashlib.sha256(original_text.encode()).hexdigest(),
            'native_package_preserved': True, 'native_reference_fields': '保留并锁定原缓存；原附件目录页码沿用原文件',
            'visual_review': 'pending'}

    def load(self, item):
        return Splitter(item['path'], self.prefix, content_start=item['prefix_count'])

    def finish(self, builder):
        # Only volume files are reread here; no whole-book Word ever exists.
        owners = {}
        for i, item in enumerate(self.shards):
            s = self.load(item)
            for name in bookmark_names(s.content):
                if name in owners: raise SplitError('重复书签：' + name)
                owners[name] = i
        outputs = []; output_hash = hashlib.sha256(); blocks = 0
        for i, item in enumerate(self.shards):
            s = self.load(item)
            ranges = [(0, 0)] * len(self.shards); ranges[i] = (0, len(s.content))
            data, info = s.render(ranges, i, owners)
            if len(data) >= self.limit:
                raise SplitError('加入目录及跨卷链接后超过上限，已停止且未发布：' + s.filename(i))
            target = self.stage / s.filename(i); target.write_bytes(data)
            output_hash.update(texts(s.content).encode()); blocks += len(s.content)
            outputs.append(info)
        if output_hash.digest() != self.source_hash.digest() or blocks != self.source_blocks:
            raise SplitError('分卷正文顺序或覆盖检查失败，未发布。')
        return outputs


def run_stream_export(hdx, chapters, out, out_dir, prefix, max_mb=10, allow_large=False,
                      dry_run=False, merge_word_attachments=False, allow_large_attachments=False):
    started = time.perf_counter(); out_dir = Path(out_dir).resolve(); out = Path(out).resolve()
    if not math.isfinite(max_mb) or max_mb <= 0: raise PackageError('--max-mb 必须为正数。')
    if not prefix or re.search(r'[<>:"/\\|?*]', prefix) or prefix.endswith((' ', '.')):
        raise PackageError('分卷前缀不是有效文件名。')
    if out_dir.exists(): raise PackageError('请指定一个尚不存在的新输出文件夹：' + str(out_dir))
    pkg = Package(str(hdx))
    try:
        _, nodes = select_nodes(pkg, chapters)
        print(f'范围：{"；".join(chapters)}；节点：{len(nodes)}\n生成过程中分卷，每份小于 {max_mb:g} MB。\n输出：{out_dir}', flush=True)
        if dry_run: return {'node_count': len(nodes), 'output': str(out_dir), 'mode': 'streaming'}
        if len(nodes) > 200 and not allow_large:
            raise PackageError(f'本次包含 {len(nodes)} 个节点，超过默认的 200 项保护阈值。确认范围后加 --allow-large 再运行。')
        out_dir.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=out_dir.parent, prefix='.hdx-stream-') as tmp:
            stage = Path(tmp); writer = VolumeWriter(stage, prefix, max_mb)
            builder = WordBuilder(pkg, nodes, out, chapters)
            builder.build(writer.page, writer.attachment if merge_word_attachments else None)
            if builder.errors: raise PackageError('正文转换检查未通过，未发布：' + json.dumps(builder.errors[:3], ensure_ascii=False))
            assets = []
            for name, data in builder.attachments.items():
                if len(data) >= writer.limit and not allow_large_attachments:
                    raise PackageError('未合并附件超过上限：' + name + '；未发布分卷。DOC/DOCX 可使用 --merge-word-attachments。')
                rel = out.stem + '_附件/' + name
                path = local_path(stage, rel); path.parent.mkdir(parents=True, exist_ok=True); path.write_bytes(data)
                assets.append({'file': rel, 'bytes': len(data), 'sha256': hashlib.sha256(data).hexdigest()})
            parts = writer.finish(builder)
            report = {'tool_version': VERSION, 'mode': 'streaming', 'source_hdx': str(Path(hdx).resolve()),
                      'chapter_paths': chapters, 'node_count': len(nodes), 'whole_book_created': False,
                      'limit_bytes': writer.limit, 'parts': parts, 'pages': builder.page_reports,
                      'attachments': assets, 'all_docx_within_limit': True,
                      'all_delivered_files_within_limit': all(a['bytes'] < writer.limit for a in assets),
                      'all_selected_content_merged_into_docx': all(p.get('body_merged', True) for p in builder.page_reports),
                      'coverage_contiguous': True, 'source_body_blocks': writer.source_blocks,
                      'body_text_sha256': writer.source_hash.hexdigest(), 'warnings': builder.warnings,
                      'errors': [], 'automatic_checks_passed': True,
                      'visual_review': '未执行；需要在 Word 中检查附件转换、分页、目录与跨卷链接',
                      'elapsed_seconds': round(time.perf_counter() - started, 3)}
            # Publish only validated volumes and assets; temporary source chunks never ship.
            publish = stage / 'publish'; publish.mkdir()
            for info in parts: (stage / info['file']).replace(publish / info['file'])
            for asset in assets:
                dest = local_path(publish, asset['file']); dest.parent.mkdir(parents=True, exist_ok=True)
                local_path(stage, asset['file']).replace(dest)
            (publish / '分卷报告.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
            (publish / '阅读说明.txt').write_text('按文件编号顺序阅读；所有分卷保持在同一文件夹。\n附件正文已合并的部分会单独起卷，保留其原生样式。\n目录需在 Word 中更新。更新并保存后请重新检查大小。\n自动检查不能替代 Word 中的分页、表格与附件转换验收。\n', encoding='utf-8')
            publish.rename(out_dir)
            print(f'完成：{len(parts)} 份 Word，未生成整本文件。耗时 {report["elapsed_seconds"]} 秒。', flush=True)
            return report
    finally:
        pkg.z.close()
