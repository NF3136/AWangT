# -*- coding: utf-8 -*-
"""Small synthetic integration tests; never exports the actual full manual."""
import base64
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile
from lxml import etree

from export_word import run_export, PackageError, W, decode

PNG = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jK1sAAAAASUVORK5CYII=")


def fixture(path, broken=False):
    nav = '''<topics><topic txt="手册" url="a/index.html">
      <topic txt="第二章" url="b/index.html"/>
    </topic><topic txt="其他" url="other/index.html"/></topics>'''
    one = '''<html><head><meta charset="utf-8"></head><body><a name="top"></a>
    <div class="articleBoxWithoutHead"><h1>手册</h1><p>中文正文与<b>粗体</b>。</p>
    <h2 id="first">示例</h2><p><a href="../B/INDEX.html?mode=1#sec">跨页引用</a>，
    <a href="../other/index.html">章节外同名页</a>，<a href="../files/附件.txt">附件</a>。</p>
    <table id="table"><caption>表一</caption><thead><tr><th>参数</th><th>说明</th></tr></thead>
    <tbody><tr><td rowspan="2">合并竖向</td><td>第一行</td></tr><tr><td>第二行</td></tr>
    <tr><td colspan="2">合并横向</td></tr></tbody></table>
    <p><img src="../img/p.png" alt="示意图"></p><pre>  line1\n  line2 &lt;tag&gt;</pre>
    <ul><li>列表一<ol><li>子列表</li></ol></li><li>列表二</li></ul></div></body></html>'''
    two = '''<html><head><meta charset="gb2312"></head><body><div class="articleBoxWithoutHead">
    <h1>第二章</h1><h2 id="sec">目标小节</h2><p>第二页中文。</p>
    <p><a href="../a/index.html#table">返回表格</a></p></div></body></html>'''
    with ZipFile(path, "w") as z:
        z.writestr("resources/navi.xml", nav.encode())
        z.writestr("profile.xml", '<profile><libName>测试文档</libName><productVersion>V1</productVersion></profile>'.encode())
        z.writestr("resources/a/index.html", one.encode())
        z.writestr("resources/b/index.html", two.encode("gb18030"))
        z.writestr("resources/other/index.html", b'<html><body>outside</body></html>')
        if not broken: z.writestr("resources/img/p.png", PNG)
        z.writestr("resources/files/附件.txt", '附件内容'.encode())


class ExportWordTests(unittest.TestCase):
    def test_decode_reports_filename_and_byte_position(self):
        with self.assertRaisesRegex(PackageError, "resources/bad.html") as exc:
            decode(b'<meta charset="utf-8">\xff', "resources/bad.html")
        self.assertIn("字节位置", str(exc.exception))

    def test_bom_encodings_are_strict_and_supported(self):
        text = '<html><body>中文测试</body></html>'
        for encoding in ("utf-8-sig", "utf-16", "utf-32"):
            self.assertEqual(decode(text.encode(encoding))[0], text)

    def export(self, *args, **kwargs):
        with redirect_stdout(StringIO()):
            return run_export(*args, **kwargs)

    def test_links_tables_text_images_and_attachments(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d); source = p / "sample.hdx"; out = p / "result.docx"
            fixture(source)
            r = self.export(source, ["手册"], out)
            self.assertEqual(r["node_count"], 2)
            self.assertFalse(r["errors"], r["errors"])
            self.assertEqual(r["counts"]["tables"], 1)
            self.assertEqual(r["counts"]["merged_cells"], 2)
            self.assertEqual(r["counts"]["image_occurrences"], 1)
            self.assertTrue(all(p["text_intact"] for p in r["pages"]))
            self.assertEqual(r["outside_references"][0]["target"], "resources/other/index.html")
            self.assertEqual((p / "result_附件/resources/files/附件.txt").read_text(encoding="utf-8"), "附件内容")
            with ZipFile(out) as z:
                root = etree.fromstring(z.read("word/document.xml"))
                bookmarks = [e.get(W + "name") for e in root.iter(W + "bookmarkStart")]
                self.assertEqual(len(bookmarks), len(set(bookmarks)))
                self.assertTrue(all(e.get(W + "anchor") in bookmarks for e in root.iter(W + "hyperlink") if e.get(W + "anchor")))
                self.assertTrue(list(root.iter(W + "gridSpan")))
                self.assertTrue(list(root.iter(W + "vMerge")))
                text = "".join(e.text or "" for e in root.iter(W + "t"))
                for phrase in ["合并竖向", "第二行", "合并横向", "第二页中文", "列表一", "子列表"]:
                    self.assertIn(phrase, text)
                media = [n for n in z.namelist() if n.startswith("word/media/")]
                self.assertEqual(len(media), 1)
                self.assertEqual(z.read(media[0]), PNG)

    def test_existing_output_is_protected(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d); fixture(p / "sample.hdx")
            (p / "result.docx").write_bytes(b"do not overwrite")
            with self.assertRaises(PackageError):
                self.export(p / "sample.hdx", ["手册"], p / "result.docx")
            self.assertEqual((p / "result.docx").read_bytes(), b"do not overwrite")

    def test_missing_image_is_a_failure_not_a_silent_pass(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d); fixture(p / "sample.hdx", broken=True)
            r = self.export(p / "sample.hdx", ["手册"], p / "result.docx")
            self.assertFalse(r["automatic_checks_passed"])
            self.assertTrue(any("图片" in e["reason"] for e in r["errors"]))

    def test_dry_run_does_not_write(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d); fixture(p / "sample.hdx")
            r = self.export(p / "sample.hdx", ["手册"], p / "result.docx", dry_run=True)
            self.assertEqual(r["node_count"], 2)
            self.assertFalse((p / "result.docx").exists())

    def test_selection_overlap_is_rejected(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d); fixture(p / "sample.hdx")
            with self.assertRaises(PackageError):
                self.export(p / "sample.hdx", ["手册", "手册 > 第二章"], p / "result.docx")

    def test_shared_body_and_multi_selection(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d); source = p / "shared.hdx"
            with ZipFile(source, "w") as z:
                z.writestr("resources/navi.xml", '<topics><topic txt="范围一" url="a.html"><topic txt="别名" url="a.html"/></topic><topic txt="范围二" url="b.html"/></topics>'.encode())
                z.writestr("resources/a.html", '<html><meta charset="utf-8"><body><h1>范围一</h1><p>共用正文只出现一次。</p></body></html>'.encode())
                z.writestr("resources/b.html", '<html><meta charset="utf-8"><body><h1>范围二</h1><p>第二个范围。</p></body></html>'.encode())
            r = self.export(source, ["范围一", "范围二"], p / "result.docx")
            self.assertEqual(r["node_count"], 3)
            self.assertEqual(r["unique_body_files"], 2)
            self.assertFalse(r["errors"])
            self.assertTrue(r["pages"][1]["shared_body"])
            with ZipFile(p / "result.docx") as z:
                root = etree.fromstring(z.read("word/document.xml"))
                text = "".join(e.text or "" for e in root.iter(W + "t"))
                self.assertEqual(text.count("共用正文只出现一次。"), 1)

    def test_large_scope_requires_explicit_flag(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d); source = p / "large.hdx"
            with ZipFile(source, "w") as z:
                topics = ''.join(f'<topic txt="子项{i}" url="a.html"/>' for i in range(200))
                z.writestr("resources/navi.xml", ('<topics><topic txt="手册" url="a.html">' + topics + '</topic></topics>').encode())
                z.writestr("resources/a.html", b'<html><body>test</body></html>')
            with self.assertRaisesRegex(PackageError, "200"):
                self.export(source, ["手册"], p / "result.docx")
            self.assertFalse((p / "result.docx").exists())

    def test_office_directory_node_is_preserved_not_decoded(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d); source = p / "office.hdx"
            binary = bytes.fromhex("d0 cf 11 e0 a1 b1 1a e1") + bytes(range(256))
            with ZipFile(source, "w") as z:
                z.writestr("resources/navi.xml", '<topics><topic txt="手册" url="a.html"><topic txt="原始指导书" url="Files/Manual.doc"/></topic></topics>'.encode())
                z.writestr("resources/a.html", '<html><meta charset="utf-8"><body><h1>手册</h1><p><a href="files/manual.doc">查看指导书</a></p></body></html>'.encode())
                z.writestr("resources/files/manual.doc", binary)
            out = p / "result.docx"
            r = self.export(source, ["手册"], out)
            self.assertFalse(r["errors"], r["errors"])
            self.assertEqual(r["node_count"], 2)
            self.assertEqual(r["unique_body_files"], 1)
            self.assertEqual(r["attachment_source_files"], 1)
            self.assertFalse(r["all_selected_content_merged_into_docx"])
            self.assertFalse(r["pages"][1]["body_merged"])
            self.assertEqual((p / "result_附件/resources/files/manual.doc").read_bytes(), binary)
            with ZipFile(out) as z:
                root = etree.fromstring(z.read("word/document.xml"))
                names = {e.get(W + "name") for e in root.iter(W + "bookmarkStart")}
                self.assertTrue(all(e.get(W + "anchor") in names for e in root.iter(W + "hyperlink") if e.get(W + "anchor")))
                self.assertIn('result_', z.read('word/_rels/document.xml.rels').decode())


if __name__ == "__main__":
    unittest.main(verbosity=2)
