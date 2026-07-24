import tempfile
import unittest
import zipfile
from pathlib import Path

from lxml import etree

from main import FB2_NS, NS, XHTML_NS, Converter, Settings


DOCUMENT_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
 <w:body>
  <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Глава</w:t></w:r></w:p>
  <w:p><w:r><w:t>До *несколько слов* и **жирный текст** и ***оба вида*** после[1].</w:t></w:r></w:p>
  <w:p><w:r><w:t>&gt;&gt; вложенная *цитата из нескольких слов*</w:t></w:r></w:p>
  <w:p><w:r><w:t>[1] Текст сноски</w:t></w:r></w:p>
  <w:sectPr/>
 </w:body>
</w:document>'''.encode("utf-8")

STYLES_XML = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
 <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>
</w:styles>'''


def create_docx(path):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", DOCUMENT_XML)
        archive.writestr("word/styles.xml", STYLES_XML)


class ConverterTest(unittest.TestCase):
    def test_normalization_and_both_outputs(self):
        with tempfile.TemporaryDirectory() as temporary:
            folder = Path(temporary)
            create_docx(folder / "source.docx")
            (folder / "metadata.txt").write_text(
                "title=Тестовая книга\nauthor=Иван Иванов\ngenre=fantasy\nlang=ru\n",
                encoding="utf-8",
            )
            (folder / "config.ini").write_text(
                """[book]
input=source.docx
metadata=metadata.txt
cover=missing.jpg
normalized=normalized.docx
output_dir=out
[output]
formats=fb2, epub
[normalization]
enabled=yes
report=report.txt
""",
                encoding="utf-8",
            )

            Converter(Settings(folder / "config.ini")).run()

            with zipfile.ZipFile(folder / "normalized.docx") as archive:
                normalized = etree.fromstring(archive.read("word/document.xml"))
            normalized_text = "".join(normalized.itertext())
            self.assertNotIn("*", normalized_text)
            quote_style = normalized.xpath(
                ".//w:pPr/w:pStyle[starts-with(@w:val, 'KONVERTQuote')]",
                namespaces=NS,
            )[0]
            self.assertEqual(quote_style.attrib["{%s}val" % NS["w"]], "KONVERTQuote2")
            self.assertGreaterEqual(len(normalized.findall(".//w:b", NS)), 2)
            self.assertGreaterEqual(len(normalized.findall(".//w:i", NS)), 3)

            fb2 = etree.parse(str(folder / "out" / "book.fb2"))
            fb2_ns = {"f": FB2_NS}
            self.assertTrue(fb2.xpath("boolean(//f:strong)", namespaces=fb2_ns))
            self.assertTrue(fb2.xpath("boolean(//f:emphasis)", namespaces=fb2_ns))
            cite_text = "".join(fb2.xpath("//f:cite/f:p//text()", namespaces=fb2_ns))
            self.assertTrue(cite_text.startswith("> "))

            with zipfile.ZipFile(folder / "out" / "book.epub") as epub:
                self.assertEqual(epub.namelist()[0], "mimetype")
                self.assertEqual(epub.getinfo("mimetype").compress_type, zipfile.ZIP_STORED)
                xhtml = etree.fromstring(epub.read("OEBPS/text.xhtml"))
                package = etree.fromstring(epub.read("OEBPS/content.opf"))
            xhtml_ns = {"x": XHTML_NS}
            self.assertTrue(xhtml.xpath("boolean(//x:blockquote/x:blockquote)", namespaces=xhtml_ns))
            self.assertTrue(xhtml.xpath("boolean(//x:strong)", namespaces=xhtml_ns))
            self.assertTrue(xhtml.xpath("boolean(//x:em)", namespaces=xhtml_ns))
            self.assertEqual(package.attrib["unique-identifier"], "book-id")


if __name__ == "__main__":
    unittest.main()
