import argparse
import base64
import configparser
from datetime import datetime, timezone
import mimetypes
import re
import uuid
import zipfile
from collections import deque
from pathlib import Path

from lxml import etree


FB2_NS = "http://www.gribuser.ru/xml/fictionbook/2.0"
XLINK_NS = "http://www.w3.org/1999/xlink"
XHTML_NS = "http://www.w3.org/1999/xhtml"
OPF_NS = "http://www.idpf.org/2007/opf"
DC_NS = "http://purl.org/dc/elements/1.1/"

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}
W = "{%s}" % NS["w"]


class Settings:
    """Configuration resolved relative to config.ini."""

    def __init__(self, config_path: Path):
        self.config_path = config_path.resolve()
        self.base = self.config_path.parent
        cfg = configparser.ConfigParser()
        if not self.config_path.exists():
            raise FileNotFoundError(f"Config not found: {self.config_path}")
        cfg.read(self.config_path, encoding="utf-8")

        def path(section, key, default):
            value = cfg.get(section, key, fallback=default).strip()
            candidate = Path(value)
            return candidate if candidate.is_absolute() else self.base / candidate

        self.input_docx = path("book", "input", "book.docx")
        self.metadata = path("book", "metadata", "metadata.txt")
        self.cover = path("book", "cover", "cover.jpg")
        self.normalized_docx = path("book", "normalized", "book.normalized.docx")
        self.output_dir = path("book", "output_dir", ".")
        raw_formats = cfg.get("output", "formats", fallback="fb2, epub")
        self.formats = {x.strip().lower() for x in raw_formats.split(",") if x.strip()}
        unknown = self.formats - {"fb2", "epub"}
        if unknown or not self.formats:
            raise ValueError("formats must contain fb2 and/or epub")
        self.normalize = cfg.getboolean("normalization", "enabled", fallback=True)
        self.report = path("normalization", "report", "conversion-report.txt")


class Metadata:
    def __init__(self, path: Path):
        self.data = {}
        if path.exists():
            for raw_line in path.read_text(encoding="utf-8-sig").splitlines():
                line = raw_line.strip()
                if line and not line.startswith("#") and "=" in line:
                    key, value = line.split("=", 1)
                    self.data[key.strip()] = value.strip()

    def get(self, key, default=""):
        return self.data.get(key, default)


class DocxNormalizer:
    MARKER_RE = re.compile(r"\*{1,3}")
    QUOTE_RE = re.compile(r"^(?:>[ \t]*)+")
    QUOTE_STYLE_RE = re.compile(r"^KONVERTQuote(\d+)$")

    def __init__(self):
        self.messages = []
        self.stats = {"italic": 0, "bold": 0, "bold_italic": 0, "quotes": 0}

    @staticmethod
    def _enabled(node, name):
        prop = node.find(f"w:{name}", NS) if node is not None else None
        if prop is None:
            return False
        value = prop.attrib.get(W + "val", "true").lower()
        return value not in {"0", "false", "off", "none"}

    def _characters(self, paragraph):
        result = []
        for run in paragraph.findall(".//w:r", NS):
            rpr = run.find("w:rPr", NS)
            bold = self._enabled(rpr, "b")
            italic = self._enabled(rpr, "i")
            for text_node in run.findall(".//w:t", NS):
                for char in text_node.text or "":
                    result.append([char, bold, italic])
        return result

    @staticmethod
    def _is_simple(paragraph):
        forbidden = ("hyperlink", "drawing", "pict", "object", "fldSimple", "instrText", "tab", "br")
        return not any(paragraph.findall(f".//w:{name}", NS) for name in forbidden)

    def _parse_markers(self, chars):
        plain = "".join(item[0] for item in chars)
        matches = list(self.MARKER_RE.finditer(plain))
        if not matches:
            return chars, False

        bold = italic = False
        changed = False
        result = []
        marker_by_start = {m.start(): m for m in matches}
        index = 0
        while index < len(chars):
            marker = marker_by_start.get(index)
            if marker:
                size = len(marker.group())
                if size == 3:
                    bold, italic = not bold, not italic
                elif size == 2:
                    bold = not bold
                else:
                    italic = not italic
                index += size
                changed = True
                continue
            char, base_bold, base_italic = chars[index]
            result.append([char, base_bold or bold, base_italic or italic])
            index += 1

        if bold or italic:
            return chars, False
        for marker in matches:
            size = len(marker.group())
            key = {1: "italic", 2: "bold", 3: "bold_italic"}[size]
            self.stats[key] += 1
        return result, changed

    @staticmethod
    def _set_quote_style(paragraph, level):
        ppr = paragraph.find("w:pPr", NS)
        if ppr is None:
            ppr = etree.Element(W + "pPr")
            paragraph.insert(0, ppr)
        style = ppr.find("w:pStyle", NS)
        if style is None:
            style = etree.Element(W + "pStyle")
            ppr.insert(0, style)
        style.attrib[W + "val"] = f"KONVERTQuote{level}"

    @staticmethod
    def _rebuild(paragraph, chars):
        ppr = paragraph.find("w:pPr", NS)
        for child in list(paragraph):
            if child is not ppr:
                paragraph.remove(child)
        groups = []
        for char, bold, italic in chars:
            if groups and groups[-1][1:] == [bold, italic]:
                groups[-1][0] += char
            else:
                groups.append([char, bold, italic])
        for text, bold, italic in groups:
            run = etree.SubElement(paragraph, W + "r")
            if bold or italic:
                rpr = etree.SubElement(run, W + "rPr")
                if bold:
                    etree.SubElement(rpr, W + "b")
                if italic:
                    etree.SubElement(rpr, W + "i")
            node = etree.SubElement(run, W + "t")
            if text.startswith(" ") or text.endswith(" "):
                node.attrib["{http://www.w3.org/XML/1998/namespace}space"] = "preserve"
            node.text = text

    def _ensure_quote_styles(self, styles_xml, levels):
        if not levels:
            return styles_xml
        if styles_xml:
            root = etree.fromstring(styles_xml)
        else:
            root = etree.Element(W + "styles", nsmap={"w": NS["w"]})
        existing = {node.attrib.get(W + "styleId") for node in root.findall("w:style", NS)}
        for level in range(1, max(levels) + 1):
            style_id = f"KONVERTQuote{level}"
            if style_id in existing:
                continue
            style = etree.SubElement(root, W + "style", {W + "type": "paragraph", W + "styleId": style_id})
            etree.SubElement(style, W + "name", {W + "val": f"KONVERT Quote {level}"})
            etree.SubElement(style, W + "basedOn", {W + "val": "Normal"})
            ppr = etree.SubElement(style, W + "pPr")
            etree.SubElement(ppr, W + "ind", {W + "left": str(720 * level)})
            borders = etree.SubElement(ppr, W + "pBdr")
            etree.SubElement(borders, W + "left", {
                W + "val": "single", W + "sz": "12", W + "space": "8", W + "color": "808080"
            })
        return etree.tostring(root, encoding="utf-8", xml_declaration=True, standalone=True)

    def normalize(self, source: Path, destination: Path, report: Path):
        if source.resolve() == destination.resolve():
            raise ValueError("normalized DOCX must not overwrite the source DOCX")
        if not source.exists():
            raise FileNotFoundError(f"DOCX not found: {source}")
        destination.parent.mkdir(parents=True, exist_ok=True)
        quote_levels = set()

        with zipfile.ZipFile(source, "r") as src:
            document = etree.fromstring(src.read("word/document.xml"))
            paragraphs = document.findall(".//w:p", NS)
            table_count = len(document.findall(".//w:tbl", NS))
            if table_count:
                self.messages.append(f"Предупреждение: таблиц в DOCX: {table_count}; они не конвертируются")
            for number, paragraph in enumerate(paragraphs, 1):
                chars = self._characters(paragraph)
                text = "".join(item[0] for item in chars)
                unsupported = []
                for tag, label in (("u", "подчёркивание"), ("strike", "зачёркивание"),
                                   ("dstrike", "двойное зачёркивание"), ("vertAlign", "индекс")):
                    if paragraph.find(f".//w:{tag}", NS) is not None:
                        unsupported.append(label)
                if unsupported:
                    self.messages.append(
                        f"Абзац {number}: неподдерживаемое форматирование: {', '.join(unsupported)}"
                    )
                quote_match = self.QUOTE_RE.match(text)
                has_markers = bool(self.MARKER_RE.search(text))
                if not quote_match and not has_markers:
                    continue
                if not self._is_simple(paragraph):
                    self.messages.append(f"Абзац {number}: пропущен из-за сложных объектов Word")
                    continue

                level = 0
                if quote_match:
                    prefix_end = quote_match.end()
                    level = quote_match.group().count(">")
                    chars = chars[prefix_end:]
                    quote_levels.add(level)
                    self.stats["quotes"] += 1
                    self._set_quote_style(paragraph, level)
                chars, marker_changed = self._parse_markers(chars)
                if has_markers and not marker_changed:
                    self.messages.append(f"Абзац {number}: несбалансированные маркеры *, оставлен без изменения")
                if level or marker_changed:
                    self._rebuild(paragraph, chars)
                    changes = []
                    if level:
                        changes.append(f"цитата уровня {level}")
                    if marker_changed:
                        changes.append("форматирование *")
                    self.messages.append(f"Абзац {number}: " + ", ".join(changes))

            document_xml = etree.tostring(document, encoding="utf-8", xml_declaration=True, standalone=True)
            styles_original = src.read("word/styles.xml") if "word/styles.xml" in src.namelist() else None
            styles_xml = self._ensure_quote_styles(styles_original, quote_levels)
            temporary = destination.with_suffix(destination.suffix + ".tmp")
            with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as dst:
                for item in src.infolist():
                    if item.filename == "word/document.xml":
                        dst.writestr(item, document_xml)
                    elif item.filename == "word/styles.xml":
                        dst.writestr(item, styles_xml)
                    else:
                        dst.writestr(item, src.read(item.filename))
                if "word/styles.xml" not in src.namelist():
                    dst.writestr("word/styles.xml", styles_xml)
            temporary.replace(destination)

        report.parent.mkdir(parents=True, exist_ok=True)
        summary = [
            f"Источник: {source}",
            f"Нормализованный файл: {destination}",
            f"Курсивных маркеров: {self.stats['italic'] // 2}",
            f"Жирных маркеров: {self.stats['bold'] // 2}",
            f"Маркеров жирного курсива: {self.stats['bold_italic'] // 2}",
            f"Абзацев-цитат: {self.stats['quotes']}",
            "",
        ]
        report.write_text("\n".join(summary + self.messages) + "\n", encoding="utf-8")


class DocxParser:
    QUOTE_STYLE_RE = re.compile(r"^KONVERTQuote(\d+)$")

    def __init__(self, path: Path):
        self.zip = zipfile.ZipFile(path)
        self.document = etree.fromstring(self.zip.read("word/document.xml"))
        self.relationships = self._load_relationships()

    def _load_relationships(self):
        rels = {}
        try:
            xml = etree.fromstring(self.zip.read("word/_rels/document.xml.rels"))
            for rel in xml:
                rid, target = rel.attrib.get("Id"), rel.attrib.get("Target")
                if rid and target:
                    rels[rid] = target
        except (KeyError, etree.XMLSyntaxError) as error:
            print("Relationships error:", error)
        return rels

    def blocks(self):
        body = self.document.find("w:body", NS)
        if body is None:
            return []
        return [(etree.QName(child).localname, child) for child in body if etree.QName(child).localname in {"p", "tbl"}]

    def paragraph_style(self, paragraph):
        style = paragraph.find("w:pPr/w:pStyle", NS)
        return style.attrib.get(W + "val") if style is not None else None

    def quote_level(self, paragraph):
        match = self.QUOTE_STYLE_RE.match(self.paragraph_style(paragraph) or "")
        return int(match.group(1)) if match else 0

    def paragraph_text(self, paragraph):
        return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS)).strip()

    @staticmethod
    def _enabled(rpr, name):
        prop = rpr.find(f"w:{name}", NS) if rpr is not None else None
        if prop is None:
            return False
        return prop.attrib.get(W + "val", "true").lower() not in {"0", "false", "off", "none"}

    def paragraph_runs(self, paragraph):
        result = []
        for run in paragraph.findall(".//w:r", NS):
            text = "".join(node.text or "" for node in run.findall(".//w:t", NS))
            if text:
                rpr = run.find("w:rPr", NS)
                result.append({"text": text, "bold": self._enabled(rpr, "b"), "italic": self._enabled(rpr, "i")})
        return result

    def paragraph_images(self, paragraph):
        result = []
        for blip in paragraph.xpath(".//a:blip", namespaces=NS):
            rid = blip.attrib.get("{%s}embed" % NS["r"])
            target = self.relationships.get(rid)
            if not target:
                continue
            target = target.replace("\\", "/")
            full_target = target if target.startswith("word/") else "word/" + target.lstrip("/")
            try:
                result.append((Path(full_target).name, self.zip.read(full_target)))
            except KeyError as error:
                print("Image read error:", full_target, error)
        return result


class FB2Builder:
    NOTE_RE = re.compile(r"\[(\d+)\]")

    def __init__(self, meta: Metadata):
        self.meta = meta
        self.root = etree.Element("FictionBook", nsmap={None: FB2_NS, "l": XLINK_NS})
        self.description = etree.SubElement(self.root, "description")
        self.body = etree.SubElement(self.root, "body")
        self.notes = etree.SubElement(self.root, "body", name="notes")
        self.sections = deque([(0, self.body)])
        self.active_cite = None
        self.added_images = set()
        self._build_metadata()

    def _build_metadata(self):
        title_info = etree.SubElement(self.description, "title-info")
        etree.SubElement(title_info, "genre").text = self.meta.get("genre", "prose_contemporary")
        author = etree.SubElement(title_info, "author")
        name = self.meta.get("author", "Unknown").strip()
        parts = name.split()
        if len(parts) == 2:
            etree.SubElement(author, "first-name").text = parts[0]
            etree.SubElement(author, "last-name").text = parts[1]
        else:
            etree.SubElement(author, "nickname").text = name
        etree.SubElement(title_info, "book-title").text = self.meta.get("title", "Untitled")
        annotation = self.meta.get("annotation")
        if annotation:
            ann = etree.SubElement(title_info, "annotation")
            for value in annotation.split("||"):
                if value.strip():
                    etree.SubElement(ann, "p").text = value.strip()
        etree.SubElement(title_info, "lang").text = self.meta.get("lang", "ru")
        sequence = self.meta.get("sequence")
        if sequence:
            attrs = {"name": sequence}
            if self.meta.get("sequence_number"):
                attrs["number"] = self.meta.get("sequence_number")
            etree.SubElement(title_info, "sequence", **attrs)
        document_info = etree.SubElement(self.description, "document-info")
        doc_author = etree.SubElement(document_info, "author")
        etree.SubElement(doc_author, "nickname").text = "KONVERT"
        etree.SubElement(document_info, "program-used").text = "KONVERT"
        if self.meta.get("date"):
            etree.SubElement(document_info, "date").text = self.meta.get("date")
        etree.SubElement(document_info, "id").text = str(uuid.uuid4())
        etree.SubElement(document_info, "version").text = "1.0"

    def current_section(self):
        # FB2 body must contain sections, not bare paragraphs.
        if self.sections[-1][1] is self.body:
            section = etree.SubElement(self.body, "section")
            self.sections.append((7, section))
        return self.sections[-1][1]

    def open_section(self, level, title):
        self.active_cite = None
        while self.sections and self.sections[-1][0] >= level:
            self.sections.pop()
        parent = self.sections[-1][1]
        section = etree.SubElement(parent, "section")
        heading = etree.SubElement(section, "title")
        etree.SubElement(heading, "p").text = title
        self.sections.append((level, section))

    @staticmethod
    def append_text(parent, text):
        if len(parent):
            parent[-1].tail = (parent[-1].tail or "") + text
        else:
            parent.text = (parent.text or "") + text

    def _formatted(self, parent, text, bold, italic):
        if bold and italic:
            strong = etree.SubElement(parent, "strong")
            etree.SubElement(strong, "emphasis").text = text
        elif bold:
            etree.SubElement(parent, "strong").text = text
        elif italic:
            etree.SubElement(parent, "emphasis").text = text
        else:
            self.append_text(parent, text)

    def _add_inline(self, parent, runs):
        for run in runs:
            pos = 0
            for match in self.NOTE_RE.finditer(run["text"]):
                self._formatted(parent, run["text"][pos:match.start()], run["bold"], run["italic"])
                link = etree.SubElement(parent, "a", type="note")
                link.attrib["{%s}href" % XLINK_NS] = f"#n_{match.group(1)}"
                link.text = match.group()
                pos = match.end()
            self._formatted(parent, run["text"][pos:], run["bold"], run["italic"])

    def add_runs(self, runs, quote_level=0):
        if not runs:
            return
        if quote_level:
            if self.active_cite is None:
                self.active_cite = etree.SubElement(self.current_section(), "cite")
            paragraph = etree.SubElement(self.active_cite, "p")
            if quote_level > 1:
                paragraph.text = "> " * (quote_level - 1)
        else:
            self.active_cite = None
            paragraph = etree.SubElement(self.current_section(), "p")
        self._add_inline(paragraph, runs)

    def add_note(self, note_id, text):
        section = etree.SubElement(self.notes, "section", id=f"n_{note_id}")
        title = etree.SubElement(section, "title")
        etree.SubElement(title, "p").text = note_id
        etree.SubElement(section, "p").text = text

    def add_binary(self, image_id, data):
        if image_id in self.added_images:
            return
        mime = mimetypes.guess_type(image_id)[0]
        if mime not in {"image/jpeg", "image/png", "image/gif"}:
            mime = "image/jpeg"
        binary = etree.SubElement(self.root, "binary", id=image_id)
        binary.attrib["content-type"] = mime
        binary.text = base64.b64encode(data).decode("ascii")
        self.added_images.add(image_id)

    def add_image(self, image_id, data):
        self.active_cite = None
        self.add_binary(image_id, data)
        paragraph = etree.SubElement(self.current_section(), "p")
        image = etree.SubElement(paragraph, "image")
        image.attrib["{%s}href" % XLINK_NS] = f"#{image_id}"

    def set_cover(self, image_id, data):
        self.add_binary(image_id, data)
        title_info = self.description.find("title-info")
        cover = etree.Element("coverpage")
        language = title_info.find("lang")
        title_info.insert(title_info.index(language), cover)
        image = etree.SubElement(cover, "image")
        image.attrib["{%s}href" % XLINK_NS] = f"#{image_id}"

    def save(self, path):
        path.write_bytes(etree.tostring(self.root, encoding="utf-8", pretty_print=True, xml_declaration=True))


class EPUBBuilder:
    NOTE_RE = re.compile(r"\[(\d+)\]")

    def __init__(self, meta: Metadata):
        self.meta = meta
        self.identifier = str(uuid.uuid4())
        self.html = etree.Element("{%s}html" % XHTML_NS, nsmap={None: XHTML_NS})
        head = etree.SubElement(self.html, "{%s}head" % XHTML_NS)
        etree.SubElement(head, "{%s}title" % XHTML_NS).text = meta.get("title", "Untitled")
        etree.SubElement(head, "{%s}link" % XHTML_NS, rel="stylesheet", type="text/css", href="style.css")
        self.body = etree.SubElement(self.html, "{%s}body" % XHTML_NS)
        self.quote_stack = []
        self.toc = []
        self.images = {}
        self.cover_id = None

    def _reset_quotes(self):
        self.quote_stack = []

    def open_section(self, level, title):
        self._reset_quotes()
        anchor = f"heading-{len(self.toc) + 1}"
        heading = etree.SubElement(self.body, "{%s}h%d" % (XHTML_NS, level), id=anchor)
        heading.text = title
        self.toc.append((level, title, anchor))

    @staticmethod
    def _append_text(parent, text):
        if not text:
            return
        if len(parent):
            parent[-1].tail = (parent[-1].tail or "") + text
        else:
            parent.text = (parent.text or "") + text

    def _formatted(self, parent, text, bold, italic):
        if not text:
            return
        current = parent
        if bold:
            current = etree.SubElement(current, "{%s}strong" % XHTML_NS)
        if italic:
            current = etree.SubElement(current, "{%s}em" % XHTML_NS)
        if current is parent:
            self._append_text(parent, text)
        else:
            current.text = text

    def add_runs(self, runs, quote_level=0):
        if not runs:
            return
        if quote_level:
            while len(self.quote_stack) > quote_level:
                self.quote_stack.pop()
            while len(self.quote_stack) < quote_level:
                parent = self.quote_stack[-1] if self.quote_stack else self.body
                self.quote_stack.append(etree.SubElement(parent, "{%s}blockquote" % XHTML_NS))
            parent = self.quote_stack[-1]
        else:
            self._reset_quotes()
            parent = self.body
        paragraph = etree.SubElement(parent, "{%s}p" % XHTML_NS)
        for run in runs:
            pos = 0
            for match in self.NOTE_RE.finditer(run["text"]):
                self._formatted(paragraph, run["text"][pos:match.start()], run["bold"], run["italic"])
                link = etree.SubElement(paragraph, "{%s}a" % XHTML_NS, href=f"#note-{match.group(1)}", **{"class": "note-link"})
                link.text = match.group()
                pos = match.end()
            self._formatted(paragraph, run["text"][pos:], run["bold"], run["italic"])

    def add_note(self, note_id, text):
        self._reset_quotes()
        aside = etree.SubElement(self.body, "{%s}aside" % XHTML_NS, id=f"note-{note_id}", **{"class": "footnote"})
        etree.SubElement(aside, "{%s}p" % XHTML_NS).text = f"[{note_id}] {text}"

    def add_image(self, image_id, data):
        self._reset_quotes()
        self.images.setdefault(image_id, data)
        paragraph = etree.SubElement(self.body, "{%s}p" % XHTML_NS, **{"class": "image"})
        etree.SubElement(paragraph, "{%s}img" % XHTML_NS, src=f"images/{image_id}", alt="")

    def set_cover(self, image_id, data):
        self.images[image_id] = data
        self.cover_id = image_id

    def _nav(self):
        html = etree.Element("{%s}html" % XHTML_NS, nsmap={None: XHTML_NS, "epub": "http://www.idpf.org/2007/ops"})
        head = etree.SubElement(html, "{%s}head" % XHTML_NS)
        etree.SubElement(head, "{%s}title" % XHTML_NS).text = "Оглавление"
        body = etree.SubElement(html, "{%s}body" % XHTML_NS)
        nav = etree.SubElement(body, "{%s}nav" % XHTML_NS)
        nav.attrib["{http://www.idpf.org/2007/ops}type"] = "toc"
        etree.SubElement(nav, "{%s}h1" % XHTML_NS).text = "Оглавление"
        ordered = etree.SubElement(nav, "{%s}ol" % XHTML_NS)
        entries = self.toc or [(1, self.meta.get("title", "Книга"), "book-start")]
        if not self.toc:
            self.body.attrib["id"] = "book-start"
        for _, title, anchor in entries:
            item = etree.SubElement(ordered, "{%s}li" % XHTML_NS)
            etree.SubElement(item, "{%s}a" % XHTML_NS, href=f"text.xhtml#{anchor}").text = title
        return etree.tostring(html, encoding="utf-8", xml_declaration=True, doctype="<!DOCTYPE html>")

    def _package(self):
        package = etree.Element(
            "{%s}package" % OPF_NS,
            nsmap={None: OPF_NS, "dc": DC_NS},
            attrib={"version": "3.0", "unique-identifier": "book-id"},
        )
        metadata = etree.SubElement(package, "{%s}metadata" % OPF_NS)
        etree.SubElement(metadata, "{%s}identifier" % DC_NS, id="book-id").text = self.identifier
        etree.SubElement(metadata, "{%s}title" % DC_NS).text = self.meta.get("title", "Untitled")
        etree.SubElement(metadata, "{%s}creator" % DC_NS).text = self.meta.get("author", "Unknown")
        etree.SubElement(metadata, "{%s}language" % DC_NS).text = self.meta.get("lang", "ru")
        modified = etree.SubElement(metadata, "{%s}meta" % OPF_NS, property="dcterms:modified")
        modified.text = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        if self.meta.get("date"):
            etree.SubElement(metadata, "{%s}date" % DC_NS).text = self.meta.get("date")
        if self.meta.get("annotation"):
            etree.SubElement(metadata, "{%s}description" % DC_NS).text = self.meta.get("annotation").replace("||", "\n")
        manifest = etree.SubElement(package, "{%s}manifest" % OPF_NS)
        etree.SubElement(manifest, "{%s}item" % OPF_NS, attrib={"id": "text", "href": "text.xhtml", "media-type": "application/xhtml+xml"})
        etree.SubElement(manifest, "{%s}item" % OPF_NS, attrib={"id": "nav", "href": "nav.xhtml", "media-type": "application/xhtml+xml", "properties": "nav"})
        etree.SubElement(manifest, "{%s}item" % OPF_NS, attrib={"id": "css", "href": "style.css", "media-type": "text/css"})
        for number, image_id in enumerate(self.images, 1):
            attrs = {"id": f"image-{number}", "href": f"images/{image_id}", "media-type": mimetypes.guess_type(image_id)[0] or "image/jpeg"}
            if image_id == self.cover_id:
                attrs["properties"] = "cover-image"
            etree.SubElement(manifest, "{%s}item" % OPF_NS, attrib=attrs)
        spine = etree.SubElement(package, "{%s}spine" % OPF_NS)
        etree.SubElement(spine, "{%s}itemref" % OPF_NS, idref="text")
        return etree.tostring(package, encoding="utf-8", xml_declaration=True, pretty_print=True)

    def save(self, path):
        container = b'''<?xml version="1.0" encoding="UTF-8"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
  <rootfiles><rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/></rootfiles>
</container>'''
        css = b'''body { font-family: serif; line-height: 1.45; }
blockquote { border-left: 0.25em solid #888; margin: 1em 0 1em 1em; padding-left: 1em; }
.image { text-align: center; } img { max-width: 100%; } .footnote { font-size: 0.9em; }'''
        path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("mimetype", "application/epub+zip", compress_type=zipfile.ZIP_STORED)
            archive.writestr("META-INF/container.xml", container, compress_type=zipfile.ZIP_DEFLATED)
            archive.writestr("OEBPS/content.opf", self._package(), compress_type=zipfile.ZIP_DEFLATED)
            archive.writestr("OEBPS/nav.xhtml", self._nav(), compress_type=zipfile.ZIP_DEFLATED)
            archive.writestr("OEBPS/text.xhtml", etree.tostring(self.html, encoding="utf-8", xml_declaration=True, doctype="<!DOCTYPE html>"), compress_type=zipfile.ZIP_DEFLATED)
            archive.writestr("OEBPS/style.css", css, compress_type=zipfile.ZIP_DEFLATED)
            for image_id, data in self.images.items():
                archive.writestr(f"OEBPS/images/{image_id}", data, compress_type=zipfile.ZIP_DEFLATED)


class Converter:
    HEADINGS = {f"Heading{level}": level for level in range(1, 7)}
    NOTE_LINE = re.compile(r"^\[(\d+)\]\s*(.+)$")

    def __init__(self, settings: Settings):
        self.settings = settings
        self.meta = Metadata(settings.metadata)

    def run(self):
        source = self.settings.input_docx
        if self.settings.normalize:
            normalizer = DocxNormalizer()
            normalizer.normalize(source, self.settings.normalized_docx, self.settings.report)
            source = self.settings.normalized_docx
            print("Normalized:", source)
            print("Report:", self.settings.report)

        document = DocxParser(source)
        builders = []
        if "fb2" in self.settings.formats:
            builders.append(("fb2", FB2Builder(self.meta)))
        if "epub" in self.settings.formats:
            builders.append(("epub", EPUBBuilder(self.meta)))

        if self.settings.cover.exists():
            cover_data = self.settings.cover.read_bytes()
            for _, builder in builders:
                builder.set_cover(self.settings.cover.name, cover_data)

        pending_notes = []
        referenced_notes = set()
        validation_messages = []
        for required in ("title", "author"):
            if not self.meta.get(required):
                validation_messages.append(f"Предупреждение: metadata.txt не содержит {required}")
        for kind, block in document.blocks():
            if kind == "tbl":
                print("Warning: table skipped")
                validation_messages.append("Предупреждение: таблица пропущена")
                continue
            text = document.paragraph_text(block)
            runs = document.paragraph_runs(block)
            images = document.paragraph_images(block)
            if not text and not images:
                continue
            note = self.NOTE_LINE.match(text) if text else None
            if note:
                pending_notes.append((note.group(1), note.group(2)))
                continue
            referenced_notes.update(FB2Builder.NOTE_RE.findall(text))
            style = document.paragraph_style(block)
            if text and style in self.HEADINGS:
                for _, builder in builders:
                    builder.open_section(self.HEADINGS[style], text)
                continue
            if runs:
                quote_level = document.quote_level(block)
                for _, builder in builders:
                    builder.add_runs(runs, quote_level)
            for image_id, data in images:
                if image_id.lower() != self.settings.cover.name.lower():
                    for _, builder in builders:
                        builder.add_image(image_id, data)

        for note_id, note_text in pending_notes:
            for _, builder in builders:
                builder.add_note(note_id, note_text)

        defined_notes = {note_id for note_id, _ in pending_notes}
        for note_id in sorted(referenced_notes - defined_notes, key=int):
            validation_messages.append(f"Предупреждение: для ссылки [{note_id}] отсутствует текст сноски")
        for note_id in sorted(defined_notes - referenced_notes, key=int):
            validation_messages.append(f"Предупреждение: сноска [{note_id}] нигде не используется")
        if self.settings.normalize and validation_messages:
            with self.settings.report.open("a", encoding="utf-8") as report_file:
                report_file.write("\nПроверка конвертера:\n")
                report_file.write("\n".join(validation_messages) + "\n")

        self.settings.output_dir.mkdir(parents=True, exist_ok=True)
        for extension, builder in builders:
            output = self.settings.output_dir / f"book.{extension}"
            builder.save(output)
            print("Saved:", output)


def find_config(argument):
    target = Path(argument)
    return target / "config.ini" if target.is_dir() else target


def main():
    parser = argparse.ArgumentParser(description="DOCX to FB2/EPUB converter")
    parser.add_argument("config", nargs="?", default="config.ini", help="Path to config.ini or its folder")
    args = parser.parse_args()
    Converter(Settings(find_config(args.config))).run()


if __name__ == "__main__":
    main()
