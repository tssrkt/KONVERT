import argparse
import base64
import mimetypes
import re
import uuid
import zipfile
from collections import deque
from pathlib import Path

from lxml import etree


# =========================================================
# NAMESPACES
# =========================================================

FB2_NS = "http://www.gribuser.ru/xml/fictionbook/2.0"
XLINK_NS = "http://www.w3.org/1999/xlink"

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
}


# =========================================================
# METADATA
# =========================================================

class Metadata:

    def __init__(self, path: Path):

        self.data = {}

        if not path.exists():
            return

        text = path.read_text(
            encoding="utf-8"
        )

        for line in text.splitlines():

            line = line.strip()

            if not line:
                continue

            if "=" not in line:
                continue

            k, v = line.split("=", 1)

            self.data[k.strip()] = v.strip()

    # -----------------------------------------------------

    def get(self, key, default=""):
        return self.data.get(key, default)


# =========================================================
# DOCX PARSER
# =========================================================

class DocxParser:

    def __init__(self, path: Path):

        self.zip = zipfile.ZipFile(path)

        self.document = etree.fromstring(
            self.zip.read("word/document.xml")
        )

        self.relationships = self._load_relationships()

    # -----------------------------------------------------

    def _load_relationships(self):

        rels = {}

        try:

            xml = etree.fromstring(
                self.zip.read(
                    "word/_rels/document.xml.rels"
                )
            )

            for rel in xml:

                rid = rel.attrib.get("Id")
                target = rel.attrib.get("Target")

                if rid and target:
                    rels[rid] = target

        except Exception as e:

            print(
                "Relationships error:",
                e
            )

        return rels

    # -----------------------------------------------------

    def blocks(self):

        body = self.document.find(
            "w:body",
            NS
        )

        if body is None:
            return []

        result = []

        for child in body:

            tag = etree.QName(child).localname

            if tag == "p":

                result.append(
                    ("p", child)
                )

            elif tag == "tbl":

                result.append(
                    ("tbl", child)
                )

        return result

    # -----------------------------------------------------

    def paragraph_style(self, p):

        ppr = p.find("w:pPr", NS)

        if ppr is None:
            return None

        style = ppr.find("w:pStyle", NS)

        if style is None:
            return None

        return style.attrib.get(
            f"{{{NS['w']}}}val"
        )

    # -----------------------------------------------------

    def paragraph_text(self, p):

        out = []

        for t in p.findall(".//w:t", NS):

            if t.text:
                out.append(t.text)

        return "".join(out).strip()

    # -----------------------------------------------------

    def paragraph_runs(self, p):

        result = []

        field_active = False
        field_display = False
        field_instruction = []
        field_href = None

        for r in p.findall(".//w:r", NS):

            hyperlink = self._ancestor(
                r,
                "hyperlink",
                p
            )

            simple_field = self._ancestor(
                r,
                "fldSimple",
                p
            )

            href = None

            if hyperlink is not None:
                href = self._hyperlink_target(
                    hyperlink
                )

            elif simple_field is not None:
                instruction = simple_field.attrib.get(
                    f"{{{NS['w']}}}instr",
                    ""
                )

                href = self._field_hyperlink(
                    instruction
                )

            field_chars = r.findall(
                ".//w:fldChar",
                NS
            )

            for field_char in field_chars:

                field_type = field_char.attrib.get(
                    f"{{{NS['w']}}}fldCharType"
                )

                if field_type == "begin":
                    field_active = True
                    field_display = False
                    field_instruction = []
                    field_href = None

                elif (
                    field_type == "separate"
                    and
                    field_active
                ):
                    field_href = self._field_hyperlink(
                        "".join(field_instruction)
                    )
                    field_display = True

            if field_active and not field_display:

                field_instruction.extend(
                    node.text or ""
                    for node in r.findall(
                        ".//w:instrText",
                        NS
                    )
                )

            if (
                href is None
                and
                field_active
                and
                field_display
            ):
                href = field_href

            text = "".join(
                t.text or ""
                for t in r.findall(".//w:t", NS)
            )

            if text:

                rpr = r.find("w:rPr", NS)

                bold = False
                italic = False

                if rpr is not None:

                    bold = (
                        rpr.find("w:b", NS)
                        is not None
                    )

                    italic = (
                        rpr.find("w:i", NS)
                        is not None
                    )

                result.append({
                    "text": text,
                    "bold": bold,
                    "italic": italic,
                    "href": href,
                })

            if any(
                field_char.attrib.get(
                    f"{{{NS['w']}}}fldCharType"
                ) == "end"
                for field_char in field_chars
            ):
                field_active = False
                field_display = False
                field_instruction = []
                field_href = None

        return result

    # -----------------------------------------------------

    def _ancestor(self, element, name, stop):

        parent = element.getparent()

        while parent is not None and parent is not stop:

            if etree.QName(parent).localname == name:
                return parent

            parent = parent.getparent()

        return None

    # -----------------------------------------------------

    def _hyperlink_target(self, hyperlink):

        rid = hyperlink.attrib.get(
            f"{{{NS['r']}}}id"
        )

        if rid:

            target = self.relationships.get(rid)

            if target:
                return target

        anchor = hyperlink.attrib.get(
            f"{{{NS['w']}}}anchor"
        )

        if anchor:
            return f"#{anchor}"

        return None

    # -----------------------------------------------------

    def _field_hyperlink(self, instruction):

        if not instruction:
            return None

        if not re.search(
            r"\bHYPERLINK\b",
            instruction,
            re.IGNORECASE
        ):
            return None

        bookmark = re.search(
            r'\\l\s+"([^"]+)"',
            instruction,
            re.IGNORECASE
        )

        if bookmark:
            return f"#{bookmark.group(1)}"

        target = re.search(
            r'\bHYPERLINK\s+(?:"([^"]+)"|(\S+))',
            instruction,
            re.IGNORECASE
        )

        if not target:
            return None

        return target.group(1) or target.group(2)

    # -----------------------------------------------------

    def paragraph_images(self, p):

        result = []

        blips = p.xpath(".//a:blip | .//w:drawing//a:blip", namespaces={
            "a": NS["a"],
            "w": NS["w"],
            "r": NS["r"]
        })

        for blip in blips:

            rid = blip.attrib.get(
                f"{{{NS['r']}}}embed"
            )

            if not rid:
                continue

            target = self.relationships.get(rid)

            if not target:
                continue

            target = target.replace(
                "\\",
                "/"
            )

            if not target.startswith("word/"):

                target = (
                    "word/"
                    + target.lstrip("/")
                )

            try:

                data = self.zip.read(target)

                result.append(
                    (
                        Path(target).name,
                        data
                    )
                )

            except Exception as e:

                print(
                    "Image read error:",
                    target,
                    e
                )

        return result

    # -----------------------------------------------------

    def parse_table(self, tbl):

        rows = []

        for tr in tbl.findall("./w:tr", NS):

            row = []

            for tc in tr.findall("./w:tc", NS):

                cell_paragraphs = []

                paragraphs = tc.findall("./w:p", NS)

                for p in paragraphs:

                    runs = self.paragraph_runs(p)
                    images = self.paragraph_images(p)

                    if runs or images:
                        cell_paragraphs.append({
                            "runs": runs,
                            "images": images
                        })

                row.append(cell_paragraphs)

            rows.append(row)

        return rows


# =========================================================
# FB2 BUILDER
# =========================================================

class FB2Builder:

    NOTE_RE = re.compile(
        r"\[(\d+)\]"
    )

    def __init__(self, meta: Metadata):

        self.meta = meta

        self.root = etree.Element(
            "FictionBook",
            nsmap={
                None: FB2_NS,
                "l": XLINK_NS
            }
        )

        self.description = etree.SubElement(
            self.root,
            "description"
        )

        self.body = etree.SubElement(
            self.root,
            "body"
        )

        self.notes = etree.SubElement(
            self.root,
            "body",
            name="notes"
        )

        self.sections = deque([
            (0, self.body)
        ])

        self.added_images = set()

        self._build_metadata()

    # -----------------------------------------------------

    def _build_metadata(self):

        ti = etree.SubElement(
            self.description,
            "title-info"
        )

        etree.SubElement(
            ti,
            "book-title"
        ).text = self.meta.get(
            "title",
            "Untitled"
        )

        author_name = self.meta.get(
            "author",
            "Unknown"
        ).strip()

        author = etree.SubElement(
            ti,
            "author"
        )

        parts = author_name.split()

        if len(parts) == 1:

            etree.SubElement(
                author,
                "nickname"
            ).text = parts[0]

        elif len(parts) == 2:

            etree.SubElement(
                author,
                "first-name"
            ).text = parts[0]

            etree.SubElement(
                author,
                "last-name"
            ).text = parts[1]

        else:

            etree.SubElement(
                author,
                "nickname"
            ).text = author_name

        etree.SubElement(
            ti,
            "lang"
        ).text = self.meta.get(
            "lang",
            "ru"
        )

        annotation = self.meta.get(
            "annotation"
        )

        if annotation:

            ann = etree.SubElement(
                ti,
                "annotation"
            )

            for paragraph in annotation.split("||"):
                paragraph = paragraph.strip()

                if paragraph:
                    etree.SubElement(
                        ann,
                        "p"
                    ).text = paragraph

        sequence = self.meta.get(
            "sequence"
        )

        if sequence:

            seq = etree.SubElement(
                ti,
                "sequence"
            )

            seq.attrib["name"] = sequence

        doc = etree.SubElement(
            self.description,
            "document-info"
        )

        etree.SubElement(
            doc,
            "id"
        ).text = str(uuid.uuid4())

    # -----------------------------------------------------

    def current_section(self):

        return self.sections[-1][1]

    # -----------------------------------------------------

    def open_section(self, level, title):

        while (
            self.sections
            and
            self.sections[-1][0] >= level
        ):
            self.sections.pop()

        parent = self.sections[-1][1]

        sec = etree.SubElement(
            parent,
            "section"
        )

        t = etree.SubElement(
            sec,
            "title"
        )

        etree.SubElement(
            t,
            "p"
        ).text = title

        self.sections.append(
            (level, sec)
        )

    # -----------------------------------------------------

    def append_text(
        self,
        parent,
        text
    ):

        if not text:
            return

        if len(parent):

            parent[-1].tail = (
                parent[-1].tail or ""
            ) + text

        else:

            parent.text = (
                parent.text or ""
            ) + text

    # -----------------------------------------------------

    def _append_formatted(
        self,
        parent,
        text,
        bold,
        italic
    ):

        if not text:
            return

        if bold and italic:

            strong = etree.SubElement(
                parent,
                "strong"
            )

            em = etree.SubElement(
                strong,
                "emphasis"
            )

            em.text = text

        elif bold:

            node = etree.SubElement(
                parent,
                "strong"
            )

            node.text = text

        elif italic:

            node = etree.SubElement(
                parent,
                "emphasis"
            )

            node.text = text

        else:

            self.append_text(
                parent,
                text
            )

    # -----------------------------------------------------

    def add_runs(
        self,
        runs,
        parent=None
    ):

        if not runs:
            return

        if parent is None:

            parent = etree.SubElement(
                self.current_section(),
                "p"
            )

        for run in runs:

            text = run["text"]

            bold = run["bold"]
            italic = run["italic"]

            href = run.get("href")

            run_parent = parent

            if href:

                run_parent = etree.SubElement(
                    parent,
                    "a"
                )

                run_parent.attrib[
                    f"{{{XLINK_NS}}}href"
                ] = href

                self._append_formatted(
                    run_parent,
                    text,
                    bold,
                    italic
                )

                continue

            pos = 0

            for m in self.NOTE_RE.finditer(text):

                start, end = m.span()

                before = text[pos:start]

                note_id = m.group(1)

                if before:

                    self._append_formatted(
                        run_parent,
                        before,
                        bold,
                        italic
                    )

                note_link = etree.SubElement(
                    run_parent,
                    "a"
                )

                note_link.attrib[
                    f"{{{XLINK_NS}}}href"
                ] = f"#n_{note_id}"

                note_link.attrib["type"] = "note"

                note_link.text = f"[{note_id}]"

                pos = end

            remain = text[pos:]

            if remain:

                self._append_formatted(
                    run_parent,
                    remain,
                    bold,
                    italic
                )

    # -----------------------------------------------------

    def add_note(
        self,
        nid,
        text
    ):

        section = etree.SubElement(
            self.notes,
            "section"
        )

        section.attrib["id"] = f"n_{nid}"

        title = etree.SubElement(
            section,
            "title"
        )

        etree.SubElement(
            title,
            "p"
        ).text = nid

        p = etree.SubElement(
            section,
            "p"
        )

        p.text = text

    # -----------------------------------------------------

    def add_binary(
        self,
        image_id,
        data
    ):

        if image_id in self.added_images:
            return

        mime = mimetypes.guess_type(
            image_id
        )[0]

        if mime not in (
            "image/jpeg",
            "image/png"
        ):
            mime = "image/jpeg"

        binary = etree.SubElement(
            self.root,
            "binary"
        )

        binary.attrib["id"] = image_id

        binary.attrib[
            "content-type"
        ] = mime

        binary.text = base64.b64encode(
            data
        ).decode("ascii")

        self.added_images.add(image_id)

    # -----------------------------------------------------

    def add_image(
        self,
        image_id,
        data
    ):

        self.add_binary(
            image_id,
            data
        )

        p = etree.SubElement(
            self.current_section(),
            "p"
        )

        img = etree.SubElement(
            p,
            "image"
        )

        img.attrib[
            f"{{{XLINK_NS}}}href"
        ] = f"#{image_id}"

    # -----------------------------------------------------

    def set_cover(
        self,
        image_id,
        data
    ):

        self.add_binary(
            image_id,
            data
        )

        ti = self.description.find(
            "title-info"
        )

        cover = etree.SubElement(
            ti,
            "coverpage"
        )

        img = etree.SubElement(
            cover,
            "image"
        )

        img.attrib[
            f"{{{XLINK_NS}}}href"
        ] = f"#{image_id}"

    # -----------------------------------------------------

    def add_table(self, rows):

        if not rows:
            return

        table = etree.SubElement(
            self.current_section(),
            "table"
        )

        for row in rows:

            tr = etree.SubElement(
                table,
                "tr"
            )

            for cell_paragraphs in row:

                td = etree.SubElement(
                    tr,
                    "td"
                )

                for item in cell_paragraphs:
                    p = etree.SubElement(td, "p")
                    self.add_runs(item.get("runs"), parent=p)

                    for image_id, data in item.get("images", []):
                        if image_id.lower() == "cover.jpg":
                            continue
                        self.add_image(image_id, data)

    # -----------------------------------------------------

    def save(self, path: Path):

        xml = etree.tostring(
            self.root,
            encoding="utf-8",
            pretty_print=True,
            xml_declaration=True
        )

        path.write_bytes(xml)


# =========================================================
# CONVERTER
# =========================================================

class Converter:

    HEADINGS = {
        "Heading1": 1,
        "Heading2": 2,
        "Heading3": 3,
        "Heading4": 4,
        "Heading5": 5,
        "Heading6": 6,
    }

    NOTE_LINE = re.compile(
        r"^\[(\d+)\]\s*(.+)$"
    )

    def __init__(self, folder):

        self.folder = Path(folder)

        self.docx = (
            self.folder / "book.docx"
        )

        self.meta = Metadata(
            self.folder / "metadata.txt"
        )

        self.cover = (
            self.folder / "cover.jpg"
        )

        self.doc = DocxParser(
            self.docx
        )

        self.fb2 = FB2Builder(
            self.meta
        )

    # -----------------------------------------------------

    def process_cover(self):

        if self.cover.exists():

            self.fb2.set_cover(
                "cover.jpg",
                self.cover.read_bytes()
            )

    # -----------------------------------------------------

    def run(self):

        self.process_cover()

        pending_notes = []

        for kind, block in self.doc.blocks():

            if kind == "p":

                style = (
                    self.doc.paragraph_style(
                        block
                    )
                )

                text = (
                    self.doc.paragraph_text(
                        block
                    )
                )

                runs = (
                    self.doc.paragraph_runs(
                        block
                    )
                )

                images = (
                    self.doc.paragraph_images(
                        block
                    )
                )

                if (
                    not text
                    and
                    not images
                ):
                    continue

                if text:

                    m = self.NOTE_LINE.match(
                        text
                    )

                    if m:

                        pending_notes.append(
                            (
                                m.group(1),
                                m.group(2)
                            )
                        )

                        continue

                if (
                    text
                    and
                    style in self.HEADINGS
                ):

                    self.fb2.open_section(
                        self.HEADINGS[style],
                        text
                    )

                    continue

                if runs:

                    self.fb2.add_runs(
                        runs
                    )

                for image_id, data in images:

                    if (
                        image_id.lower()
                        ==
                        "cover.jpg"
                    ):
                        continue

                    self.fb2.add_image(
                        image_id,
                        data
                    )

            elif kind == "tbl":

                rows = (
                    self.doc.parse_table(
                        block
                    )
                )

                self.fb2.add_table(rows)

        # -------------------------------------------------
        # ДОБАВЛЯЕМ СНОСКИ ПОСЛЕ ОСНОВНОГО ТЕКСТА
        # -------------------------------------------------

        for nid, note_text in pending_notes:

            self.fb2.add_note(
                nid,
                note_text
            )

        output = (
            self.folder / "book.fb2"
        )

        self.fb2.save(output)

        print(
            "Saved:",
            output
        )


# =========================================================
# MAIN
# =========================================================

def main():

    parser = argparse.ArgumentParser()

    parser.add_argument(
        "folder",
        help="Folder with book.docx"
    )

    args = parser.parse_args()

    converter = Converter(
        args.folder
    )

    converter.run()


if __name__ == "__main__":
    main()
