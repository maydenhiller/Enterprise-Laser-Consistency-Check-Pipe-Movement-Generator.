"""Write cells into an .xlsx without opening it in a spreadsheet library.

Enterprise's laser workbooks hold about 25 charts and a set of hidden calculation
sheets. openpyxl drops charts when it saves a file it did not create, so cells are
written straight into the sheet XML and every other part of the package is copied
through byte for byte. Adapted from the Enterprise Dig File Generator's XlsmPatcher.
"""
from __future__ import annotations

import io
import re
import zipfile
from typing import Optional
from xml.sax.saxutils import escape


class XlsxPatcher:
    def __init__(self, data: bytes):
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            self.names = list(archive.namelist())
            self.parts = {name: archive.read(name) for name in self.names}
            self.infos = {info.filename: info for info in archive.infolist()}
        self._sheet_paths = self._map_sheets()

    # -- structure ------------------------------------------------------
    def _text(self, path: str) -> str:
        return self.parts[path].decode("utf-8")

    def _map_sheets(self) -> dict:
        workbook = self._text("xl/workbook.xml")
        rels = self._text("xl/_rels/workbook.xml.rels")
        targets = {
            m.group(1): m.group(2)
            for m in re.finditer(r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', rels)
        }
        mapping = {}
        for match in re.finditer(r"<sheet\b[^>]*/?>", workbook):
            tag = match.group(0)
            name = re.search(r'name="([^"]*)"', tag)
            rid = re.search(r'r:id="([^"]*)"', tag)
            if not name or not rid or rid.group(1) not in targets:
                continue
            target = targets[rid.group(1)]
            path = target.lstrip("/") if target.startswith(("xl/", "/xl/")) else "xl/" + target
            mapping[_unescape(name.group(1))] = path.replace("/./", "/")
        return mapping

    def sheet_names(self) -> list:
        return list(self._sheet_paths)

    def sheet_xml_containing(self, needle: str) -> list:
        """Names of sheets whose XML contains ``needle`` (e.g. '#REF!')."""
        return [name for name, path in self._sheet_paths.items()
                if needle in self._text(path)]

    # -- cells ----------------------------------------------------------
    def set_values(self, sheet_name: str, values: dict) -> list:
        """Write {'A8': value}. None clears the cell (keeps its formatting).

        Text is always written as text, never as a formula. Returns cells that
        could not be placed (row missing from the sheet).
        """
        path = self._sheet_paths[sheet_name]
        xml = self._text(path)
        missing = []
        for ref, value in values.items():
            xml, ok = _write_cell(xml, ref, value)
            if not ok:
                missing.append(ref)
        self.parts[path] = xml.encode("utf-8")
        return missing

    def set_cached(self, sheet_name: str, values: dict) -> None:
        """Refresh what a formula cell last computed, so previews that do not
        recalculate (phones, mail clients) show the right numbers."""
        path = self._sheet_paths[sheet_name]
        xml = self._text(path)
        for ref, value in values.items():
            xml = _set_cached_value(xml, ref, value)
        self.parts[path] = xml.encode("utf-8")

    def force_full_recalc(self) -> None:
        path = "xl/workbook.xml"
        xml = self._text(path)
        match = re.search(r"<calcPr\b[^>]*/?>", xml)
        if match:
            tag = match.group(0)
            if "fullCalcOnLoad" not in tag:
                updated = tag.rstrip("/>").rstrip() + ' fullCalcOnLoad="1"/>'
                xml = xml[: match.start()] + updated + xml[match.end():]
        else:
            xml = xml.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
        self.parts[path] = xml.encode("utf-8")

    def to_bytes(self) -> bytes:
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
            for name in self.names:
                info = self.infos[name]
                new = zipfile.ZipInfo(name, date_time=info.date_time)
                new.external_attr = info.external_attr
                new.compress_type = zipfile.ZIP_DEFLATED
                archive.writestr(new, self.parts[name])
        return buffer.getvalue()


# ---------------------------------------------------------------------------
def _unescape(text: str) -> str:
    return (text.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
            .replace("&apos;", "'").replace("&amp;", "&"))


def _cell_pattern(ref: str) -> re.Pattern:
    return re.compile(rf'<c r="{ref}"(?P<attrs>[^>]*?)(?:/>|>(?P<body>.*?)</c>)', re.S)


def _style_of(attrs: str) -> str:
    match = re.search(r'\ss="(\d+)"', attrs or "")
    return f' s="{match.group(1)}"' if match else ""


def _render(ref: str, style: str, value) -> str:
    if value is None:
        return f'<c r="{ref}"{style}/>'
    if isinstance(value, bool):
        return f'<c r="{ref}"{style} t="b"><v>{int(value)}</v></c>'
    if isinstance(value, (int, float)):
        return f'<c r="{ref}"{style}><v>{value!r}</v></c>'
    text = escape(str(value))
    return f'<c r="{ref}"{style} t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'


def _write_cell(xml: str, ref: str, value) -> tuple:
    match = _cell_pattern(ref).search(xml)
    if match:
        body = match.group("body") or ""
        if "<f" in body:
            raise ValueError(f"Refusing to overwrite formula cell {ref}")
        new = _render(ref, _style_of(match.group("attrs")), value)
        return xml[: match.start()] + new + xml[match.end():], True
    if value is None:
        return xml, True  # nothing there to clear
    inserted = _insert_cell(xml, ref, value)
    return (inserted, True) if inserted is not None else (xml, False)


def _set_cached_value(xml: str, ref: str, value) -> str:
    match = _cell_pattern(ref).search(xml)
    if not match or not match.group("body"):
        return xml
    formula = re.search(r"<f\b[^>]*?(?:/>|>.*?</f>)", match.group("body"), re.S)
    if not formula:
        return xml
    attrs = re.sub(r'\st="[^"]*"', "", match.group("attrs") or "")
    if value is None or value == "":
        kind, text = ' t="str"', ""
    elif isinstance(value, (int, float)) and not isinstance(value, bool):
        kind, text = "", repr(value)
    else:
        kind, text = ' t="str"', escape(str(value))
    cell = f'<c r="{ref}"{attrs}{kind}>{formula.group(0)}<v>{text}</v></c>'
    return xml[: match.start()] + cell + xml[match.end():]


_REF_RE = re.compile(r"([A-Z]+)(\d+)")


def column_index(letters: str) -> int:
    index = 0
    for ch in letters:
        index = index * 26 + (ord(ch) - 64)
    return index


def column_letters(index: int) -> str:
    letters = ""
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _insert_cell(xml: str, ref: str, value) -> Optional[str]:
    parsed = _REF_RE.fullmatch(ref)
    if not parsed:
        return None
    target, number = column_index(parsed.group(1)), parsed.group(2)
    row = re.search(rf'<row[^>]*\br="{number}"[^>]*>(.*?)</row>', xml, re.S)
    if not row:
        return None
    body, style, insert_at = row.group(1), "", len(row.group(1))
    for cell in re.finditer(r'<c r="([A-Z]+)\d+"([^>]*?)(?:/>|>.*?</c>)', body, re.S):
        if column_index(cell.group(1)) < target:
            style = _style_of(cell.group(2)) or style
        else:
            insert_at = cell.start()
            break
    new_body = body[:insert_at] + _render(ref, style, value) + body[insert_at:]
    return xml[: row.start(1)] + new_body + xml[row.end(1):]
