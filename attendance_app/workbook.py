"""Small OOXML cell patcher: unrelated ZIP members remain byte-identical.

The utility deliberately does not round-trip a workbook through a spreadsheet
object model, so screenshots, drawings, comments and unsupported Excel features
survive an attendance update. Formula cells are never replaced by input values.
"""

from copy import deepcopy
from io import BytesIO
from pathlib import Path
import math
import os
import posixpath
import re
import tempfile
from xml.sax.saxutils import escape
import xml.etree.ElementTree as ET
from zipfile import ZipFile

from .ooxml import prune_package, rels_path

SS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"s": SS}
ET.register_namespace("", SS)


def equivalent(left, right):
    if left in (None, "") and right in (None, ""):
        return True
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        return math.isclose(left, right, rel_tol=0, abs_tol=1e-10)
    return left == right


class Workbook:
    def __init__(self, source):
        self.original = source if isinstance(source, bytes) else Path(source).read_bytes()
        with ZipFile(BytesIO(self.original)) as z:
            self.infos = z.infolist()
            if len({i.filename for i in self.infos}) != len(self.infos):
                raise ValueError("중복 항목이 있는 Excel 파일입니다.")
            self.parts = {i.filename: z.read(i.filename) for i in self.infos}
        self.strings = []
        if "xl/sharedStrings.xml" in self.parts:
            root = ET.fromstring(self.parts["xl/sharedStrings.xml"])
            self.strings = ["".join(si.itertext()) for si in root]
        wb = ET.fromstring(self.parts["xl/workbook.xml"])
        rels = ET.fromstring(self.parts["xl/_rels/workbook.xml.rels"])
        targets = {r.get("Id"): r.get("Target") for r in rels}
        self.sheets = {}
        for sheet in wb.find("s:sheets", NS):
            target = targets[sheet.get(f"{{{REL}}}id")]
            self.sheets[sheet.get("name")] = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
        self.style_cache = {}

    def tree(self, sheet):
        if sheet not in self.sheets:
            raise ValueError(f"'{sheet}' 시트가 없습니다.")
        return ET.fromstring(self.parts[self.sheets[sheet]])

    def cells(self, sheet):
        return {c.get("r"): c for c in self.tree(sheet).findall("s:sheetData/s:row/s:c", NS)}

    def value(self, cell):
        if cell is None:
            return None
        formula = cell.find("s:f", NS)
        if formula is not None:
            return {"formula": formula.text or "", "attributes": formula.attrib}
        if cell.get("t") == "inlineStr":
            return "".join(t.text or "" for t in cell.findall(".//s:t", NS))
        v = cell.find("s:v", NS)
        if v is None or v.text is None:
            return None
        if cell.get("t") == "s":
            return self.strings[int(v.text)]
        if cell.get("t") in ("str", "e", "d"):
            return v.text
        if cell.get("t") == "b":
            return bool(int(v.text))
        return float(v.text)

    def values(self, sheet):
        return {ref: self.value(cell) for ref, cell in self.cells(sheet).items()}

    def roster(self):
        values = self.values("출결정보")
        result = {}
        for ref, name in values.items():
            if not re.fullmatch(r"D\d+", ref):
                continue
            row = int(ref[1:])
            number = values.get(f"B{row}")
            if isinstance(number, (int, float)) and number > 0 and number == int(number) and name:
                result[row] = str(name).strip()
        if not result:
            raise ValueError("출결정보 B열 순번·D열 성명에서 수강생을 찾지 못했습니다.")
        if len(set(result.values())) != len(result):
            raise ValueError("출석부에 동명이인이 있습니다. 이름을 구분한 뒤 진행하세요.")
        return result

    def highlighted_cells(self, sheet):
        styles = ET.fromstring(self.parts["xl/styles.xml"])
        fills = styles.find("s:fills", NS)
        xfs = styles.find("s:cellXfs", NS)
        highlighted_styles = set()
        for index, xf in enumerate(xfs):
            pattern = fills[int(xf.get("fillId", "0"))].find("s:patternFill", NS)
            color = pattern.find("s:fgColor", NS) if pattern is not None else None
            if pattern is not None and pattern.get("patternType") == "solid" and color is not None and color.get("rgb", "").upper().endswith("FFB366"):
                highlighted_styles.add(index)
        return {ref for ref, cell in self.cells(sheet).items() if int(cell.get("s", "0")) in highlighted_styles}

    def _append_style_element(self, group, element):
        xml = self.parts["xl/styles.xml"].decode("utf-8")
        pattern = rf'<{group}\b[^>]*?(?:/>|>.*?</{group}>)'
        match = re.search(pattern, xml, re.S)
        if match:
            # Elements inherit the stylesheet namespace; parse within a wrapper.
            root = ET.fromstring(f'<styleSheet xmlns="{SS}">{match[0]}</styleSheet>')[0]
            index = len(root)
            root.append(element)
            root.set("count", str(len(root)))
            replacement = ET.tostring(root, encoding="unicode")
            xml = xml[:match.start()] + replacement + xml[match.end():]
        else:
            if group != "numFmts":
                raise ValueError(f"Excel 스타일 '{group}'이 없습니다.")
            index = 0
            root = ET.Element(f"{{{SS}}}{group}", count="1")
            root.append(element)
            xml = re.sub(r'(<styleSheet\b[^>]*>)', lambda m: m[0] + ET.tostring(root, encoding="unicode"), xml, count=1)
        self.parts["xl/styles.xml"] = xml.encode("utf-8")
        return index

    def style(self, base, kind, highlight):
        key = (base, kind, highlight)
        if key in self.style_cache:
            return self.style_cache[key]
        if not kind and not highlight:
            return base
        root = ET.fromstring(self.parts["xl/styles.xml"])
        xfs = root.find("s:cellXfs", NS)
        xf = deepcopy(xfs[base])
        if kind in ("time", "duration"):
            fmt = "hh:mm:ss" if kind == "time" else "[h]:mm:ss"
            formats = root.findall("s:numFmts/s:numFmt", NS)
            existing = next((int(n.get("numFmtId")) for n in formats if n.get("formatCode") == fmt), None)
            if existing is None:
                existing = max([163] + [int(n.get("numFmtId")) for n in formats]) + 1
                self._append_style_element("numFmts", ET.Element(f"{{{SS}}}numFmt", numFmtId=str(existing), formatCode=fmt))
            xf.set("numFmtId", str(existing))
            xf.set("applyNumberFormat", "1")
        if highlight:
            fills = root.find("s:fills", NS)
            fill_id = None
            for i, fill in enumerate(fills):
                pattern = fill.find("s:patternFill", NS)
                color = pattern.find("s:fgColor", NS) if pattern is not None else None
                if pattern is not None and pattern.get("patternType") == "solid" and color is not None and color.get("rgb") == "00FFB366":
                    fill_id = i
            if fill_id is None:
                fill = ET.Element(f"{{{SS}}}fill")
                pattern = ET.SubElement(fill, f"{{{SS}}}patternFill", patternType="solid")
                ET.SubElement(pattern, f"{{{SS}}}fgColor", rgb="00FFB366")
                ET.SubElement(pattern, f"{{{SS}}}bgColor", rgb="00FFB366")
                fill_id = self._append_style_element("fills", fill)
            xf.set("fillId", str(fill_id))
            xf.set("applyFill", "1")
        style_id = self._append_style_element("cellXfs", xf)
        self.style_cache[key] = style_id
        return style_id

    def patch(self, sheet, changes, highlight=False):
        """changes: ref -> (typed value, format kind). Existing row required."""
        path = self.sheets[sheet]
        xml = self.parts[path].decode("utf-8")
        cells = self.cells(sheet)
        for ref, (value, kind) in changes.items():
            if not re.fullmatch(r"[A-Z]+[1-9]\d*", ref):
                raise ValueError("잘못된 셀 주소입니다.")
            old = cells.get(ref)
            if old is not None and old.find("s:f", NS) is not None:
                raise ValueError(f"{ref}에는 수식이 있어 덮어쓸 수 없습니다.")
            cell = deepcopy(old) if old is not None else ET.Element(f"{{{SS}}}c", r=ref)
            for child in list(cell):
                if child.tag in (f"{{{SS}}}v", f"{{{SS}}}is"):
                    cell.remove(child)
            cell.attrib.pop("t", None)
            base = int(cell.get("s", "0"))
            style_id = self.style(base, kind, highlight)
            if style_id or "s" in cell.attrib:
                cell.set("s", str(style_id))
            if isinstance(value, str):
                cell.set("t", "inlineStr")
                text = ET.SubElement(ET.SubElement(cell, f"{{{SS}}}is"), f"{{{SS}}}t")
                text.text = value
                text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            elif value is not None:
                if not isinstance(value, (float, int)) or not math.isfinite(value):
                    raise ValueError("지원하지 않는 셀 값입니다.")
                ET.SubElement(cell, f"{{{SS}}}v").text = format(value, ".15g")
            replacement = ET.tostring(cell, encoding="unicode")
            pattern = rf'<c\b(?=[^>]*\br="{ref}")[^>]*?(?:/>|>.*?</c>)'
            xml, count = re.subn(pattern, lambda _: replacement, xml, count=1, flags=re.S)
            if not count:
                row_number = int(re.search(r"\d+", ref)[0])
                row_pattern = rf'<row\b(?=[^>]*\br="{row_number}")[^>]*>.*?</row>'
                row_match = re.search(row_pattern, xml, re.S)
                if not row_match:
                    raise ValueError(f"{ref}가 속한 행이 없습니다. 원본 양식을 확인하세요.")
                row_xml = row_match[0]
                column_number = lambda name: sum((ord(ch) - 64) * 26 ** power for power, ch in enumerate(reversed(re.match(r"[A-Z]+", name)[0])))
                offset = row_xml.rfind("</row>")
                for next_cell in re.finditer(r'<c\b[^>]*\br="([A-Z]+\d+)"', row_xml):
                    if column_number(next_cell[1]) > column_number(ref):
                        offset = next_cell.start()
                        break
                row_xml = row_xml[:offset] + replacement + row_xml[offset:]
                xml = xml[:row_match.start()] + row_xml + xml[row_match.end():]
        self.parts[path] = xml.encode("utf-8")

    def refresh_attendance_counts(self):
        """Refresh the template's simple COUNTIF caches; preserve formula text."""
        path = self.sheets["출결정보"]
        xml = self.parts[path].decode("utf-8")
        cells = self.cells("출결정보")
        values = {ref: self.value(c) for ref, c in cells.items()}
        for ref, cell in cells.items():
            f = cell.find("s:f", NS)
            if f is None:
                continue
            match = re.fullmatch(r"COUNTIF\(\$?([KLM])\$?(\d+):\$?\1\$?(\d+),\$?([A-Z]+)\$?(\d+)\)", f.text or "", re.I)
            if not match:
                continue
            col, first, last, criterion_col, criterion_row = match.groups()
            criterion = values.get(criterion_col + criterion_row)
            count = sum(str(values.get(f"{col}{n}", "") or "").lower() == str(criterion or "").lower() for n in range(int(first), int(last) + 1))
            replacement = deepcopy(cell)
            replacement.attrib.pop("t", None)
            v = replacement.find("s:v", NS)
            if v is None:
                v = ET.SubElement(replacement, f"{{{SS}}}v")
            v.text = str(count)
            pattern = rf'<c\b(?=[^>]*\br="{ref}")[^>]*?(?:/>|>.*?</c>)'
            xml = re.sub(pattern, lambda _: ET.tostring(replacement, encoding="unicode"), xml, count=1, flags=re.S)
        self.parts[path] = xml.encode("utf-8")
        wb = self.parts["xl/workbook.xml"].decode("utf-8")
        def update_calc(m):
            tag = m[0]
            for name, value in (("fullCalcOnLoad", "1"), ("forceFullCalc", "1")):
                tag = re.sub(rf'\s{name}="[^"]*"', "", tag)
                tag = tag[:-2] + f' {name}="{value}"/>'
            return tag
        if re.search(r'<calcPr\b[^>]*/>', wb):
            wb = re.sub(r'<calcPr\b[^>]*/>', update_calc, wb, count=1)
        elif "<calcPr" not in wb:
            wb = wb.replace("</workbook>", '<calcPr fullCalcOnLoad="1" forceFullCalc="1"/></workbook>')
        self.parts["xl/workbook.xml"] = wb.encode("utf-8")

    def prepare_web_export(self):
        """Match the reference web download; never used by original-file updates."""
        for part in self.sheets.values():
            xml = self.parts[part].decode("utf-8")
            # The reference serializer makes Excel's base column width explicit.
            xml = re.sub(r'<sheetFormatPr\b[^>]*/>', lambda m: m[0] if 'baseColWidth=' in m[0] else m[0][:-2] + ' baseColWidth="8"/>', xml)
            self.parts[part] = xml.encode("utf-8")
            relationship = rels_path(part)
            if relationship not in self.parts:
                continue
            relations = ET.fromstring(self.parts[relationship])
            removed = {r.get("Id") for r in relations if r.get("Type", "").endswith("/drawing")}
            if not removed:
                continue
            relxml = self.parts[relationship].decode("utf-8")
            relxml = re.sub(r'<Relationship\b[^>]*/>', lambda m: "" if ET.fromstring(m[0]).get("Id") in removed else m[0], relxml)
            if list(ET.fromstring(relxml)):
                self.parts[relationship] = relxml.encode("utf-8")
            else:
                del self.parts[relationship]
            xml = self.parts[part].decode("utf-8")
            for rid in removed:
                xml = re.sub(rf'<drawing\b[^>]*r:id="{re.escape(rid)}"[^>]*/>', "", xml)
            self.parts[part] = xml.encode("utf-8")
        self.parts = prune_package(self.parts)

    def to_bytes(self):
        stream = BytesIO()
        with ZipFile(stream, "w") as z:
            for info in self.infos:
                if info.filename in self.parts:
                    z.writestr(info, self.parts[info.filename])
        result = stream.getvalue()
        with ZipFile(BytesIO(result)) as z:
            if z.testzip():
                raise ValueError("생성한 Excel 파일을 검증하지 못했습니다.")
            for name in (self.sheets["출결정보"], "xl/styles.xml", "xl/workbook.xml"):
                ET.fromstring(z.read(name))
        return result


def write_new(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("xb") as stream:
        stream.write(data)


def atomic_replace(path, data, expected_bytes):
    path = Path(path)
    if path.read_bytes() != expected_bytes:
        raise ValueError("처리 중 원본 파일이 변경되었습니다. 다시 분석하세요.")
    fd, temporary = tempfile.mkstemp(prefix=".attendance_", suffix=".xlsx", dir=path.parent)
    try:
        with os.fdopen(fd, "wb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
        if path.read_bytes() != expected_bytes:
            raise ValueError("저장 직전 원본이 변경되었습니다. 다시 분석하세요.")
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)
