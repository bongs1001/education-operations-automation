from pathlib import Path
from zipfile import ZipFile
import csv
import io
import json
import sys
import xml.etree.ElementTree as ET

if sys.stdout is not None and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def read_csv(path):
    raw = path.read_bytes()
    for encoding in ("utf-16" if raw.startswith((b"\xff\xfe", b"\xfe\xff")) else "utf-8-sig", "cp949"):
        try:
            text = raw.decode(encoding)
            break
        except UnicodeError:
            continue
    else:
        raise ValueError("Unknown encoding")
    delimiter = "\t" if "\t" in text.splitlines()[0] or raw.startswith((b"\xff\xfe", b"\xfe\xff")) else ","
    return encoding, delimiter, list(csv.reader(io.StringIO(text), delimiter=delimiter))


def read_workbook(path):
    with ZipFile(path) as z:
        strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            strings = ["".join(t.text or "" for t in si.iterfind(".//s:t", NS)) for si in root]
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        targets = {r.attrib["Id"]: r.attrib["Target"] for r in rels}
        result = []
        for sheet in wb.find("s:sheets", NS):
            rid = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            target = targets[rid]
            entry = target.lstrip("/") if target.startswith("/") else "xl/" + target
            root = ET.fromstring(z.read(entry))
            cells = {}
            formulas = {}
            for c in root.findall(".//s:sheetData/s:row/s:c", NS):
                v = c.find("s:v", NS)
                value = v.text if v is not None else None
                if c.get("t") == "s" and value is not None:
                    value = strings[int(value)]
                elif c.get("t") == "inlineStr":
                    value = "".join(t.text or "" for t in c.findall(".//s:t", NS))
                f = c.find("s:f", NS)
                if f is not None:
                    formulas[c.get("r")] = f.text
                if value is not None:
                    cells[c.get("r")] = value
            dimension = root.find("s:dimension", NS)
            result.append({"sheet": sheet.get("name"), "dimension": dimension.get("ref") if dimension is not None else "", "cells": cells, "formulas": formulas, "merges": [m.get("ref") for m in root.findall("s:mergeCells/s:mergeCell", NS)]})
        return {"sheets": result, "media_count": len([n for n in z.namelist() if n.startswith("xl/media/")]), "annotation_parts": [n for n in z.namelist() if "comment" in n.lower()]}
