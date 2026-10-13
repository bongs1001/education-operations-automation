"""Create dated attendance workbooks and roster-matched template presentations."""

from collections import defaultdict
from copy import deepcopy
from datetime import date, timedelta
from io import BytesIO
import csv
import json
from pathlib import Path
import posixpath
import re
from xml.sax.saxutils import escape
import xml.etree.ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED

from attendance_reference.inspect_samples import read_csv
from .workbook import Workbook, NS, SS, write_new
from .service import unique_folder
from .ooxml import prune_package, rels_path

PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
DRAW = "http://schemas.openxmlformats.org/drawingml/2006/main"
PRES = "http://schemas.openxmlformats.org/presentationml/2006/main"
OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def clean_attendance(source, day):
    book = Workbook(source)
    roster = book.roster()
    values = book.values("출결정보")
    daily_rows = {int(ref[1:]) for ref, number in values.items() if re.fullmatch(r"B\d+", ref) and isinstance(number, (int, float)) and number > 0 and number == int(number)}
    changes = {f"{col}{row}": (None, "") for row in daily_rows for col in "EFGHIJKLMNOP"}
    changes["B4"] = (f"일자 : {day:%Y.%m.%d}", "")
    book.patch("출결정보", changes)
    if "증빙자료" in book.sheets:
        values = book.values("증빙자료")
        book.patch("증빙자료", {ref: (None, "") for ref, value in values.items() if ref != "B2" and value is not None and not isinstance(value, dict)})
    # Clearing daily evidence is intentional only for NEW template generation.
    for sheet in ("출결정보", "증빙자료"):
        if sheet not in book.sheets:
            continue
        part = book.sheets[sheet]
        xml = book.parts[part].decode("utf-8")
        relationship = rels_path(part)
        if relationship in book.parts:
            relxml = book.parts[relationship].decode("utf-8")
            removed_ids = []
            def remove_daily_rel(m):
                if any(f'/{kind}"' in m[0] for kind in ("drawing", "vmlDrawing", "comments", "threadedComment", "threadedComments")):
                    removed_ids.append(re.search(r'\bId="([^"]+)"', m[0])[1])
                    return ""
                return m[0]
            relxml = re.sub(r'<Relationship\b[^>]*/>', remove_daily_rel, relxml)
            for rid in removed_ids:
                xml = re.sub(rf'<(?:drawing|legacyDrawing|legacyDrawingHF)\b[^>]*r:id="{re.escape(rid)}"[^>]*/>', "", xml)
            book.parts[relationship] = relxml.encode("utf-8")
        book.parts[part] = xml.encode("utf-8")
    book.refresh_attendance_counts()
    book.parts = prune_package(book.parts)
    return book.to_bytes()


def replace_text_nodes(xml, values):
    """Replace a complete text sequence, retaining each existing run's styling."""
    items = iter(values)
    def replace(m):
        try:
            value = next(items)
        except StopIteration:
            raise ValueError("PPT 텍스트 구조가 예상한 양식과 다릅니다.")
        return m[1] + escape(str(value)) + m[3]
    xml = re.sub(r'(<a:t\b[^>]*>)(.*?)(</a:t>)', replace, xml, flags=re.S)
    if next(items, None) is not None:
        raise ValueError("PPT 텍스트 개수가 예상한 양식과 다릅니다.")
    return xml


def text_values(xml):
    return [t.text or "" for t in ET.fromstring(xml).iter(f"{{{DRAW}}}t")]


def set_table_cell(xml, value):
    matches = list(re.finditer(r'(<a:t\b[^>]*>)(.*?)(</a:t>)', xml, re.S))
    if matches:
        values = [value] + [""] * (len(matches) - 1)
        return replace_text_nodes(xml, values)
    if value == "":
        return xml
    # Blank capacity cells retain their cell and paragraph formatting.
    run = f'<a:r><a:rPr lang="ko-KR" sz="1800"/><a:t>{escape(str(value))}</a:t></a:r>'
    return xml.replace("</a:p>", run + "</a:p>", 1)


def roster_slide(xml, roster, day):
    xml, dates = re.subn(r'(<a:t\b[^>]*>)\d{4}-\d{2}-\d{2}(</a:t>)', lambda m: m[1] + day.isoformat() + m[2], xml)
    if dates != 1:
        raise ValueError("산출물 2번 슬라이드에서 날짜 자리 1개를 찾지 못했습니다.")
    tables = list(re.finditer(r'<a:tbl\b[^>]*>.*?</a:tbl>', xml, re.S))
    if len(tables) != 1:
        raise ValueError("산출물 양식의 명단 표가 예상과 다릅니다.")
    table = tables[0][0]
    rows = list(re.finditer(r'<a:tr\b[^>]*>.*?</a:tr>', table, re.S))
    if len(rows) != 13:
        raise ValueError("지원 양식은 명단 표가 머리글 1행·학생 12행이어야 합니다.")
    if len(roster) > 36:
        raise ValueError("이 PPT 양식은 명단 표에 최대 36명까지 지원합니다.")
    row_cursor = -1
    def update_row(match):
        nonlocal row_cursor
        row_cursor += 1
        if row_cursor == 0:
            return match[0]
        cell_cursor = -1
        def update_cell(m):
            nonlocal cell_cursor
            cell_cursor += 1
            slot = (cell_cursor // 2) * 12 + row_cursor - 1
            value = "" if slot >= len(roster) else roster[slot][cell_cursor % 2]
            return set_table_cell(m[0], value)
        return re.sub(r'<a:tc\b[^>]*>.*?</a:tc>', update_cell, match[0], flags=re.S)
    new_table = re.sub(r'<a:tr\b[^>]*>.*?</a:tr>', update_row, table, flags=re.S)
    return xml[:tables[0].start()] + new_table + xml[tables[0].end():]


def generate_presentation(template, roster, day):
    with ZipFile(template) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    presentation = parts["ppt/presentation.xml"].decode("utf-8")
    original_rels = ET.fromstring(parts["ppt/_rels/presentation.xml.rels"])
    targets = {r.get("Id"): posixpath.normpath("ppt/" + r.get("Target")) for r in original_rels if r.get("Type", "").endswith("/slide")}
    tree = ET.fromstring(presentation)
    ids = tree.find(f"{{{PRES}}}sldIdLst")
    ordered = [targets[n.get(f"{{{OFFICE_REL}}}id")] for n in ids]
    if len(ordered) < 8 or (len(ordered) - 4) % 4:
        raise ValueError("산출물 양식은 앞 4장과 학생별 4장으로 구성되어야 합니다.")
    blocks = {}
    for offset in range(4, len(ordered), 4):
        header = text_values(parts[ordered[offset]])
        match = re.fullmatch(r"\s*\d+\s+(.+?)\s*", "".join(header))
        if not match:
            raise ValueError("학생 표지 구조가 선택한 산출물 양식과 다릅니다.")
        name = match[1]
        if name in blocks:
            raise ValueError("PPT 양식의 학생 이름이 중복되어 있습니다.")
        blocks[name] = ordered[offset:offset + 4]
    selected = [(path, None, None) for path in ordered[:4]]
    for number, name in roster:
        # The input is a blank evidence template. A new learner uses the same blank layout.
        block = blocks.get(name, ordered[4:8])
        selected.extend((path, (number, name), index) for index, path in enumerate(block))
    out = dict(parts)
    all_slide_paths = set(ordered)
    for path in all_slide_paths:
        out.pop(path, None)
        out.pop(rels_path(path), None)
    relxml = parts["ppt/_rels/presentation.xml.rels"].decode("utf-8")
    relxml = re.sub(r'<Relationship\b[^>]*Type="[^"]*/slide"[^>]*/>', "", relxml)
    new_relations, new_ids, overrides = [], [], []
    for index, (source, student, block_index) in enumerate(selected, 1):
        path = f"ppt/slides/slide{index}.xml"
        xml = parts[source].decode("utf-8")
        if index == 2:
            xml = roster_slide(xml, roster, day)
        elif student is not None:
            number, name = student
            old = text_values(xml)
            if block_index == 0:
                new = [f"{number:02d} {name}"] if len(old) == 1 else [f"{number:02d} ", name] + [""] * (len(old) - 2)
            else:
                # Supports a label divided across several rich text runs.
                new = [f"{number}-{block_index}"] + [""] * (len(old) - 1)
            xml = replace_text_nodes(xml, new)
        ET.fromstring(xml)
        out[path] = xml.encode("utf-8")
        source_rels = rels_path(source)
        if source_rels in parts:
            relationships = parts[source_rels]
            if any(r.get("Type", "").endswith(("/notesSlide", "/slide", "/comments")) for r in ET.fromstring(relationships)):
                raise ValueError("현재 산출물 생성은 슬라이드별 노트·댓글·다른 슬라이드 링크가 없는 제공 양식을 지원합니다.")
            out[rels_path(path)] = relationships
        rid = f"rIdAttendance{index}"
        new_relations.append(f'<Relationship Id="{rid}" Type="{OFFICE_REL}/slide" Target="slides/slide{index}.xml"/>')
        new_ids.append(f'<p:sldId id="{1000 + index}" r:id="{rid}"/>')
        overrides.append(f'<Override PartName="/{path}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>')
    relxml = relxml.replace("</Relationships>", "".join(new_relations) + "</Relationships>")
    presentation = re.sub(r'<p:sldIdLst>.*?</p:sldIdLst>', lambda _: "<p:sldIdLst>" + "".join(new_ids) + "</p:sldIdLst>", presentation, count=1, flags=re.S)
    out["ppt/presentation.xml"] = presentation.encode("utf-8")
    out["ppt/_rels/presentation.xml.rels"] = relxml.encode("utf-8")
    content_types = parts["[Content_Types].xml"].decode("utf-8")
    content_types = re.sub(r'<Override\b[^>]*PartName="/ppt/slides/slide\d+\.xml"[^>]*/>', "", content_types)
    out["[Content_Types].xml"] = content_types.replace("</Types>", "".join(overrides) + "</Types>").encode("utf-8")
    if "docProps/app.xml" in out:
        app = out["docProps/app.xml"].decode("utf-8")
        app = re.sub(r'<Slides>\d+</Slides>', f"<Slides>{len(selected)}</Slides>", app)
        out["docProps/app.xml"] = app.encode("utf-8")
    out = prune_package(out)
    stream = BytesIO()
    with ZipFile(stream, "w", ZIP_DEFLATED) as z:
        for name, content in out.items():
            z.writestr(name, content)
    return stream.getvalue()


HOLIDAY_NAMES = {"추석", "설날", "연휴", "휴강", "대체공휴일", "공휴일", "광복절", "개천절", "한글날", "크리스마스", "성탄절"}


def presentation_eligible(event):
    teacher = re.sub(r"\s+", "", event.get("instructor", "")).casefold()
    unknown = teacher in {"", "강사미정", "강사미배정", "미정", "미배정", "미확정", "tbd", "tba", "none", "null", "-"} or teacher.startswith("강사미정(")
    return event.get("event_type", "").strip().casefold() in {"class", "수업"} and not unknown


def filter_events(events, kind):
    if kind not in ("attendance", "presentation", "both"):
        raise ValueError("생성할 파일 종류를 확인하세요.")
    included, excluded = [], []
    for raw in events:
        event = dict(raw)
        reason = None
        if event.get("file_required", "yes").strip().lower() != "yes":
            reason = "파일 생성 제외 일정"
        elif event.get("event_name", "").strip() in HOLIDAY_NAMES:
            reason = "휴강·공휴일"
        elif kind == "presentation" and not presentation_eligible(event):
            reason = "수업 외 일정 또는 강사 미정"
        event["instructor"] = event.get("instructor", "").strip() or "강사미정"
        if reason:
            excluded.append({**event, "excluded_reason": reason})
        else:
            included.append(event)
    return included, excluded


def load_schedule(path, first, last, kind="attendance"):
    _, _, rows = read_csv(Path(path))
    required = {"date", "event_name", "instructor", "file_required"}
    if kind == "presentation":
        required.add("event_type")
    if not rows or not required.issubset(rows[0]):
        raise ValueError("일정 CSV 필수 열을 확인하세요: " + ", ".join(sorted(required)))
    events = []
    for row in rows[1:]:
        if not row:
            continue
        if len(row) != len(rows[0]):
            raise ValueError("일정 CSV의 열 개수가 일치하지 않습니다.")
        item = dict(zip(rows[0], row))
        day = date.fromisoformat(item["date"])
        if not first <= day <= last:
            continue
        events.append(item)
    return filter_events(events, kind)


def safe_filename(value):
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value).strip().rstrip(". ") or "미정"


def generate_files(attendance_template, ppt_template, events, output_parent, make_zip=True, kind="both"):
    if not events:
        raise ValueError("생성할 일정을 선택하세요.")
    events, excluded = filter_events(events, kind)
    if not events:
        raise ValueError("생성할 일정이 없습니다. 산출물은 강사가 지정된 수업 일정만 생성합니다." if kind == "presentation" else "생성할 일정이 없습니다.")
    create_attendance = kind in ("attendance", "both")
    ppt_events = [event for event in events if presentation_eligible(event)] if kind in ("presentation", "both") else []
    if ppt_events and not ppt_template:
        if kind == "presentation":
            raise ValueError("산출물 PPT 양식을 선택하세요.")
        ppt_events = []
    book = Workbook(attendance_template)
    values = book.values("출결정보")
    roster = [(int(values[f"B{row}"]), name) for row, name in book.roster().items()]
    # Validate PPT contract before creating a batch directory.
    first_day = min(date.fromisoformat(event["date"]) for event in events)
    first_ppt_day = min((date.fromisoformat(event["date"]) for event in ppt_events), default=None)
    first_ppt = generate_presentation(ppt_template, roster, first_ppt_day) if ppt_events else None
    first_xlsx = clean_attendance(book.original, first_day) if create_attendance else None
    days = sorted({date.fromisoformat(e["date"]) for e in events})
    batch = unique_folder(output_parent, {"attendance": "출석부_생성", "presentation": "산출물_생성", "both": "출석부_산출물"}[kind])
    outputs = []
    for day in days:
        day_ppt_events = [e for e in ppt_events if e["date"] == day.isoformat()]
        monday = day - timedelta(days=day.weekday())
        week_number = (day.day + day.replace(day=1).weekday() - 1) // 7 + 1
        span = f"{monday:%m%d}~{monday + timedelta(days=4):%m%d}"
        week = f"{day.month}월 {week_number}주차 ({span})"
        month = f"{day:%y}년 {day:%m}월"
        if create_attendance:
            destination = batch / month / "출석부 - 강사용" / week / f"{day.month}월 {day.day:02d}일" / f"{day:%m%d} 출석부.xlsx"
            write_new(destination, first_xlsx if day == first_day else clean_attendance(book.original, day))
            outputs.append(str(destination.relative_to(batch)))
        if day_ppt_events:
            teachers = sorted({e["instructor"].strip() for e in day_ppt_events})
            ppt = first_ppt if day == first_ppt_day else generate_presentation(ppt_template, roster, day)
            used_names = set()
            for teacher in teachers:
                filename = f"{day:%y%m%d}_산출물_{safe_filename(teacher)}.pptx"
                if filename in used_names:
                    raise ValueError("강사명에서 파일명으로 바꿀 때 중복이 발생했습니다.")
                used_names.add(filename)
                destination = batch / month / "산출물 - 수강생용" / week / f"{day.month}월 {day.day:02d}일" / filename
                write_new(destination, ppt)
                outputs.append(str(destination.relative_to(batch)))
    manifest = {"kind": kind, "student_count": len(roster), "ppt_slides": 4 + len(roster) * 4 if ppt_events else 0, "dates": [d.isoformat() for d in days], "events": events, "excluded_events": excluded, "files": outputs}
    write_new(batch / "생성내역.json", json.dumps(manifest, ensure_ascii=False, indent=2).encode("utf-8"))
    archive = None
    if make_zip:
        archive = batch.with_suffix(".zip")
        with ZipFile(archive, "x", ZIP_DEFLATED) as z:
            for path in batch.rglob("*"):
                if path.is_file():
                    z.write(path, str(path.relative_to(batch)))
    return {"folder": str(batch), "zip": str(archive) if archive else None, **manifest}
