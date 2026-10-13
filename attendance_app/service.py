"""Attendance analysis, export and optimistic original-file updates."""

from datetime import datetime
from html import escape
import hashlib
import json
from pathlib import Path
import re

from attendance_reference.analyze_samples import analyze, parse_duration, parse_timestamp
from .workbook import Workbook, equivalent, atomic_replace, write_new


def unique_folder(parent, label):
    folder = Path(parent) / f"{label}_{datetime.now():%Y%m%d_%H%M%S_%f}"
    folder.mkdir(parents=True, exist_ok=False)
    return folder


def analyze_files(workbook, middle, teams, class_start="09:10", class_end="17:50", lunch="12:50-14:00", breaks=(), mode="regular", morning_end="09:20", final_start="17:40"):
    paths = [Path(workbook).resolve(), Path(middle).resolve() if middle else None, Path(teams).resolve()]
    for path in paths:
        if path is None:
            continue
        if not path.is_file():
            raise ValueError(f"파일을 찾을 수 없습니다: {path.name}")
    hashes = [hashlib.sha256(p.read_bytes()).hexdigest() if p else None for p in paths]
    data = analyze(*paths, class_start, class_end, lunch, breaks, mode, morning_end, final_start)
    book = Workbook(paths[0])
    cells = book.cells("출결정보")
    highlighted = book.highlighted_cells("출결정보")
    changes = []
    for change in data.pop("source_updates"):
        before = book.value(cells.get(change["cell"]))
        if isinstance(before, dict):
            raise ValueError(f"{change['cell']}에는 수식이 있어 자동 변경할 수 없습니다.")
        after = change["after"]
        format_only = equivalent(before, after)
        if not format_only or change["cell"] not in highlighted:
            changes.append({**change, "before": before, "after": after, "display_after": after, "format_only": format_only, "source": change["source"] + (" · 강조색 보완" if format_only else "")})
    data["proposed_changes"] = changes
    data["schema"] = "attendance-changes-v2"
    data["created_at"] = datetime.now().isoformat(timespec="seconds")
    data["source_workbook"] = str(paths[0])
    data["roster"] = [{"row": row, "name": name} for row, name in book.roster().items()]
    if hashes != [hashlib.sha256(p.read_bytes()).hexdigest() if p else None for p in paths]:
        raise ValueError("분석 중 입력 파일이 변경되었습니다. 다시 실행하세요.")
    return data


def validate_journal(data):
    if data.get("schema") == "attendance-changes-v1":
        raise ValueError("1.0 버전의 변경내역입니다. 기존 웹앱과 동일한 표기로 저장하려면 입력 파일을 다시 분석하세요.")
    if data.get("schema") != "attendance-changes-v2":
        raise ValueError("이 프로그램에서 저장한 변경내역 JSON을 선택하세요.")
    datetime.strptime(data["date"], "%Y-%m-%d")
    seen = set()
    for c in data["proposed_changes"]:
        if not re.fullmatch(r"[HIJK][1-9]\d*", c["cell"]) or c["cell"] in seen:
            raise ValueError("변경내역에 허용되지 않거나 중복된 셀 주소가 있습니다.")
        seen.add(c["cell"])
        if c["cell"].startswith("K"):
            if c["after"] not in ("O", "x") or c["type"] != "text":
                raise ValueError("중간출결 변경값이 올바르지 않습니다.")
        else:
            expected = "duration" if c["cell"].startswith("J") else "time"
            if c["type"] != expected or not isinstance(c["after"], str):
                raise ValueError("Teams 변경값이 올바르지 않습니다.")
            if expected == "time":
                if parse_timestamp(c["after"]).date().isoformat() != data["date"]:
                    raise ValueError("Teams 변경값의 날짜가 출석부 날짜와 다릅니다.")
            elif not 0 <= parse_duration(c["after"]) <= 86400:
                raise ValueError("Teams 참석 시간이 하루 범위를 벗어났습니다.")
    return data


def load_journal(path):
    return validate_journal(json.loads(Path(path).read_text(encoding="utf-8")))


def prepare_patch(data, target, highlight):
    validate_journal(data)
    book = Workbook(target)
    values = book.values("출결정보")
    if data["date"].replace("-", ".") not in str(values.get("B4", "")):
        raise ValueError("반영 대상 출석부 날짜가 변경내역 날짜와 다릅니다.")
    roster = book.roster()
    expected_roster = {s["row"]: s["name"] for s in data["roster"]}
    if roster != expected_roster:
        raise ValueError("반영 대상 명단 또는 행 위치가 변경되었습니다. 새 출석부로 다시 분석하세요.")
    conflicts, patch = [], {}
    highlighted = book.highlighted_cells("출결정보") if highlight else set()
    for c in data["proposed_changes"]:
        row = int(c["cell"][1:])
        if roster.get(row) != c["name"]:
            raise ValueError("변경내역의 이름과 셀 위치가 일치하지 않습니다.")
        current = values.get(c["cell"])
        if equivalent(current, c["after"]) and (not highlight or c["cell"] in highlighted):
            continue
        if not equivalent(current, c["after"]) and not equivalent(current, c["before"]):
            conflicts.append(f"{c['name']} {c['cell']}")
        # The web app copies CSV strings verbatim and retains the original number format.
        patch[c["cell"]] = c["after"], ""
    if conflicts:
        raise ValueError("분석 후 수정된 셀이 있어 전체 반영을 중단했습니다: " + ", ".join(conflicts[:12]))
    if patch:
        book.patch("출결정보", patch, highlight=highlight)
        book.refresh_attendance_counts()
    return book, len(patch)


def export_result(data, destination):
    book, count = prepare_patch(data, data["source_workbook"], highlight=True)
    book.prepare_web_export()
    # Compute bytes before creating output files, so validation errors leave no partial workbook.
    content = book.to_bytes()
    folder = unique_folder(destination, data["date"] + "_출결결과")
    output = folder / (f"updated_{datetime.now():%Y%m%d_%H%M%S}_" + Path(data["source_workbook"]).name)
    write_new(output, content)
    write_new(folder / "변경내역.json", json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8"))
    write_new(folder / "출결검토.html", report_html(data).encode("utf-8"))
    return {"folder": str(folder), "workbook": str(output), "count": count}


def apply_original(data, target, destination):
    target = Path(target).resolve()
    book, count = prepare_patch(data, target, highlight=True)
    if not count:
        return {"count": 0, "backup": None, "target": str(target), "folder": None, "workbook": None}
    content = book.to_bytes()
    destination = Path(destination).resolve()
    backup = destination / "출석부_백업" / f"{target.stem}_{datetime.now():%Y%m%d_%H%M%S_%f}.xlsx"
    write_new(backup, book.original)
    folder = unique_folder(destination, data["date"] + "_원본반영")
    saved_copy = folder / target.name
    # Write review artifacts first: disk errors must not silently lose the user's original.
    write_new(folder / "변경내역.json", json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8"))
    write_new(folder / "출결검토.html", report_html(data).encode("utf-8"))
    write_new(saved_copy, content)
    atomic_replace(target, content, book.original)
    return {"count": count, "backup": str(backup), "target": str(target), "folder": str(folder), "workbook": str(saved_copy)}


def display_value(value, kind="text"):
    if value is None:
        return "(빈 셀)"
    if kind in ("time", "duration") and isinstance(value, (int, float)):
        seconds = round(value * 86400)
        return f"{seconds // 3600:02d}:{seconds % 3600 // 60:02d}:{seconds % 60:02d}"
    return str(value)


def report_html(data):
    e = lambda value: escape(str(value))
    def table(headers, rows):
        return "<table><thead><tr>" + "".join(f"<th>{e(h)}</th>" for h in headers) + "</tr></thead><tbody>" + "".join("<tr>" + "".join(f"<td>{e(v)}</td>" for v in row) + "</tr>" for row in rows) + "</tbody></table>"
    schedule, matching = data["schedule"], data["matching"]
    mode = schedule.get("mode", "regular")
    result = f'''<!doctype html><html lang="ko"><meta charset="utf-8"><title>{e(data['date'])} 출결 검토</title>
<style>body{{font:15px 'Malgun Gothic',sans-serif;color:#223047;background:#f4f6fa;margin:0}}main{{max-width:1100px;margin:30px auto;background:white;padding:36px;border-radius:14px}}h1{{font-size:26px}}h2{{margin-top:32px;font-size:19px}}p{{line-height:1.8}}table{{width:100%;border-collapse:collapse}}th,td{{padding:10px;text-align:left;border-bottom:1px solid #e0e6ee}}th{{background:#edf2f8}}small{{color:#65748a}}.note{{padding:14px;background:#fff2db;border-radius:8px}}@media print{{main{{margin:0;padding:0}}body{{background:white}}}}</style><main>
<h1>{e(data['date'])} 출결 검토</h1><p>운영 방식 {e(schedule.get('mode_label', '일반 수업'))} · 수강생 {data['workbook']['roster_count']}명 · 중간출결 연결 {matching['middle_matched']}명 · Teams 연결 {matching['teams_matched']}명 · 변경 후보 {len(data['proposed_changes'])}셀</p>'''
    if not schedule.get("middle_enabled", True):
        result += '<p class="note">중간출결: 해당 없음 (CSV 미사용). K열 기존 값은 유지하며 누락 경고를 표시하지 않습니다.</p>'
    if mode == "project":
        summary = data.get("project_summary", {})
        result += f'''<p>수업 {e(schedule['class_start'])}~{e(schedule['class_end'])}<br>오전 확인 {e(schedule['class_start'])}~{e(schedule['morning_end'])}<br>종료 확인 {e(schedule['final_start'])}~{e(schedule['class_end'])}</p>
<p class="note">프로젝트 팀 참석 보고서가 없으므로 낮 동안의 공백과 외출은 판정하지 않습니다. H·I·J에는 전체 모임방의 첫 입장·마지막 퇴장·누적 참여시간 원자료를 기록합니다.</p>
<h2>프로젝트 출석 확인</h2><p>오전 확인 {summary.get('morning_confirmed', 0)}명 · 오전 미확인 {summary.get('morning_unconfirmed', 0)}명 · 종료 확인 {summary.get('final_confirmed', 0)}명 · 종료 미확인 {summary.get('final_unconfirmed', 0)}명</p>'''
        result += table(["이름", "오전", "종료", "중간", "첫 입장", "마지막 퇴장"], [[c["name"], c["morning"], c["final"], c["middle"], c["first"], c["last"]] for c in data.get("project_attendance_checks", [])])
    else:
        result += f'''<p>수업 {e(schedule['class_start'])}~{e(schedule['class_end'])}<br>점심 {e(schedule['lunch'] or '제외 없음')}<br>추가 휴게: {e(', '.join(schedule['additional_breaks']) or '미입력')}</p>
<p class="note">외출 후보는 입력된 휴게시간만 제외한 10분 이상 연속 접속 공백입니다. 실제 휴게시간을 확인하세요. 최초 입장 전과 최종 퇴장 후는 외출에 포함하지 않으며, 기존 지각·조퇴·결석 판정을 자동으로 바꾸지 않습니다.</p>
<h2>외출 확인 후보</h2>'''
        result += table(["이름", "시작", "종료", "연속 공백"], [[g["name"], g["start"][11:], g["end"][11:], g["duration"]] for g in data["class_gap_candidates"]])
        if not data["class_gap_candidates"]:
            result += "<p>입력 기준에서 10분 이상 구간을 찾지 못했습니다.</p>"
    result += "<h2>확인할 기록</h2>" + table(["이름", "구분", "내용"], [[f["name"], f["kind"], f["detail"]] for f in data["findings"]])
    result += f"<p>중간출결 명단 밖: {e(', '.join(matching['middle_outside_roster']) or '없음')}<br>Teams 명단 밖: {e(', '.join(matching['teams_outside_roster']) or '없음')}</p>"
    if mode == "regular":
        result += "<h2>전체 접속 공백 (휴게시간 제외 전)</h2>" + table(["이름", "퇴장", "재입장", "공백"], [[g["name"], g["start"][11:], g["end"][11:], g["duration"]] for g in data["disconnection_candidates"]])
    result += "<h2>변경 셀</h2>" + table(["이름", "셀", "기존", "변경", "근거"], [[c["name"], c["cell"], display_value(c["before"], c["type"]), display_value(c["after"], c["type"]), c["source"]] for c in data["proposed_changes"]])
    result += "<h2>입력 파일</h2><p>" + "<br>".join(e(Path(s["path"]).name) for s in data["sources"] if s["path"]) + "</p><small>저장한 분석 시점의 검토 결과입니다. 입력 파일이나 운영 방식·확인 시간을 바꾸면 다시 분석하세요.</small></main></html>"
    return result
