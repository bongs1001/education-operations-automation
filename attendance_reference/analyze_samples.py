"""Read-only reference analysis. Never saves or changes the source workbook."""

import argparse
from collections import Counter, defaultdict
from datetime import datetime
import hashlib
import json
from pathlib import Path
import re
import unicodedata

try:
    from .inspect_samples import read_csv, read_workbook
except ImportError:
    from inspect_samples import read_csv, read_workbook


def name_key(value):
    return unicodedata.normalize("NFC", value).strip()


def teams_section(rows, title):
    start = next((i for i, row in enumerate(rows) if row and row[0].strip() == title), None)
    if start is None or start + 1 >= len(rows):
        raise ValueError(f"Teams 보고서에 '{title}' 섹션이 없습니다. 참가자·활동 기록이 포함된 보고서를 선택하세요.")
    headers = rows[start + 1]
    records = []
    for row in rows[start + 2:]:
        if not row or re.match(r"^\d+\. ", row[0]):
            break
        if len(row) != len(headers):
            raise ValueError(f"{title}: 열 개수가 맞지 않는 행이 있습니다.")
        records.append(dict(zip(headers, row)))
    return records


def parse_timestamp(value):
    return datetime.strptime(value, "%m/%d/%y, %I:%M:%S %p")


def parse_duration(value):
    parts = re.fullmatch(r"\s*(?:(\d+)시간\s*)?(?:(\d+)분\s*)?(?:(\d+)초\s*)?", value)
    if not parts or not value.strip():
        raise ValueError(f"Invalid duration: {value}")
    h, m, s = (int(v or 0) for v in parts.groups())
    return h * 3600 + m * 60 + s


def disconnected_intervals(intervals, threshold_seconds=600):
    """Union overlapping sessions before finding gaps; threshold is inclusive."""
    merged = []
    for start, end in sorted(intervals):
        if end < start:
            raise ValueError("Session ends before it starts")
        if merged and start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    return [(left[1], right[0]) for left, right in zip(merged, merged[1:])
            if (right[0] - left[1]).total_seconds() >= threshold_seconds]


def subtract_breaks(start, end, breaks):
    parts = [(start, end)] if start < end else []
    for break_start, break_end in sorted(breaks):
        if break_start >= break_end:
            raise ValueError("Break start must precede its end")
        remaining = []
        for left, right in parts:
            if break_end <= left or break_start >= right:
                remaining.append((left, right))
            else:
                if left < break_start:
                    remaining.append((left, break_start))
                if break_end < right:
                    remaining.append((break_end, right))
        parts = remaining
    return parts


def class_gap_candidates(intervals, lesson_start, lesson_end, breaks, threshold_seconds=600):
    if lesson_start >= lesson_end:
        raise ValueError("Lesson start must precede its end")
    candidates = []
    for start, end in disconnected_intervals(intervals, threshold_seconds=0):
        for left, right in subtract_breaks(max(start, lesson_start), min(end, lesson_end), breaks):
            if (right - left).total_seconds() >= threshold_seconds:
                candidates.append((left, right))
    return candidates


def overlaps_window(intervals, window_start, window_end):
    """Return whether any Teams session touches an attendance-check window."""
    if window_start >= window_end:
        raise ValueError("Attendance window start must precede its end")
    return any(start <= window_end and end >= window_start for start, end in intervals)


def unique_index(records, field):
    grouped = defaultdict(list)
    for row in records:
        grouped[name_key(row[field])].append(row)
    duplicates = [key for key, rows in grouped.items() if len(rows) > 1]
    if duplicates:
        raise ValueError(f"이름이 중복되어 자동 연결할 수 없습니다: {', '.join(duplicates)}. 원본에서 구분 가능한 이름으로 정리하세요.")
    return {name: rows[0] for name, rows in grouped.items()}


def clock_text(seconds):
    return f"{seconds // 3600:02d}:{seconds % 3600 // 60:02d}:{seconds % 60:02d}"


def analyze(workbook_path, middle_path, teams_path, class_start="09:10", class_end="17:50", lunch="12:50-14:00", extra_breaks=(), mode="regular", morning_end="09:20", final_start="17:40"):
    if mode not in ("regular", "project"):
        raise ValueError("운영 방식은 일반 수업 또는 프로젝트여야 합니다.")
    workbook = read_workbook(workbook_path)
    sheet = next((s for s in workbook["sheets"] if s["sheet"] == "출결정보"), None)
    if sheet is None:
        raise ValueError("출석부에 '출결정보' 시트가 없습니다.")
    cells = sheet["cells"]
    roster = [{"row": int(ref[1:]), "name": value, "number": int(cells["B" + ref[1:]])}
              for ref, value in cells.items()
              if re.fullmatch(r"D\d+", ref) and cells.get("B" + ref[1:], "").isdigit() and int(cells["B" + ref[1:]]) > 0]
    if not roster:
        raise ValueError("출결정보 시트 B열 순번·D열 성명에서 수강생을 찾지 못했습니다.")
    unique_index(roster, "name")
    middle_encoding = middle_delimiter = None
    middle_records = []
    middle = {}
    if middle_path:
        middle_encoding, middle_delimiter, middle_rows = read_csv(middle_path)
        if not middle_rows or not {"날짜", "이름", "출석여부"}.issubset(middle_rows[0]):
            raise ValueError("중간출결 CSV에 날짜·이름·출석여부 열이 필요합니다.")
        if any(len(row) != len(middle_rows[0]) for row in middle_rows[1:] if row):
            raise ValueError("중간출결 CSV의 열 개수가 맞지 않습니다.")
        middle_records = [dict(zip(middle_rows[0], row)) for row in middle_rows[1:] if row]
        middle = unique_index(middle_records, "이름")
    teams_encoding, teams_delimiter, teams_rows = read_csv(teams_path)
    participants = teams_section(teams_rows, "2. 참가자")
    participant_index = unique_index(participants, "이름")
    activities = teams_section(teams_rows, "3. 모임 내 활동")
    date_match = re.search(r"\d{4}\.\d{2}\.\d{2}", cells.get("B4", ""))
    if not date_match:
        raise ValueError("출결정보 B4 날짜는 '일자 : YYYY.MM.DD' 형식이어야 합니다.")
    date = datetime.strptime(date_match[0], "%Y.%m.%d").date()
    def on_date(clock):
        return datetime.combine(date, datetime.strptime(clock, "%H:%M").time())
    lesson_start, lesson_end = on_date(class_start), on_date(class_end)
    if lesson_start >= lesson_end:
        raise ValueError("수업 종료는 시작보다 늦어야 합니다.")
    morning_check_end = final_check_start = None
    if mode == "project":
        morning_check_end, final_check_start = on_date(morning_end), on_date(final_start)
        if not (lesson_start < morning_check_end <= final_check_start < lesson_end):
            raise ValueError("프로젝트 확인 시간은 수업 시작 < 오전 확인 종료 ≤ 종료 확인 시작 < 수업 종료 순서여야 합니다.")
    break_ranges = []
    if mode == "regular":
        for value in [lunch, *extra_breaks]:
            if value:
                left, right = value.split("-")
                interval = on_date(left), on_date(right)
                if interval[0] >= interval[1]:
                    raise ValueError("휴게 종료는 시작보다 늦어야 합니다.")
                break_ranges.append(interval)
    if middle_path and {row["날짜"] for row in middle_records} != {date.isoformat()}:
        raise ValueError("중간출결 CSV 날짜와 출석부 날짜가 다릅니다.")
    sessions = defaultdict(list)
    for row in activities:
        start, end = parse_timestamp(row["참가 시간"]), parse_timestamp(row["나간 시간"])
        if start.date() != date or end.date() != date:
            raise ValueError("Teams 활동 날짜와 출석부 날짜가 다릅니다.")
        if end < start:
            raise ValueError("Teams 퇴장 시각이 입장 시각보다 빠른 기록이 있습니다.")
        name = name_key(row["이름"])
        if name not in participant_index or row["이메일"].casefold() != participant_index[name]["이메일"].casefold():
            raise ValueError(f"Teams identity mismatch: {name}")
        sessions[name].append((start, end))
    changes, updates, gaps, candidates, findings, project_checks = [], [], [], [], [], []
    for student in roster:
        row_number, name = student["row"], name_key(student["name"])
        if name in participant_index:
            p = participant_index[name]
            first, last = parse_timestamp(p["처음 들어온 시간"]), parse_timestamp(p["마지막 나간 시간"])
            if first.date() != date or last.date() != date or last < first:
                raise ValueError("Teams participant date differs from workbook date")
            if not sessions[name]:
                raise ValueError(f"{name}: 참가자 요약은 있으나 활동 세션이 없습니다.")
            parse_duration(p["모임 참여 시간"])
            values = {"H": p["처음 들어온 시간"], "I": p["마지막 나간 시간"], "J": p["모임 참여 시간"]}
            for col, value in values.items():
                ref = f"{col}{row_number}"
                change = {"cell": ref, "name": name, "before": cells.get(ref), "after": value, "source": "Teams 2. 참가자", "type": "duration" if col == "J" else "time"}
                updates.append(change)
                if cells.get(ref) != value:
                    changes.append(change)
            if mode == "regular":
                for start, end in disconnected_intervals(sessions[name]):
                    seconds = int((end - start).total_seconds())
                    gaps.append({"name": name, "start": start.isoformat(), "end": end.isoformat(), "seconds": seconds, "duration": clock_text(seconds), "status": "휴게시간 미제외·외출 여부 미확정"})
                for start, end in class_gap_candidates(sessions[name], lesson_start, lesson_end, break_ranges):
                    seconds = int((end - start).total_seconds())
                    candidates.append({"name": name, "start": start.isoformat(), "end": end.isoformat(), "seconds": seconds, "duration": clock_text(seconds), "status": "입력된 휴게시간 제외·실제 휴게시간 및 외출 여부 확인 필요"})
        else:
            findings.append({"name": name, "kind": "Teams 미매칭", "detail": "Teams 셀의 기존 값을 유지하고 확인 대상으로 표시"})
        if name in middle:
            status = middle[name]["출석여부"]
            if status not in ("출석", "미출석"):
                raise ValueError(f"Unknown middle attendance status: {status}")
            ref = f"K{row_number}"
            value = "O" if status == "출석" else "x"
            change = {"cell": ref, "name": name, "before": cells.get(ref), "after": value, "source": "중간출결 CSV", "type": "text"}
            updates.append(change)
            if cells.get(ref) != value:
                changes.append(change)
            if status == "미출석" and cells.get(f"M{row_number}", "").lower() == "o":
                findings.append({"name": name, "kind": "자료 간 불일치", "detail": "중간출결 미출석 / 퇴실 호명 o; 기존 호명값은 유지"})
        elif middle_path:
            findings.append({"name": name, "kind": "중간출결 미매칭", "detail": "K열 기존 값을 유지하고 확인 대상으로 표시"})
        if cells.get(f"O{row_number}") == "결석" and name in participant_index:
            findings.append({"name": name, "kind": "수기 확정출결과 접속 기록 병존", "detail": "확정출결 결석 / Teams 접속 기록 있음; 수기 판정 유지"})
        if mode == "project":
            has_teams = name in participant_index
            morning_present = has_teams and overlaps_window(sessions[name], lesson_start, morning_check_end)
            final_present = has_teams and overlaps_window(sessions[name], final_check_start, lesson_end)
            if has_teams and not morning_present:
                findings.append({"name": name, "kind": "오전 출석 미확인", "detail": f"{class_start}~{morning_end} 전체 모임방 접속 기록 없음"})
            if has_teams and not final_present:
                findings.append({"name": name, "kind": "종료 출석 미확인", "detail": f"{final_start}~{class_end} 전체 모임방 접속 기록 없음"})
            project_checks.append({
                "name": name,
                "morning": "확인" if morning_present else "미확인",
                "final": "확인" if final_present else "미확인",
                "middle": "해당 없음" if not middle_path else "O" if name in middle and middle[name]["출석여부"] == "출석" else "x" if name in middle else "미확인",
                "first": participant_index[name]["처음 들어온 시간"] if has_teams else "",
                "last": participant_index[name]["마지막 나간 시간"] if has_teams else "",
            })
    names = {name_key(s["name"]) for s in roster}
    project_summary = {
        "morning_confirmed": sum(c["morning"] == "확인" for c in project_checks),
        "morning_unconfirmed": sum(c["morning"] == "미확인" for c in project_checks),
        "final_confirmed": sum(c["final"] == "확인" for c in project_checks),
        "final_unconfirmed": sum(c["final"] == "미확인" for c in project_checks),
    }
    return {
        "date": date.isoformat(),
        "sources": [{"path": str(p.resolve()) if p else "", "sha256": hashlib.sha256(p.read_bytes()).hexdigest() if p else None} for p in (workbook_path, middle_path, teams_path)],
        "source_formats": {"middle": {"encoding": middle_encoding, "delimiter": middle_delimiter}, "teams": {"encoding": teams_encoding, "delimiter": teams_delimiter}},
        "workbook": {"sheets": [s["sheet"] for s in workbook["sheets"]], "media_count": workbook["media_count"], "student_range": f"D{min(s['row'] for s in roster)}:D{max(s['row'] for s in roster)}", "update_range": f"H{min(s['row'] for s in roster)}:K{max(s['row'] for s in roster)}", "roster_count": len(roster)},
        "matching": {"middle_records": len(middle), "middle_matched": len(names & middle.keys()), "middle_outside_roster": sorted(middle.keys() - names), "middle_matched_status_counts": dict(Counter(middle[n]["출석여부"] for n in names & middle.keys())), "teams_participants": len(participant_index), "teams_matched": len(names & participant_index.keys()), "teams_missing": sorted(names - participant_index.keys()), "teams_outside_roster": sorted(participant_index.keys() - names), "activity_rows": len(activities)},
        "schedule": {"middle_enabled": bool(middle_path), "mode": mode, "mode_label": "프로젝트" if mode == "project" else "일반 수업", "class_start": class_start, "class_end": class_end, "lunch": lunch, "additional_breaks": list(extra_breaks), "morning_end": morning_end, "final_start": final_start, "threshold_seconds": 600, "method": "전체 모임방의 오전·종료 확인 구간에 접속 세션이 겹치는지 확인; 낮 동안 외출 미판정" if mode == "project" else "수업시간으로 제한한 뒤 입력된 휴게시간을 제외하고 남은 각 연속 구간에 10분 이상 기준 적용", "source": "사용자 제공 운영시간; 프로젝트 팀 보고서가 없으면 낮 동안 접속·외출 검증 불가" if mode == "project" else "사용자 제공 운영시간; 유동적인 점심·휴게시간은 날짜별 실제 시각 입력 필요"},
        "proposed_changes": changes, "source_updates": updates, "disconnection_candidates": gaps, "class_gap_candidates": candidates, "project_attendance_checks": project_checks, "project_summary": project_summary, "findings": findings,
        "limitations": (["프로젝트 팀 참석 보고서가 없어 낮 동안의 접속 공백·외출은 판정하지 않음", "H·I·J는 전체 모임방의 첫 입장·마지막 퇴장·누적 참여시간 원자료", "오전·종료·중간 확인 결과로 지각·조퇴·결석을 자동 확정하지 않음", "원본 파일 수정 없음"] if mode == "project" else ["추가 휴게시간은 입력된 것만 제외; 유동적인 실제 휴게시간은 별도 확인 필요", "최초 입장 전·최종 퇴장 후는 외출 공백에서 제외", "외출·지각·조퇴·결석 자동 확정 없음", "원본 파일 수정 없음"]),
    }


def write_report(data, destination):
    match = data["matching"]
    lines = ["# 출석 자동화 샘플 분석", "", f"분석 기준일: {data['date']}", "", "## 입력 파일과 연결 결과", "",
        "| 항목 | 확인 결과 |", "| --- | --- |",
        f"| 출석부 | 수강생 {data['workbook']['roster_count']}명, 출결정보 D7:D36 |",
        f"| 중간출결 | {match['middle_records']}명 중 명단과 {match['middle_matched']}명 일치 |",
        f"| 중간출결 상태 | 출석 {match['middle_matched_status_counts'].get('출석', 0)}명 / 미출석 {match['middle_matched_status_counts'].get('미출석', 0)}명 |",
        f"| Teams | 참가자 {match['teams_participants']}명 중 수강생 {match['teams_matched']}명 일치, 활동 기록 {match['activity_rows']}행 |",
        f"| 증빙 | 이미지 {data['workbook']['media_count']}개; 기존 시트·그림·서식·수식 보존 필요 |", "",
        f"- 중간출결 명단 밖 인원: {', '.join(match['middle_outside_roster'])}",
        f"- Teams 명단 밖 인원: {', '.join(match['teams_outside_roster'])}",
        f"- Teams 기록이 없는 수강생: {', '.join(match['teams_missing'])}",
        "- 출석부 순번, 중간출결 학생ID, Teams 계정 번호는 서로 다르므로 같은 ID로 취급하지 않는다.",
        "- 이번 샘플은 이름으로 정확히 연결 가능하다. 동명이인·중복 이름은 자동 연결하지 않고 매핑을 요구한다.", "",
        "## 출석부 반영 규칙 초안", "", "| 대상 | 원본 데이터 |", "| --- | --- |",
        "| H7:H36 | Teams 참가자 섹션의 처음 들어온 시간 |",
        "| I7:I36 | Teams 참가자 섹션의 마지막 나간 시간 |",
        "| J7:J36 | Teams 참가자 섹션의 모임 참여 시간; 마지막 시간−처음 시간으로 대체하지 않음 |",
        "| K7:K36 | 중간출결 출석 → o, 미출석 → x |", "",
        f"현재 파일에서는 H~J {sum(c['source'].startswith('Teams') for c in data['proposed_changes'])}개, K {sum(c['source'] == '중간출결 CSV' for c in data['proposed_changes'])}개, 총 {len(data['proposed_changes'])}개 셀이 업데이트 후보이다.",
        "LMS(E:G), 호명(L:M), 특이사항(N), 확정출결(O), 비고(P)는 자동으로 덮어쓰지 않는다. K37/K38 집계는 기존 수식에 따라 26/4가 되는지 업데이트 후 검증한다.",
        "별도 결과 파일은 변경 셀을 주황색으로 표시한다. 기존 출석부 반영 기능은 변경 내역의 이전 값과 현재 값을 비교하고, 충돌 시 덮어쓰지 않는다. 원본 백업·임시 저장 후 교체를 적용한다.", "",
        "## 10분 이상 접속이 끊긴 구간", "",
        "Teams ‘3. 모임 내 활동’에서 같은 사람의 세션을 시간순으로 정렬하고 겹치는 세션을 합친 뒤, 퇴장부터 다음 재입장까지 600초 이상인 구간을 계산했다. 최초 입장 전과 최종 퇴장 후는 이 표에 포함하지 않았다.",
        "점심·휴게시간을 아직 제외하지 않은 접속 공백이며, 외출 확정 기록이 아니다.", "",
        "| 이름 | 퇴장 | 재입장 | 접속 공백 |", "| --- | --- | --- | --- |"]
    for gap in data["disconnection_candidates"]:
        lines.append(f"| {gap['name']} | {gap['start'][11:]} | {gap['end'][11:]} | {gap['duration']} |")
    schedule = data["schedule"]
    lines += ["", "## 수업·점심시간을 반영한 확인 후보", "",
        f"사용자 제공 수업시간 {schedule['class_start']}~{schedule['class_end']}, 기본 점심시간 {schedule['lunch']}를 적용했다.",
        "점심은 강사에 따라 12:41~12:45에 시작할 수 있으며, 매시간 45~50분 무렵 시작하는 10분 휴게도 유동적이다. 고정 반복 휴게로 추정하여 제외하지 않고 날짜별 실제 시간을 입력받는다.",
        f"추가 입력 휴게시간: {', '.join(schedule['additional_breaks']) or '없음(실제 휴게가 없었다는 뜻은 아님)'}. 휴게시간을 제외하고 남은 각 연속 구간이 10분 이상인 경우만 아래에 표시한다.", "",
        "| 이름 | 시작 | 종료 | 연속 공백 |", "| --- | --- | --- | --- |"]
    for gap in data["class_gap_candidates"]:
        lines.append(f"| {gap['name']} | {gap['start'][11:]} | {gap['end'][11:]} | {gap['duration']} |")
    lines += ["", "이 결과는 외출 확정이 아니다. 실제 휴게시간을 추가하면 구간이 분리되거나 줄어들 수 있다. 최종 퇴장 후 재입장이 없는 기록은 외출 구간에 섞지 않고 기존 조퇴·결석 판단의 참고자료로 둔다."]
    lines += ["", "## 확인 대상", ""]
    lines.extend(f"- {f['name']}: {f['kind']} — {f['detail']}" for f in data["findings"])
    lines += ["", "## 개발에 필요한 남은 기준", "",
        "1. 날짜별 실제 점심·휴게시간. 기본값은 위 운영시간을 사용하며 실제 시각으로 수정 후 재계산한다. 휴게시간을 사이에 둔 짧은 공백들을 합산해 10분 외출로 판정하지 않는다.",
        "2. 산출물 종류와 파일 양식. 현재 첨부는 출석부·출결자료이며 산출물 템플릿은 아직 없다.",
        "3. 출석부 일괄 생성 시 공휴일·주말 수업과 보강일 처리. 실제 수업일 목록을 우선 적용할 수 있도록 설계한다.", "",
        "## 확인 범위", "", "원본 Excel·CSV는 읽기만 수행했다. 이 문서는 입력 구조와 계산 결과 분석이며, 업데이트된 Excel 파일이나 완성된 실행 프로그램은 아직 생성하지 않았다.", "",
        "입력 파일:", ""]
    lines.extend(f"- `{Path(source['path']).name}` (SHA-256: `{source['sha256']}`)" for source in data["sources"])
    destination.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", required=True, type=Path)
    parser.add_argument("--middle", type=Path)
    parser.add_argument("--teams", required=True, type=Path)
    parser.add_argument("--class-start", default="09:10")
    parser.add_argument("--class-end", default="17:50")
    parser.add_argument("--lunch", default="12:50-14:00", help="Actual lunch interval, HH:MM-HH:MM")
    parser.add_argument("--break", dest="extra_breaks", action="append", default=[], help="Additional actual break, HH:MM-HH:MM; repeat as needed")
    parser.add_argument("--output", type=Path, default=Path(__file__).parent / "results")
    args = parser.parse_args()
    data = analyze(args.workbook, args.middle, args.teams, args.class_start, args.class_end, args.lunch, args.extra_breaks)
    args.output.mkdir(parents=True, exist_ok=True)
    (args.output / "sample_analysis.json").write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    write_report(data, args.output / "샘플_분석_및_개발기준.md")
    print(json.dumps({"matching": data["matching"], "changes": len(data["proposed_changes"]), "raw_gap_count": len(data["disconnection_candidates"]), "schedule": data["schedule"], "class_gap_candidates": data["class_gap_candidates"], "findings": data["findings"]}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
