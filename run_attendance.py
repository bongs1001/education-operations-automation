import json
from pathlib import Path
import sys

from attendance_app.gui import Application, asset_root, main

if __name__ == "__main__":
    if len(sys.argv) == 3 and sys.argv[1] == "--self-test":
        output = Path(sys.argv[2])
        try:
            app = Application(settings_file=output.with_suffix(".settings.json"))
            app.withdraw()
            app.update_idletasks()
            assets = asset_root()
            required = [assets / name for name in ("출석부_양식.xlsx", "산출물_양식.pptx", "교육일정.csv")]
            if not all(p.is_file() for p in required):
                raise RuntimeError("Bundled templates missing")
            tabs = [app.notebook.tab(tab, "text") for tab in app.notebook.tabs()]
            if tabs != ["출석부 업데이트", "출석부 생성", "산출물 생성", "사용 안내"]:
                raise RuntimeError("Unexpected application tabs")
            if app.value("update_mode") != "일반 수업" or "project" not in app.tables:
                raise RuntimeError("Attendance modes missing")
            skip = next((b for b in app.action_buttons if b.cget('text') == '중간출결 안 함'), None)
            if skip is None:
                raise RuntimeError('Middle attendance skip button missing')
            app.var('middle').set('previous.csv')
            skip.invoke()
            if app.value('middle'):
                raise RuntimeError('Middle attendance skip failed')
            app.var("update_mode").set("프로젝트")
            app.update_idletasks()
            if not app.regular_time_entries[0].instate(["disabled"]) or not app.project_time_entries[0].instate(["!disabled"]):
                raise RuntimeError("Project mode controls are not configured")
            result = {
                "ok": True,
                "title": app.title(),
                "tabs": tabs,
                "attendance_modes": ["일반 수업", "프로젝트"],
                "generation_modes": sorted(app.event_tables),
                "assets": [p.name for p in required],
            }
            app.destroy()
        except Exception as error:
            result = {"ok": False, "error": str(error)}
        output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        sys.exit(0 if result["ok"] else 1)
    else:
        main()
