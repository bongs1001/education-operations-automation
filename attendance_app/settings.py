"""Remember the common output folder and user-selected template paths."""

import json
import os
from pathlib import Path
import tempfile

FIELDS = ("output", "attendance_template", "ppt_template", "schedule")


def default_settings_path():
    return Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData/Local")) / "AttendanceUtility/settings.json"


def load_settings(path=None):
    try:
        data = json.loads(Path(path or default_settings_path()).read_text(encoding="utf-8"))
        return {key: value for key, value in data.items() if key in FIELDS and isinstance(value, str) and value.strip()}
    except (OSError, ValueError, AttributeError):
        return {}


def save_settings(data, path=None):
    path = Path(path or default_settings_path())
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary = tempfile.mkstemp(prefix=".settings_", suffix=".json", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as stream:
            json.dump({key: data[key] for key in FIELDS if data.get(key)}, stream, ensure_ascii=False, indent=2)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)
