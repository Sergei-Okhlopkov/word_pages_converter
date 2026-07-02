import json
from pathlib import Path


def get_settings_path() -> Path:
    settings_dir = Path.home() / "AppData" / "Local" / "WordPagesToImages"
    settings_dir.mkdir(parents=True, exist_ok=True)
    return settings_dir / "settings.json"


def load_settings() -> dict:
    settings_path = get_settings_path()
    if not settings_path.exists():
        return {"theme": "light"}
    try:
        return json.loads(settings_path.read_text(encoding="utf-8"))
    except Exception:
        return {"theme": "light"}


def save_settings(settings: dict) -> None:
    settings_path = get_settings_path()
    settings_path.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")
