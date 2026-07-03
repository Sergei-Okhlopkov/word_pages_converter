import json
from pathlib import Path


def get_settings_path() -> Path:
    settings_dir = Path.home() / "AppData" / "Local" / "WordWorkTool"
    settings_dir.mkdir(parents=True, exist_ok=True)
    return settings_dir / "settings.json"


def load_settings() -> dict:
    settings_path = get_settings_path()
    if not settings_path.exists():
        return {"theme": "light", "tools": {}}
    try:
        data = json.loads(settings_path.read_text(encoding="utf-8"))
    except Exception:
        return {"theme": "light", "tools": {}}
    data.setdefault("theme", "light")
    data.setdefault("tools", {})
    return data


def get_tool_settings(settings: dict, tool_id: str) -> dict:
    tool_settings = settings.setdefault("tools", {}).setdefault(tool_id, {})
    tool_settings.setdefault("output_dir", "")
    return tool_settings


def save_tool_output_dir(settings: dict, tool_id: str, output_dir: str) -> None:
    get_tool_settings(settings, tool_id)["output_dir"] = output_dir
    save_settings(settings)


def save_settings(settings: dict) -> None:
    settings_path = get_settings_path()
    settings_path.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")
