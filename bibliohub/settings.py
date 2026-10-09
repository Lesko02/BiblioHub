import json
import os
from pathlib import Path


def default_settings_path() -> Path:
    base = os.environ.get("APPDATA") or Path.home()
    return Path(base) / "BiblioHub" / "settings.json"


class Settings:
    """Preferenze persistenti (ultimi percorsi usati, ecc.), divise per sezione."""

    def __init__(self, path: Path | None = None):
        self.path = path or default_settings_path()
        try:
            self._data = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            self._data = {}

    def get(self, section: str, key: str, default=None):
        return self._data.get(section, {}).get(key, default)

    def set(self, section: str, key: str, value) -> None:
        self._data.setdefault(section, {})[key] = value

    def save(self) -> None:
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            self.path.write_text(json.dumps(self._data, indent=2, ensure_ascii=False),
                                 encoding="utf-8")
        except OSError:
            pass
