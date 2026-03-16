from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any


class SettingsStore:
    def __init__(self, path: Path | None = None):
        if path is not None:
            self.path = path
            return

        config_dir = os.environ.get("FIBIONIC_SCALE_CONFIG_DIR")
        if config_dir:
            self.path = Path(config_dir) / "settings.json"
            return

        self.path = Path.home() / ".fibionic-scale" / "settings.json"

    def load(self) -> dict[str, Any]:
        if not self.path.exists():
            return {}

        try:
            return json.loads(self.path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return {}

    def save(self, data: dict[str, Any]) -> None:
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            self.path.write_text(json.dumps(data, indent=2), encoding="utf-8")
        except OSError:
            fallback = Path.cwd() / ".fibionic-scale" / "settings.json"
            fallback.parent.mkdir(parents=True, exist_ok=True)
            fallback.write_text(json.dumps(data, indent=2), encoding="utf-8")
            self.path = fallback
