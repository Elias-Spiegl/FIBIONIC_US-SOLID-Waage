from __future__ import annotations

import os
import sys
from pathlib import Path


def runtime_support_issue(
    platform_name: str | None = None,
    version_info: tuple[int, int, int] | None = None,
) -> str | None:
    platform_name = platform_name or sys.platform
    if version_info is None:
        current = sys.version_info
        version_info = (current.major, current.minor, current.micro)

    if platform_name == "darwin" and version_info >= (3, 14, 0):
        return (
            "Diese App startet auf macOS aktuell nicht zuverlässig mit Python 3.14, "
            "weil PySide6/Qt beim Erzeugen des Fensters abstuerzt. "
            "Bitte ein Python-3.13-venv verwenden und die Abhaengigkeiten dort neu installieren."
        )

    return None


def configure_qt_runtime(
    platform_name: str | None = None,
    library_info=None,
) -> None:
    if os.environ.get("FIBIONIC_SKIP_QT_ENV_SETUP") == "1":
        return

    platform_name = platform_name or sys.platform
    if platform_name != "darwin":
        return

    if library_info is None:
        try:
            from PySide6.QtCore import QLibraryInfo
        except Exception:
            return
        library_info = QLibraryInfo

    plugins_path = str(library_info.path(library_info.LibraryPath.PluginsPath) or "").strip()
    if plugins_path:
        os.environ["QT_PLUGIN_PATH"] = plugins_path
        os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = str(Path(plugins_path) / "platforms")

    os.environ.setdefault("QT_QPA_PLATFORM", "cocoa")


def ensure_runtime_supported() -> None:
    if os.environ.get("FIBIONIC_ALLOW_UNSUPPORTED_RUNTIME") == "1":
        return

    issue = runtime_support_issue()
    if issue:
        raise SystemExit(issue)

    configure_qt_runtime()
