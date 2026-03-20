from __future__ import annotations

import os
import sys


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


def ensure_runtime_supported() -> None:
    if os.environ.get("FIBIONIC_ALLOW_UNSUPPORTED_RUNTIME") == "1":
        return

    issue = runtime_support_issue()
    if issue:
        raise SystemExit(issue)
