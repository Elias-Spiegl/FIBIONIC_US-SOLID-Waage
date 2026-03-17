from __future__ import annotations

import unittest
from pathlib import Path
from unittest.mock import patch

import fibionic_scale_app.settings_store as settings_store_module
from fibionic_scale_app.settings_store import SettingsStore


class SettingsStoreTests(unittest.TestCase):
    def test_default_settings_path_uses_project_local_directory(self) -> None:
        store = SettingsStore()
        expected = Path(__file__).resolve().parents[1] / ".fibionic-scale" / "settings.json"
        self.assertEqual(store.path.resolve(), expected.resolve())

    def test_frozen_build_uses_executable_directory(self) -> None:
        fake_executable = Path("/tmp/fibionic/fibionic-gewichtslogging.exe")

        with patch.object(settings_store_module.sys, "frozen", True, create=True):
            with patch.object(settings_store_module.sys, "executable", str(fake_executable)):
                store = SettingsStore()

        expected = fake_executable.parent / ".fibionic-scale" / "settings.json"
        self.assertEqual(store.path.resolve(), expected.resolve())


if __name__ == "__main__":
    unittest.main()
