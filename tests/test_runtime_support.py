from __future__ import annotations

import os
import unittest
from unittest.mock import patch

from fibionic_scale_app.runtime_support import configure_qt_runtime, runtime_support_issue


class RuntimeSupportTests(unittest.TestCase):
    def test_allows_python_313_on_macos(self) -> None:
        self.assertIsNone(runtime_support_issue(platform_name="darwin", version_info=(3, 13, 7)))

    def test_blocks_python_314_on_macos(self) -> None:
        issue = runtime_support_issue(platform_name="darwin", version_info=(3, 14, 0))

        self.assertIsNotNone(issue)
        assert issue is not None
        self.assertIn("Python 3.14", issue)
        self.assertIn("Python-3.13-venv", issue)

    def test_allows_python_314_on_non_macos_platforms(self) -> None:
        self.assertIsNone(runtime_support_issue(platform_name="win32", version_info=(3, 14, 0)))

    def test_configures_qt_plugin_paths_on_macos(self) -> None:
        class DummyLibraryInfo:
            class LibraryPath:
                PluginsPath = "plugins"

            @staticmethod
            def path(_library_path):
                return "/tmp/pyside6/plugins"

        with patch.dict("os.environ", {}, clear=True):
            configure_qt_runtime(platform_name="darwin", library_info=DummyLibraryInfo)

            self.assertEqual(os.environ["QT_PLUGIN_PATH"], "/tmp/pyside6/plugins")
            self.assertEqual(os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"], "/tmp/pyside6/plugins/platforms")
            self.assertEqual(os.environ["QT_QPA_PLATFORM"], "cocoa")

    def test_keeps_explicit_qt_platform_override(self) -> None:
        class DummyLibraryInfo:
            class LibraryPath:
                PluginsPath = "plugins"

            @staticmethod
            def path(_library_path):
                return "/tmp/pyside6/plugins"

        with patch.dict("os.environ", {"QT_QPA_PLATFORM": "offscreen"}, clear=True):
            configure_qt_runtime(platform_name="darwin", library_info=DummyLibraryInfo)

            self.assertEqual(os.environ["QT_QPA_PLATFORM"], "offscreen")


if __name__ == "__main__":
    unittest.main()
