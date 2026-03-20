from __future__ import annotations

import unittest

from fibionic_scale_app.runtime_support import runtime_support_issue


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


if __name__ == "__main__":
    unittest.main()
