from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
import os

from openpyxl import load_workbook

from fibionic_scale_app.excel_writer import (
    EXCEL_MODE_AUTO,
    EXCEL_MODE_FILE,
    LIVE_BACKEND,
    ExcelSession,
    LiveExcelUnavailableError,
)
from fibionic_scale_app.models import ExcelSettings


class ExcelSessionTests(unittest.TestCase):
    def test_writes_value_to_configured_cell_and_advances_row(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "messwerte.xlsx"
            session = ExcelSession(
                ExcelSettings(
                    path=str(path),
                    sheet_name="Produktion",
                    column="F",
                    start_row=3,
                    auto_advance=True,
                    mode=EXCEL_MODE_FILE,
                )
            )

            session.reset_row(3)
            result = session.write_value(12.34)

            workbook = load_workbook(path)
            worksheet = workbook["Produktion"]
            self.assertEqual(worksheet["F3"].value, 12.34)
            self.assertEqual(result.cell, "F3")
            self.assertEqual(session.current_row, 4)
            workbook.close()

    def test_detects_first_empty_row(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "messwerte.xlsx"
            session = ExcelSession(
                ExcelSettings(
                    path=str(path),
                    sheet_name="Produktion",
                    column="B",
                    start_row=2,
                    auto_advance=True,
                    mode=EXCEL_MODE_FILE,
                )
            )
            session.reset_row(2)
            session.write_value(1.23)
            session.write_value(4.56)

            detected = ExcelSession(session.settings).detect_current_row()
            self.assertEqual(detected, 4)

    def test_auto_mode_falls_back_to_file_writer(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            path = Path(tmpdir) / "messwerte.xlsx"
            session = ExcelSession(
                ExcelSettings(
                    path=str(path),
                    sheet_name="Produktion",
                    column="C",
                    start_row=2,
                    auto_advance=True,
                    mode=EXCEL_MODE_AUTO,
                )
            )
            session.reset_row(2)

            with patch.object(
                LIVE_BACKEND,
                "write_value",
                side_effect=LiveExcelUnavailableError("Excel nicht erreichbar"),
            ):
                result = session.write_value(7.89)

            workbook = load_workbook(path)
            worksheet = workbook["Produktion"]
            self.assertEqual(worksheet["C2"].value, 7.89)
            self.assertEqual(result.backend, EXCEL_MODE_FILE)
            self.assertEqual(session.current_row, 3)
            workbook.close()

    def test_live_backend_path_match_is_exact(self) -> None:
        class DummyBook:
            def __init__(self, fullname: str):
                self.fullname = fullname

        from fibionic_scale_app.excel_writer import LiveExcelBackend

        with tempfile.TemporaryDirectory() as tmpdir:
            target = Path(tmpdir) / "messwerte.xlsx"
            target.touch()
            book = DummyBook(str(target))

            self.assertTrue(LiveExcelBackend._book_matches_path(book, target.resolve()))

    def test_macos_onedrive_root_is_derived_from_selected_file(self) -> None:
        from fibionic_scale_app.excel_writer import LiveExcelBackend

        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir) / "Library" / "CloudStorage" / "OneDrive-fibionicGmbH"
            file_path = root / "Ordner" / "messwerte.xlsx"
            file_path.parent.mkdir(parents=True, exist_ok=True)
            file_path.touch()

            with patch("fibionic_scale_app.excel_writer.sys.platform", "darwin"):
                with patch.dict(os.environ, {}, clear=True):
                    LiveExcelBackend._configure_macos_onedrive_env(file_path)
                    configured = os.environ.get("ONEDRIVE_COMMERCIAL_MAC")
                    self.assertIsNotNone(configured)
                    self.assertEqual(Path(configured).resolve(), root.resolve())
                    self.assertEqual(os.environ.get("OneDriveCommercial"), configured)

    def test_macos_onedrive_prefers_home_alias_over_cloudstorage_root(self) -> None:
        from fibionic_scale_app.excel_writer import LiveExcelBackend

        with tempfile.TemporaryDirectory() as tmpdir:
            fake_home = Path(tmpdir) / "home"
            fake_home.mkdir()
            cloud_root = fake_home / "Library" / "CloudStorage" / "OneDrive-fibionicGmbH"
            cloud_root.mkdir(parents=True)
            alias = fake_home / "OneDrive - FIBIONIC"
            alias.symlink_to(cloud_root)

            with patch("fibionic_scale_app.excel_writer.Path.home", return_value=fake_home):
                preferred = LiveExcelBackend._preferred_macos_onedrive_root(cloud_root)

            self.assertEqual(preferred, str(alias))


if __name__ == "__main__":
    unittest.main()
