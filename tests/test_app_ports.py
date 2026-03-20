from __future__ import annotations

import os
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

from fibionic_scale_app.app import ScaleLoggerWindow
from fibionic_scale_app.serial_io import SerialPortDescriptor


class ManualPortSelectionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self) -> None:
        self.patches = [
            patch("fibionic_scale_app.app.SettingsStore.load", return_value={}),
            patch("fibionic_scale_app.app.SettingsStore.save"),
            patch("fibionic_scale_app.app.list_serial_port_descriptors", return_value=self._ports()),
            patch(
                "fibionic_scale_app.app.auto_detectable_serial_ports",
                return_value=["/dev/cu.auto", "/dev/cu.manual"],
            ),
            patch("fibionic_scale_app.app.verified_serial_port", return_value="/dev/cu.auto"),
            patch("fibionic_scale_app.app.preferred_serial_port", return_value="/dev/cu.auto"),
        ]
        for current_patch in self.patches:
            current_patch.start()

        self.window = ScaleLoggerWindow()
        self.window.poll_timer.stop()

    def tearDown(self) -> None:
        self.window.close()
        for current_patch in reversed(self.patches):
            current_patch.stop()

    @staticmethod
    def _ports() -> list[SerialPortDescriptor]:
        return [
            SerialPortDescriptor(device="/dev/cu.auto", description="Auto port"),
            SerialPortDescriptor(device="/dev/cu.manual", description="Manual port"),
        ]

    def test_manual_mode_keeps_dropdown_selection_across_refresh(self) -> None:
        self.window.toggle_manual_port_selection()
        self.window.manual_port_combo.setCurrentText("/dev/cu.manual")

        self.assertEqual(self.window._selected_port(), "/dev/cu.manual")
        self.assertEqual(self.window._saved_manual_port, "/dev/cu.manual")

        self.window.refresh_ports()

        self.assertTrue(self.window.manual_port_override)
        self.assertEqual(self.window._manual_port_value(), "/dev/cu.manual")
        self.assertEqual(self.window._selected_port(), "/dev/cu.manual")
        self.assertEqual(self.window.detected_port_label.text(), "/dev/cu.manual (manuell)")

    def test_manual_dropdown_selection_updates_port_immediately(self) -> None:
        self.window.toggle_manual_port_selection()

        self.window.manual_port_combo.setCurrentText("/dev/cu.manual")

        self.assertEqual(self.window._saved_manual_port, "/dev/cu.manual")
        self.assertEqual(self.window.active_source_value.text(), "/dev/cu.manual")
        self.assertEqual(self.window.connection_note_label.text(), "Manuelle Portwahl aktiv. Verwende /dev/cu.manual.")

    def test_auto_mode_can_be_reenabled_after_manual_override(self) -> None:
        with patch("fibionic_scale_app.app.QTimer.singleShot", side_effect=lambda _delay, callback: callback()):
            self.window.toggle_manual_port_selection()
            self.window.manual_port_combo.setCurrentText("/dev/cu.manual")

            self.window.use_auto_port_selection()

        self.assertFalse(self.window.manual_port_override)
        self.assertEqual(self.window._selected_port(), "/dev/cu.auto")
        self.assertEqual(self.window.detected_port_label.text(), "/dev/cu.auto (verifiziert)")


if __name__ == "__main__":
    unittest.main()
