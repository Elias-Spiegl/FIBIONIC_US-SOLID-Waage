from __future__ import annotations

import queue
import sys
from pathlib import Path

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QCloseEvent
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from .excel_writer import (
    EXCEL_MODE_AUTO,
    ExcelSession,
    excel_mode_options,
    live_backend_status_text,
    mode_label,
)
from .models import CaptureSettings, ExcelSettings, SerialSettings
from .serial_io import ScaleStreamWorker, StreamEvent, list_serial_ports
from .settings_store import SettingsStore
from .stability import CaptureState, WeightCaptureEngine

COLORS = {
    "page": "#F4EFE6",
    "card": "#FFF9F1",
    "tile": "#F7E9D8",
    "surface": "#FFFDF9",
    "ink": "#14304A",
    "muted": "#64748B",
    "accent": "#E86A33",
    "accent_dark": "#C55121",
    "accent_soft": "#FFE3D6",
    "line": "#DCCBBB",
}


class ScaleLoggerWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("FIBIONIC | US-Solid Waage")
        self.resize(1280, 860)
        self.setMinimumSize(1120, 760)

        self.settings_store = SettingsStore()
        self.stream_worker: ScaleStreamWorker | None = None
        self.capture_engine = WeightCaptureEngine(CaptureSettings())
        self.excel_session: ExcelSession | None = None

        self._build_ui()
        self._apply_styles()
        self._load_settings()
        self.refresh_ports()
        self._refresh_excel_backend_hint()
        self._refresh_capture_window_label()

        self.poll_timer = QTimer(self)
        self.poll_timer.setInterval(120)
        self.poll_timer.timeout.connect(self._poll_worker)
        self.poll_timer.start()

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)

        layout = QHBoxLayout(central)
        layout.setContentsMargins(22, 18, 22, 18)
        layout.setSpacing(18)

        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(14)

        left_scroll = QScrollArea()
        left_scroll.setWidgetResizable(True)
        left_scroll.setFrameShape(QFrame.Shape.NoFrame)
        left_scroll.setWidget(left_panel)
        left_scroll.setMinimumWidth(430)
        left_scroll.setMaximumWidth(500)

        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(14)

        layout.addWidget(left_scroll, 0)
        layout.addWidget(right_panel, 1)

        hero = QFrame()
        hero_layout = QVBoxLayout(hero)
        hero_layout.setContentsMargins(20, 18, 20, 18)
        hero_layout.setSpacing(6)

        hero_title = QLabel("US-Solid Waage -> Excel Logger")
        hero_title.setObjectName("HeroTitle")
        hero_copy = QLabel(
            "RS232-Datenstrom stabilisieren, Zielgewicht pruefen und jeden bestaetigten Messwert sauber in Excel schreiben."
        )
        hero_copy.setWordWrap(True)
        hero_copy.setObjectName("HeroCopy")
        hero_layout.addWidget(hero_title)
        hero_layout.addWidget(hero_copy)

        right_layout.addWidget(hero)
        right_layout.addLayout(self._build_metrics_row())
        right_layout.addWidget(self._build_monitor_box(), 1)

        left_layout.addWidget(self._build_connection_box())
        left_layout.addWidget(self._build_capture_box())
        left_layout.addWidget(self._build_excel_box())
        left_layout.addStretch(1)

    def _build_connection_box(self) -> QGroupBox:
        box = QGroupBox("Verbindung")
        layout = QGridLayout(box)
        layout.setHorizontalSpacing(10)
        layout.setVerticalSpacing(8)

        layout.addWidget(self._field_label("Serieller Port"), 0, 0, 1, 2)
        self.port_combo = QComboBox()
        self.port_combo.setEditable(True)
        self.port_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        if self.port_combo.lineEdit() is not None:
            self.port_combo.lineEdit().setPlaceholderText("/dev/cu.usbserial-130")
        layout.addWidget(self.port_combo, 1, 0, 1, 2)

        self.refresh_ports_button = self._soft_button("Ports aktualisieren", self.refresh_ports)
        layout.addWidget(self.refresh_ports_button, 2, 0)
        self.simulate_checkbox = QCheckBox("Simulationsmodus auf dem Mac verwenden")
        self.simulate_checkbox.toggled.connect(self._update_source_inputs)
        layout.addWidget(self.simulate_checkbox, 2, 1)

        layout.addWidget(self._field_label("Baudrate"), 3, 0)
        self.baudrate_edit = QLineEdit("9600")
        layout.addWidget(self.baudrate_edit, 4, 0)

        self.connect_button = self._accent_button("Quelle starten", self.connect_source)
        self.disconnect_button = self._soft_button("Stoppen", self.disconnect_source)
        layout.addWidget(self.connect_button, 5, 0)
        layout.addWidget(self.disconnect_button, 5, 1)

        layout.addWidget(self._field_label("Status"), 6, 0, 1, 2)
        self.status_label = QLabel("Bereit.")
        self.status_label.setWordWrap(True)
        self.status_label.setObjectName("BodyCopy")
        layout.addWidget(self.status_label, 7, 0, 1, 2)
        return box

    def _build_capture_box(self) -> QGroupBox:
        box = QGroupBox("Erkennungslogik")
        layout = QGridLayout(box)
        layout.setHorizontalSpacing(10)
        layout.setVerticalSpacing(8)

        self.target_weight_edit = self._line_edit("12.50")
        self.target_window_edit = self._line_edit("0.50")
        self.stability_tolerance_edit = self._line_edit("0.05")
        self.stable_samples_edit = self._line_edit("6")
        self.rearm_threshold_edit = self._line_edit("0.10")
        self.minimum_weight_edit = self._line_edit("0.05")

        self._add_form_pair(layout, 0, "Zielgewicht (g)", self.target_weight_edit, "Fenster +/- (g)", self.target_window_edit)
        self._add_form_pair(
            layout,
            2,
            "Stabilitaets-Toleranz (g)",
            self.stability_tolerance_edit,
            "Benoetigte Samples",
            self.stable_samples_edit,
        )
        self._add_form_pair(
            layout,
            4,
            "Reset-Schwelle (g)",
            self.rearm_threshold_edit,
            "Mindestgewicht (g)",
            self.minimum_weight_edit,
        )
        self.require_confirmation_checkbox = QCheckBox("Vor dem Schreiben erst bestaetigen")
        layout.addWidget(self.require_confirmation_checkbox, 6, 0, 1, 2)
        layout.addWidget(self._soft_button("Erkennung neu starten", self.reset_capture_engine), 7, 0, 1, 2)
        return box

    def _build_excel_box(self) -> QGroupBox:
        box = QGroupBox("Excel-Ziel")
        layout = QGridLayout(box)
        layout.setHorizontalSpacing(10)
        layout.setVerticalSpacing(8)

        layout.addWidget(self._field_label("Excel-Modus"), 0, 0, 1, 2)
        self.excel_mode_combo = QComboBox()
        for value, label in excel_mode_options():
            self.excel_mode_combo.addItem(label, value)
        self.excel_mode_combo.currentIndexChanged.connect(self._refresh_excel_backend_hint)
        layout.addWidget(self.excel_mode_combo, 1, 0, 1, 2)

        self.excel_runtime_label = QLabel(live_backend_status_text())
        self.excel_runtime_label.setWordWrap(True)
        self.excel_runtime_label.setObjectName("BodyCopy")
        layout.addWidget(self.excel_runtime_label, 2, 0, 1, 2)

        layout.addWidget(self._field_label("Excel-Datei (.xlsx)"), 3, 0, 1, 2)
        self.excel_path_edit = QLineEdit()
        self.excel_path_edit.setPlaceholderText("/Pfad/zur/excel-test-datei.xlsx")
        layout.addWidget(self.excel_path_edit, 4, 0, 1, 2)

        layout.addWidget(self._soft_button("Datei waehlen", self.browse_excel_file), 5, 0)
        layout.addWidget(self._soft_button("Testdatei nutzen", self.use_project_excel_file), 5, 1)
        layout.addWidget(self._soft_button("Zielzelle pruefen", self.refresh_excel_target), 6, 0, 1, 2)

        self.sheet_name_edit = self._line_edit("Messwerte")
        self.column_edit = self._line_edit("A")
        self.start_row_edit = self._line_edit("2")
        self.current_row_edit = self._line_edit("2")

        self._add_form_pair(layout, 7, "Sheet", self.sheet_name_edit, "Spalte", self.column_edit)
        self._add_form_pair(layout, 9, "Startzeile", self.start_row_edit, "Aktuelle Zeile", self.current_row_edit)

        self.auto_advance_checkbox = QCheckBox("Nach jedem Eintrag eine Zeile weiter")
        self.auto_advance_checkbox.setChecked(True)
        layout.addWidget(self.auto_advance_checkbox, 11, 0, 1, 2)

        layout.addWidget(self._field_label("Naechste Zelle"), 12, 0)
        self.next_cell_value = QLabel("A2")
        self.next_cell_value.setObjectName("InlineValue")
        layout.addWidget(self.next_cell_value, 13, 0)

        layout.addWidget(self._field_label("Aktiver Writer"), 12, 1)
        self.excel_backend_label = QLabel("Auto")
        self.excel_backend_label.setWordWrap(True)
        self.excel_backend_label.setObjectName("BodyCopy")
        layout.addWidget(self.excel_backend_label, 13, 1)

        self.write_button = self._accent_button("Stabilen Wert schreiben", self.write_pending_capture)
        self.discard_button = self._soft_button("Messung verwerfen", self.discard_pending_capture)
        layout.addWidget(self.write_button, 14, 0)
        layout.addWidget(self.discard_button, 14, 1)
        return box

    def _build_metrics_row(self) -> QHBoxLayout:
        row = QHBoxLayout()
        row.setSpacing(14)

        self.live_weight_value, self.live_weight_detail, card = self._metric_card("Live-Wert")
        row.addWidget(card, 1)

        self.pending_value_label, self.capture_state_detail, card = self._metric_card("Stabile Messung")
        row.addWidget(card, 1)

        self.next_cell_metric_value, self.connection_detail_metric, card = self._metric_card("Naechste Zelle")
        row.addWidget(card, 1)

        self.capture_window_value, self.window_detail_metric, card = self._metric_card("Zielfenster")
        row.addWidget(card, 1)
        return row

    def _build_monitor_box(self) -> QGroupBox:
        box = QGroupBox("Monitoring")
        layout = QVBoxLayout(box)
        layout.setSpacing(12)

        info = QFrame()
        info_layout = QGridLayout(info)
        info_layout.setHorizontalSpacing(10)
        info_layout.setVerticalSpacing(8)

        info_layout.addWidget(self._field_label("Rohdaten"), 0, 0)
        self.raw_value_label = QLabel("-")
        self.raw_value_label.setWordWrap(True)
        self.raw_value_label.setObjectName("BodyCopy")
        info_layout.addWidget(self.raw_value_label, 0, 1)

        info_layout.addWidget(self._field_label("Verbindung"), 1, 0)
        self.connection_status_label = QLabel("Nicht verbunden")
        self.connection_status_label.setWordWrap(True)
        self.connection_status_label.setObjectName("BodyCopy")
        info_layout.addWidget(self.connection_status_label, 1, 1)

        layout.addWidget(info)

        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setObjectName("LogView")
        self.log_view.setMinimumHeight(320)
        layout.addWidget(self.log_view, 1)
        return box

    def _metric_card(self, title: str) -> tuple[QLabel, QLabel, QFrame]:
        card = QFrame()
        card.setObjectName("MetricCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(8)

        title_label = QLabel(title)
        title_label.setObjectName("MetricTitle")
        value_label = QLabel("--")
        value_label.setObjectName("MetricValue")
        detail_label = QLabel("-")
        detail_label.setWordWrap(True)
        detail_label.setObjectName("MetricDetail")

        layout.addWidget(title_label)
        layout.addWidget(value_label)
        layout.addWidget(detail_label)
        return value_label, detail_label, card

    def _field_label(self, text: str) -> QLabel:
        label = QLabel(text)
        label.setObjectName("FieldLabel")
        return label

    def _line_edit(self, value: str) -> QLineEdit:
        edit = QLineEdit(value)
        return edit

    def _add_form_pair(
        self,
        layout: QGridLayout,
        row: int,
        left_label: str,
        left_widget: QWidget,
        right_label: str,
        right_widget: QWidget,
    ) -> None:
        layout.addWidget(self._field_label(left_label), row, 0)
        layout.addWidget(self._field_label(right_label), row, 1)
        layout.addWidget(left_widget, row + 1, 0)
        layout.addWidget(right_widget, row + 1, 1)

    def _accent_button(self, text: str, handler) -> QPushButton:
        button = QPushButton(text)
        button.setObjectName("AccentButton")
        button.clicked.connect(handler)
        return button

    def _soft_button(self, text: str, handler) -> QPushButton:
        button = QPushButton(text)
        button.setObjectName("SoftButton")
        button.clicked.connect(handler)
        return button

    def _apply_styles(self) -> None:
        self.setStyleSheet(
            f"""
            QMainWindow {{
                background: {COLORS["page"]};
            }}
            QWidget {{
                color: {COLORS["ink"]};
                font-family: "Avenir Next", "Segoe UI", sans-serif;
                font-size: 10pt;
            }}
            QFrame, QGroupBox {{
                background: {COLORS["card"]};
                border: 1px solid {COLORS["line"]};
                border-radius: 18px;
            }}
            QGroupBox {{
                margin-top: 18px;
                padding-top: 20px;
                font-size: 12pt;
                font-weight: 700;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 16px;
                padding: 0 6px;
                color: {COLORS["ink"]};
            }}
            QLabel#HeroTitle {{
                font-size: 22pt;
                font-weight: 800;
                color: {COLORS["ink"]};
            }}
            QLabel#HeroCopy {{
                color: {COLORS["muted"]};
                font-size: 10.5pt;
            }}
            QLabel#FieldLabel {{
                color: {COLORS["muted"]};
                font-size: 9pt;
                font-weight: 700;
                border: none;
                background: transparent;
            }}
            QLabel#MetricTitle {{
                color: {COLORS["muted"]};
                font-size: 9pt;
                font-weight: 700;
                border: none;
                background: transparent;
            }}
            QLabel#MetricValue {{
                font-size: 20pt;
                font-weight: 800;
                color: {COLORS["ink"]};
                border: none;
                background: transparent;
            }}
            QLabel#MetricDetail, QLabel#BodyCopy {{
                color: {COLORS["ink"]};
                border: none;
                background: transparent;
            }}
            QLabel#InlineValue {{
                font-size: 16pt;
                font-weight: 800;
                color: {COLORS["ink"]};
                border: none;
                background: transparent;
            }}
            QFrame#MetricCard {{
                background: {COLORS["tile"]};
                border-radius: 20px;
                min-height: 150px;
            }}
            QLineEdit, QComboBox, QPlainTextEdit {{
                background: {COLORS["surface"]};
                border: 1px solid {COLORS["line"]};
                border-radius: 12px;
                padding: 8px 10px;
                selection-background-color: {COLORS["accent"]};
                font-size: 10.5pt;
            }}
            QLineEdit, QComboBox {{
                min-height: 30px;
            }}
            QPlainTextEdit#LogView {{
                font-family: "SF Mono", "Cascadia Code", monospace;
                font-size: 9.5pt;
            }}
            QPushButton {{
                border-radius: 12px;
                min-height: 36px;
                padding: 8px 14px;
                font-weight: 700;
                font-size: 10.5pt;
            }}
            QPushButton#AccentButton {{
                background: {COLORS["accent"]};
                color: white;
                border: none;
            }}
            QPushButton#AccentButton:hover {{
                background: {COLORS["accent_dark"]};
            }}
            QPushButton#SoftButton {{
                background: {COLORS["accent_soft"]};
                color: {COLORS["ink"]};
                border: none;
            }}
            QPushButton#SoftButton:hover {{
                background: #ffd4be;
            }}
            QCheckBox {{
                spacing: 8px;
                font-size: 10.5pt;
            }}
            """
        )

    def _refresh_excel_backend_hint(self) -> None:
        mode = self.excel_mode_combo.currentData() or EXCEL_MODE_AUTO
        if mode == EXCEL_MODE_AUTO:
            text = (
                f"{live_backend_status_text()} "
                "Im Auto-Modus versucht die App zuerst den Live-Writer und faellt sonst auf den Datei-Modus zurueck."
            )
        elif mode == "file":
            text = "Datei-Modus: Die .xlsx-Datei wird direkt auf dem Dateisystem geschrieben."
        else:
            text = (
                f"{live_backend_status_text()} "
                "Im Live-Modus schreibt die App direkt in die lokale Excel-Anwendung."
            )
        self.excel_runtime_label.setText(text)
        if self.excel_session is None or self.excel_session.settings.mode != mode:
            self._set_excel_backend(mode_label(mode))

    def refresh_ports(self) -> None:
        current = self.port_combo.currentText().strip()
        ports = list_serial_ports()
        if current and current not in ports:
            ports.append(current)

        self.port_combo.clear()
        self.port_combo.addItems(ports)
        if current:
            self.port_combo.setCurrentText(current)
        elif ports:
            self.port_combo.setCurrentText(ports[0])
        elif not self.simulate_checkbox.isChecked():
            self._set_status("Keine seriellen Ports gefunden. Du kannst den Port auch manuell eintragen.")

    def browse_excel_file(self) -> None:
        start_path = self.excel_path_edit.text().strip()
        if not start_path:
            start_path = str(Path.cwd())

        path, _ = QFileDialog.getOpenFileName(
            self,
            "Excel-Datei auswaehlen",
            start_path,
            "Excel-Datei (*.xlsx)",
        )
        if not path:
            return

        self.excel_path_edit.setText(path)
        self.refresh_excel_target()

    def use_project_excel_file(self) -> None:
        project_file = Path.cwd() / "excel-test-datei.xlsx"
        if not project_file.exists():
            QMessageBox.information(
                self,
                "Datei nicht gefunden",
                f"Im Projektordner wurde keine Datei unter\n{project_file}\ngefunden.",
            )
            return

        self.excel_path_edit.setText(str(project_file))
        self.refresh_excel_target()

    def connect_source(self) -> None:
        if self.stream_worker and self.stream_worker.is_alive():
            return

        try:
            serial_settings = self._collect_serial_settings()
            self.capture_engine.update_settings(self._collect_capture_settings())
        except ValueError as exc:
            QMessageBox.critical(self, "Ungueltige Eingabe", str(exc))
            return

        self._refresh_capture_window_label()
        self._set_pending_value("--")
        self._set_capture_state("Verbinde Quelle")
        self._set_status("Quelle wird gestartet...")

        self.stream_worker = ScaleStreamWorker(
            serial_settings,
            simulate=self.simulate_checkbox.isChecked(),
            simulation_target=self.capture_engine.settings.target_weight,
        )
        self.stream_worker.start()
        self._save_settings()

    def disconnect_source(self) -> None:
        if self.stream_worker is not None:
            self.stream_worker.stop()
            self._set_status("Quelle wird gestoppt...")

    def reset_capture_engine(self) -> None:
        try:
            self.capture_engine.update_settings(self._collect_capture_settings())
        except ValueError as exc:
            QMessageBox.critical(self, "Ungueltige Eingabe", str(exc))
            return

        self._set_pending_value("--")
        self._set_capture_state("Warte auf Gewicht")
        self._refresh_capture_window_label()
        self._log("Erkennung zurueckgesetzt.")

    def refresh_excel_target(self) -> None:
        try:
            session = self._sync_excel_session()
        except Exception as exc:
            QMessageBox.critical(self, "Excel-Konfiguration", str(exc))
            return

        self.current_row_edit.setText(str(session.current_row or 0))
        self._set_next_cell(session.preview_cell())
        self._set_excel_backend(session.backend_display_name())
        self._set_status(f"Excel-Ziel bereit: {session.preview_cell()} | {session.backend_display_name()}")

    def write_pending_capture(self) -> None:
        self._write_pending_capture(auto=False)

    def discard_pending_capture(self) -> None:
        if self.capture_engine.peek_pending_capture() is None:
            return

        self.capture_engine.discard_pending_capture()
        self._set_pending_value("--")
        self._set_capture_state("Messung verworfen")
        self._log("Stabile Messung verworfen.")

    def _write_pending_capture(self, auto: bool) -> None:
        pending_value = self.capture_engine.peek_pending_capture()
        if pending_value is None:
            if not auto:
                QMessageBox.information(self, "Keine stabile Messung", "Aktuell liegt noch keine stabile Messung vor.")
            return

        try:
            session = self._sync_excel_session()
            result = session.write_value(pending_value)
        except Exception as exc:
            self._set_status(str(exc))
            self._log(f"Excel-Fehler: {exc}")
            if not auto:
                QMessageBox.critical(self, "Excel-Fehler", str(exc))
            return

        committed_value = self.capture_engine.commit_pending_capture()
        self._set_pending_value("--")
        self._set_capture_state("Messung gespeichert")
        self.current_row_edit.setText(str(session.current_row or result.row))
        self._set_next_cell(session.preview_cell())
        self._set_excel_backend(session.backend_display_name())
        self._set_status(f"Messwert via {session.backend_display_name()} in {result.sheet_name}!{result.cell} gespeichert.")
        if committed_value is not None:
            self._log(f"{committed_value:.3f} -> {result.sheet_name}!{result.cell} ({session.backend_display_name()})")
        self._save_settings()

    def _poll_worker(self) -> None:
        if self.stream_worker is not None:
            while True:
                try:
                    event = self.stream_worker.events.get_nowait()
                except queue.Empty:
                    break
                self._handle_stream_event(event)

    def _handle_stream_event(self, event: StreamEvent) -> None:
        if event.kind == "connected":
            self._set_connection_state(event.message)
            self._set_capture_state("Warte auf stabile Werte")
            self._set_status(event.message)
            self._log(event.message)
            return

        if event.kind == "measurement" and event.measurement is not None:
            self._handle_measurement(event)
            return

        if event.kind == "raw":
            self._set_raw_text(event.raw_text or "-")
            return

        if event.kind == "error":
            self._set_connection_state("Fehler")
            self._set_status(event.message)
            self._log(f"Fehler: {event.message}")
            QMessageBox.critical(self, "Serielle Verbindung", event.message)
            return

        if event.kind == "stopped":
            self._set_connection_state("Nicht verbunden")
            self._set_status("Quelle gestoppt.")
            self._log("Quelle gestoppt.")
            self.stream_worker = None

    def _handle_measurement(self, event: StreamEvent) -> None:
        measurement = event.measurement
        if measurement is None:
            return

        self._set_live_weight(f"{measurement.value:.3f}", measurement.unit or "-")
        self._set_raw_text(event.raw_text or measurement.raw_text)

        state = self.capture_engine.process(measurement)
        self._update_capture_dashboard(state)

        if state.rearmed:
            self._log("Waage zurueckgesetzt. Naechste Waegung kann erfasst werden.")

        if state.new_candidate is not None:
            self._set_pending_value(f"{state.new_candidate:.3f}")
            self._set_capture_state("Stabile Messung erkannt")
            self._log(f"Stabiler Wert erkannt: {state.new_candidate:.3f}")
            if not self.require_confirmation_checkbox.isChecked():
                self._write_pending_capture(auto=True)

    def _update_capture_dashboard(self, state: CaptureState) -> None:
        if state.pending_capture is not None:
            self._set_pending_value(f"{state.pending_capture:.3f}")
        elif self.capture_engine.peek_pending_capture() is None:
            self._set_pending_value("--")

        if state.pending_capture is not None:
            self._set_capture_state("Stabil erkannt - bereit zum Schreiben")
            return

        if not state.armed:
            self._set_capture_state("Bitte Gewicht entfernen")
            return

        if state.stable and state.within_target:
            self._set_capture_state("Stabil im Zielbereich")
            return

        if state.within_target:
            self._set_capture_state("Im Zielbereich - warte auf Stabilitaet")
            return

        if abs(state.measurement.value) < self.capture_engine.settings.minimum_weight:
            self._set_capture_state("Warte auf Gewicht")
            return

        self._set_capture_state("Ausserhalb des Zielbereichs")

    def _update_source_inputs(self) -> None:
        self.port_combo.setEnabled(not self.simulate_checkbox.isChecked())

    def _collect_serial_settings(self) -> SerialSettings:
        if not self.simulate_checkbox.isChecked() and not self.port_combo.currentText().strip():
            raise ValueError("Bitte einen seriellen Port angeben oder den Simulationsmodus aktivieren.")

        return SerialSettings(
            port=self.port_combo.currentText().strip(),
            baudrate=self._parse_int(self.baudrate_edit.text(), "Baudrate"),
            timeout=1.0,
        )

    def _collect_capture_settings(self) -> CaptureSettings:
        target_text = self.target_weight_edit.text().strip()
        target_weight = None if not target_text else self._parse_float(target_text, "Zielgewicht")

        return CaptureSettings(
            target_weight=target_weight,
            target_window=self._parse_float(self.target_window_edit.text(), "Fenster"),
            stability_tolerance=self._parse_float(self.stability_tolerance_edit.text(), "Stabilitaets-Toleranz"),
            stable_samples=self._parse_int(self.stable_samples_edit.text(), "Samples"),
            rearm_threshold=self._parse_float(self.rearm_threshold_edit.text(), "Reset-Schwelle"),
            minimum_weight=self._parse_float(self.minimum_weight_edit.text(), "Mindestgewicht"),
            require_confirmation=self.require_confirmation_checkbox.isChecked(),
        )

    def _collect_excel_settings(self) -> ExcelSettings:
        return ExcelSettings(
            path=self.excel_path_edit.text().strip(),
            sheet_name=self.sheet_name_edit.text().strip() or "Messwerte",
            column=self.column_edit.text().strip().upper() or "A",
            start_row=self._parse_int(self.start_row_edit.text(), "Startzeile"),
            auto_advance=self.auto_advance_checkbox.isChecked(),
            mode=self.excel_mode_combo.currentData() or EXCEL_MODE_AUTO,
        )

    def _sync_excel_session(self) -> ExcelSession:
        settings = self._collect_excel_settings()
        preserve_row = self.excel_session is not None and self.excel_session.settings == settings

        if self.excel_session is None:
            self.excel_session = ExcelSession(settings)
        else:
            self.excel_session.update_settings(settings, preserve_row=preserve_row)

        current_row_text = self.current_row_edit.text().strip()
        if current_row_text:
            self.excel_session.reset_row(self._parse_int(current_row_text, "Aktuelle Zeile"))
        elif settings.path.strip():
            self.current_row_edit.setText(str(self.excel_session.detect_current_row()))
        else:
            self.excel_session.reset_row(settings.start_row)
            self.current_row_edit.setText(str(self.excel_session.current_row or settings.start_row))

        self._set_next_cell(self.excel_session.preview_cell())
        self._set_excel_backend(self.excel_session.backend_display_name())
        return self.excel_session

    def _refresh_capture_window_label(self) -> None:
        bounds = self.capture_engine.window_bounds()
        if bounds is None:
            text = f"frei ab {self.capture_engine.settings.minimum_weight:.2f}"
        else:
            text = f"{bounds[0]:.2f} bis {bounds[1]:.2f}"
        self.capture_window_value.setText(text)

    def _parse_float(self, value: str, label: str) -> float:
        try:
            return float(value.strip().replace(",", "."))
        except ValueError as exc:
            raise ValueError(f"{label} ist keine gueltige Zahl.") from exc

    def _parse_int(self, value: str, label: str) -> int:
        try:
            number = int(value.strip())
        except ValueError as exc:
            raise ValueError(f"{label} ist keine gueltige ganze Zahl.") from exc

        if number < 1:
            raise ValueError(f"{label} muss groesser als 0 sein.")
        return number

    def _set_status(self, text: str) -> None:
        self.status_label.setText(text)
        self.window_detail_metric.setText(text)

    def _set_connection_state(self, text: str) -> None:
        self.connection_status_label.setText(text)
        self.connection_detail_metric.setText(text)

    def _set_capture_state(self, text: str) -> None:
        self.capture_state_detail.setText(text)

    def _set_live_weight(self, value_text: str, unit_text: str) -> None:
        self.live_weight_value.setText(value_text)
        self.live_weight_detail.setText(unit_text)

    def _set_pending_value(self, text: str) -> None:
        self.pending_value_label.setText(text)

    def _set_next_cell(self, text: str) -> None:
        self.next_cell_value.setText(text)
        self.next_cell_metric_value.setText(text)

    def _set_excel_backend(self, text: str) -> None:
        self.excel_backend_label.setText(text)

    def _set_raw_text(self, text: str) -> None:
        self.raw_value_label.setText(text)

    def _log(self, message: str) -> None:
        self.log_view.appendPlainText(message)

    def _load_settings(self) -> None:
        data = self.settings_store.load()
        if not data:
            self.simulate_checkbox.setChecked(True)
            default_mode_index = self.excel_mode_combo.findData(EXCEL_MODE_AUTO)
            if default_mode_index >= 0:
                self.excel_mode_combo.setCurrentIndex(default_mode_index)
            self._update_source_inputs()
            self._refresh_excel_backend_hint()
            return

        self.port_combo.setCurrentText(data.get("port", ""))
        self.baudrate_edit.setText(str(data.get("baudrate", "9600")))
        self.simulate_checkbox.setChecked(bool(data.get("simulate", True)))

        self.target_weight_edit.setText(str(data.get("target_weight", "12.50")))
        self.target_window_edit.setText(str(data.get("target_window", "0.50")))
        self.stability_tolerance_edit.setText(str(data.get("stability_tolerance", "0.05")))
        self.stable_samples_edit.setText(str(data.get("stable_samples", "6")))
        self.rearm_threshold_edit.setText(str(data.get("rearm_threshold", "0.10")))
        self.minimum_weight_edit.setText(str(data.get("minimum_weight", "0.05")))
        self.require_confirmation_checkbox.setChecked(bool(data.get("require_confirmation", False)))

        self.excel_path_edit.setText(str(data.get("excel_path", "")))
        excel_mode = str(data.get("excel_mode", EXCEL_MODE_AUTO))
        excel_mode_index = self.excel_mode_combo.findData(excel_mode)
        if excel_mode_index >= 0:
            self.excel_mode_combo.setCurrentIndex(excel_mode_index)
        self.sheet_name_edit.setText(str(data.get("sheet_name", "Messwerte")))
        self.column_edit.setText(str(data.get("column", "A")))
        self.start_row_edit.setText(str(data.get("start_row", "2")))
        self.current_row_edit.setText(str(data.get("current_row", "2")))
        self.auto_advance_checkbox.setChecked(bool(data.get("auto_advance", True)))
        self._update_source_inputs()
        self._refresh_excel_backend_hint()

    def _save_settings(self) -> None:
        self.settings_store.save(
            {
                "port": self.port_combo.currentText().strip(),
                "baudrate": self.baudrate_edit.text().strip(),
                "simulate": self.simulate_checkbox.isChecked(),
                "target_weight": self.target_weight_edit.text().strip(),
                "target_window": self.target_window_edit.text().strip(),
                "stability_tolerance": self.stability_tolerance_edit.text().strip(),
                "stable_samples": self.stable_samples_edit.text().strip(),
                "rearm_threshold": self.rearm_threshold_edit.text().strip(),
                "minimum_weight": self.minimum_weight_edit.text().strip(),
                "require_confirmation": self.require_confirmation_checkbox.isChecked(),
                "excel_path": self.excel_path_edit.text().strip(),
                "excel_mode": self.excel_mode_combo.currentData() or EXCEL_MODE_AUTO,
                "sheet_name": self.sheet_name_edit.text().strip(),
                "column": self.column_edit.text().strip(),
                "start_row": self.start_row_edit.text().strip(),
                "current_row": self.current_row_edit.text().strip(),
                "auto_advance": self.auto_advance_checkbox.isChecked(),
            }
        )

    def closeEvent(self, event: QCloseEvent) -> None:
        self._save_settings()
        if self.stream_worker is not None:
            self.stream_worker.stop()
        super().closeEvent(event)


def main() -> None:
    app = QApplication.instance() or QApplication(sys.argv)
    window = ScaleLoggerWindow()
    window.show()
    app.exec()


if __name__ == "__main__":
    main()
