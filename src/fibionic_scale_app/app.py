from __future__ import annotations

import queue
import sys
from pathlib import Path

from PySide6.QtCore import Qt, QTimer, QUrl
from PySide6.QtGui import QCloseEvent, QDesktopServices, QPainter, QPixmap
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtWidgets import (
    QApplication,
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
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from .excel_writer import (
    EXCEL_MODE_AUTO,
    ExcelSession,
    build_cell_ref,
    scan_direction_options,
    workbook_path_block_reason,
)
from .models import ExcelSettings, FLOW_DOWN, SerialSettings
from .serial_io import (
    auto_detectable_serial_ports,
    SIM_PROFILE_STABLE,
    SOURCE_MODE_SERIAL,
    SOURCE_MODE_SIMULATION,
    ScaleSource,
    SerialScaleSource,
    SimulatedScaleSource,
    StreamEvent,
    list_serial_ports,
    preferred_serial_port,
    verified_serial_port,
    simulation_profile_label,
    simulation_profile_options,
    source_mode_options,
)
from .settings_store import SettingsStore
from .stability import CaptureState, WeightCaptureEngine, build_capture_settings

COLORS = {
    "page": "#E3E1E6",
    "rail": "#EEF0F3",
    "card": "#FFFFFF",
    "tile": "#F6FBFD",
    "surface": "#FFFFFF",
    "ink": "#1D1D1D",
    "muted": "#5D646D",
    "accent": "#44B6CD",
    "accent_dark": "#258EA5",
    "accent_soft": "#DFF2F6",
    "signal": "#FFE800",
    "signal_dark": "#D7C900",
    "danger": "#D94C4C",
    "danger_dark": "#B13A3A",
    "line": "#C9CED6",
}

SOURCE_CONTROL_IDLE = "idle"
SOURCE_CONTROL_RUNNING = "running"
SOURCE_CONTROL_PAUSED = "paused"


class ScaleLoggerWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("fibionic | Gewichtslogging")
        self.resize(1440, 900)
        self.setMinimumSize(1280, 800)

        self.settings_store = SettingsStore()
        self.scale_source: ScaleSource | None = None
        self.capture_engine = WeightCaptureEngine(build_capture_settings(12.5, 0.5))
        self.excel_session: ExcelSession | None = None
        self.available_ports: list[str] = []
        self.detected_port = ""
        self.verified_port = ""
        self.manual_port_override = False
        self.paused = False
        self.last_excel_error: str | None = None
        self._saved_manual_port = ""
        self._flash_widgets: tuple[QWidget, ...] = ()
        self.source_control_state = SOURCE_CONTROL_IDLE
        self._syncing_excel_cursor = False
        self._unit_error: str | None = None
        self._last_logged_value: float | None = None

        self._build_ui()
        self._refresh_source_controls()
        self._apply_styles()
        self._load_settings()
        self.refresh_ports()
        self._refresh_auto_capture_hint()
        self._refresh_target_range_display()
        self._update_source_mode_ui()
        self._set_running_state(False)
        self._set_stage(
            "Bereit zum Start",
            "Zielgewicht, Abweichung und Excel-Ziel setzen. Danach Quelle starten und auf den Datenstrom warten.",
        )
        self._refresh_excel_target(silent=True)

        self.poll_timer = QTimer(self)
        self.poll_timer.setInterval(120)
        self.poll_timer.timeout.connect(self._poll_source_events)
        self.poll_timer.start()

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)

        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(16, 12, 16, 12)
        root_layout.setSpacing(12)

        self.header_panel = self._build_header_panel()
        root_layout.addWidget(self.header_panel, 0)

        content_layout = QHBoxLayout()
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(12)
        root_layout.addLayout(content_layout, 1)

        self.left_rail = QFrame()
        self.left_rail.setObjectName("LeftRail")
        left_layout = QVBoxLayout(self.left_rail)
        left_layout.setContentsMargins(10, 10, 10, 10)
        left_layout.setSpacing(10)

        self.left_scroll = QScrollArea()
        self.left_scroll.setObjectName("LeftScroll")
        self.left_scroll.setWidgetResizable(True)
        self.left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.left_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.left_scroll.setWidget(self.left_rail)
        self.left_scroll.setMinimumWidth(440)
        self.left_scroll.setMaximumWidth(520)

        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(10)

        content_layout.addWidget(self.left_scroll, 0)
        content_layout.addWidget(right_panel, 1)

        self.scale_box = self._build_scale_box()
        self.setup_panel = QWidget()
        self.setup_panel.setObjectName("SetupPanel")
        setup_layout = QVBoxLayout(self.setup_panel)
        setup_layout.setContentsMargins(0, 0, 0, 0)
        setup_layout.setSpacing(10)
        self.capture_box = self._build_capture_box()
        self.excel_box = self._build_excel_box()
        setup_layout.addWidget(self.capture_box)
        setup_layout.addWidget(self.excel_box)
        setup_layout.addStretch(1)

        left_layout.addWidget(self.scale_box)
        left_layout.addWidget(self.setup_panel)
        left_layout.addStretch(1)

        right_layout.addWidget(self._build_status_panel())
        right_layout.addWidget(self._build_monitor_box(), 1)

    def _build_scale_box(self) -> QGroupBox:
        box = QGroupBox("Quelle")
        layout = QVBoxLayout(box)
        layout.setContentsMargins(12, 14, 12, 12)
        layout.setSpacing(10)

        self.connection_setup_panel = QWidget()
        self.connection_setup_panel.setObjectName("ConnectionSetup")
        setup_layout = QVBoxLayout(self.connection_setup_panel)
        setup_layout.setContentsMargins(0, 0, 0, 0)
        setup_layout.setSpacing(8)

        self.source_mode_combo, source_shell = self._combo_field()
        for value, label in source_mode_options():
            self.source_mode_combo.addItem(label, value)
        self.source_mode_combo.currentIndexChanged.connect(self._update_source_mode_ui)
        setup_layout.addWidget(source_shell)

        self.serial_config_panel = QWidget()
        serial_layout = QGridLayout(self.serial_config_panel)
        serial_layout.setContentsMargins(0, 0, 0, 0)
        serial_layout.setHorizontalSpacing(8)
        serial_layout.setVerticalSpacing(6)

        self.detected_port_label = QLabel("Noch keine Waage erkannt")
        self.detected_port_label.setObjectName("InlineValue")
        self.detected_port_label.setWordWrap(True)
        serial_layout.addWidget(self.detected_port_label, 0, 0, 1, 2)

        self.auto_port_button = self._soft_button("Automatisch erkennen", self.use_auto_port_selection)
        self.manual_port_button = self._soft_button("Port manuell wählen", self.toggle_manual_port_selection)
        serial_layout.addWidget(self.auto_port_button, 1, 0)
        serial_layout.addWidget(self.manual_port_button, 1, 1)

        self.manual_port_combo, self.manual_port_shell = self._combo_field(editable=True)
        if self.manual_port_combo.lineEdit() is not None:
            self.manual_port_combo.lineEdit().setPlaceholderText("/dev/cu.usbserial-130")
        serial_layout.addWidget(self.manual_port_shell, 2, 0, 1, 2)
        setup_layout.addWidget(self.serial_config_panel)

        self.simulation_config_panel = QWidget()
        simulation_layout = QGridLayout(self.simulation_config_panel)
        simulation_layout.setContentsMargins(0, 0, 0, 0)
        simulation_layout.setHorizontalSpacing(8)
        simulation_layout.setVerticalSpacing(6)

        simulation_layout.addWidget(self._field_label("Simulationsprofil"), 0, 0)
        self.simulation_profile_combo, simulation_shell = self._combo_field()
        for value, label in simulation_profile_options():
            self.simulation_profile_combo.addItem(label, value)
        self.simulation_profile_combo.currentIndexChanged.connect(self._update_source_mode_ui)
        simulation_layout.addWidget(simulation_shell, 1, 0)

        self.simulation_hint_label = QLabel(
            "Virtuelle Waage für App-, Stabilitäts- und Excel-Tests ohne echte Hardware."
        )
        self.simulation_hint_label.setObjectName("BodyCopy")
        self.simulation_hint_label.setWordWrap(True)
        simulation_layout.addWidget(self.simulation_hint_label, 2, 0)
        setup_layout.addWidget(self.simulation_config_panel)

        layout.addWidget(self.connection_setup_panel)

        self.active_source_panel = QWidget()
        self.active_source_panel.setVisible(False)
        active_layout = QVBoxLayout(self.active_source_panel)
        active_layout.setContentsMargins(0, 0, 0, 0)
        active_layout.setSpacing(4)
        self.active_source_value = QLabel("--")
        self.active_source_value.setObjectName("InlineValue")
        self.active_source_value.setWordWrap(True)
        active_layout.addWidget(self.active_source_value)
        layout.addWidget(self.active_source_panel)

        controls_layout = QHBoxLayout()
        controls_layout.setContentsMargins(0, 0, 0, 0)
        controls_layout.setSpacing(10)
        self.primary_source_button = self._accent_button("Quelle starten", self._handle_primary_source_action)
        self.stop_button = self._danger_button("Stopp", self.stop_source)
        controls_layout.addWidget(self.primary_source_button, 1)
        controls_layout.addWidget(self.stop_button, 1)
        layout.addLayout(controls_layout)

        self.connection_note_label = QLabel("Quelle noch nicht gestartet.")
        self.connection_note_label.setWordWrap(True)
        self.connection_note_label.setObjectName("BodyCopy")
        layout.addWidget(self.connection_note_label)

        return box

    def _build_capture_box(self) -> QGroupBox:
        box = QGroupBox("Messwerte")
        layout = QGridLayout(box)
        layout.setContentsMargins(12, 14, 12, 12)
        layout.setHorizontalSpacing(8)
        layout.setVerticalSpacing(6)

        self.target_weight_edit = self._line_edit("12.50")
        self.target_window_edit = self._line_edit("0.50")
        self.target_weight_edit.textChanged.connect(self._refresh_target_range_display)
        self.target_window_edit.textChanged.connect(self._refresh_target_range_display)
        self.target_weight_edit.textChanged.connect(self._refresh_auto_capture_hint)
        self.target_window_edit.textChanged.connect(self._refresh_auto_capture_hint)
        self.target_weight_edit.editingFinished.connect(self._apply_runtime_target_changes)
        self.target_window_edit.editingFinished.connect(self._apply_runtime_target_changes)

        self._add_form_pair(
            layout,
            0,
            "Zielgewicht (g)",
            self.target_weight_edit,
            "Abweichung +/- (g)",
            self.target_window_edit,
        )

        self.auto_capture_hint = QLabel()
        self.auto_capture_hint.setObjectName("BodyCopy")
        self.auto_capture_hint.setWordWrap(True)
        layout.addWidget(self.auto_capture_hint, 2, 0, 1, 2)
        return box

    def _build_excel_box(self) -> QGroupBox:
        box = QGroupBox("Excel")
        layout = QGridLayout(box)
        layout.setContentsMargins(12, 14, 12, 12)
        layout.setHorizontalSpacing(8)
        layout.setVerticalSpacing(6)

        self.excel_path_edit = QLineEdit()
        self.excel_path_edit.setVisible(False)
        self.excel_file_name_label = QLabel("")
        self.excel_file_name_label.setObjectName("InlineValue")
        self.excel_file_name_label.setWordWrap(True)
        layout.addWidget(self.excel_file_name_label, 0, 0, 1, 2)

        self.excel_buttons_row = QWidget()
        excel_buttons_layout = QHBoxLayout(self.excel_buttons_row)
        excel_buttons_layout.setContentsMargins(0, 0, 0, 0)
        excel_buttons_layout.setSpacing(8)
        self.browse_excel_button = self._soft_button("Datei auswählen", self.browse_excel_file)
        self.open_excel_button = self._soft_button("Datei öffnen", self.open_excel_file)
        excel_buttons_layout.addWidget(self.browse_excel_button, 1)
        excel_buttons_layout.addWidget(self.open_excel_button, 1)
        layout.addWidget(self.excel_buttons_row, 1, 0, 1, 2)

        self.sheet_name_edit = self._line_edit("Messwerte")
        self.column_edit = self._line_edit("A")
        self.start_row_edit = self._line_edit("2")
        self.sheet_name_edit.editingFinished.connect(self._handle_excel_settings_changed)
        self.column_edit.editingFinished.connect(self._handle_excel_settings_changed)
        self.start_row_edit.editingFinished.connect(self._handle_excel_settings_changed)

        self._add_form_pair(layout, 2, "Sheet", self.sheet_name_edit, "Spalte", self.column_edit)

        layout.addWidget(self._field_label("Zeile"), 4, 0)
        self.direction_combo, direction_shell = self._combo_field()
        for value, label in scan_direction_options():
            self.direction_combo.addItem(label, value)
        self.direction_combo.currentIndexChanged.connect(self._handle_excel_settings_changed)
        layout.addWidget(self._field_label("Logging-Format"), 4, 1)
        layout.addWidget(self.start_row_edit, 5, 0)
        layout.addWidget(direction_shell, 5, 1)

        return box

    def _build_header_panel(self) -> QFrame:
        panel = QFrame()
        panel.setObjectName("HeaderPanel")
        layout = QHBoxLayout(panel)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(14)

        copy = QVBoxLayout()
        copy.setSpacing(2)

        brand = QLabel("fibionic gmbh")
        brand.setObjectName("BrandLabel")
        title = QLabel("Gewichtslogging")
        title.setObjectName("HeaderTitle")

        copy.addWidget(brand)
        copy.addWidget(title)
        layout.addLayout(copy, 1)

        layout.addWidget(self._build_header_mark(), 0)

        return panel

    def _build_header_mark(self) -> QWidget:
        logo_root = Path(__file__).resolve().parents[2] / "logo"
        pixmap = self._load_header_logo_pixmap(logo_root)
        if pixmap is not None:
            logo_wrap = QWidget()
            logo_layout = QHBoxLayout(logo_wrap)
            logo_layout.setContentsMargins(10, 6, 10, 6)
            logo_layout.setSpacing(0)
            logo_label = QLabel()
            logo_label.setObjectName("HeaderLogo")
            logo_label.setPixmap(
                pixmap.scaledToHeight(56, Qt.TransformationMode.SmoothTransformation)
            )
            logo_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            logo_layout.addWidget(logo_label, 0, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            return logo_wrap

        ornament = QWidget()
        ornament_layout = QHBoxLayout(ornament)
        ornament_layout.setContentsMargins(0, 0, 0, 0)
        ornament_layout.setSpacing(8)
        for index, height in enumerate((34, 54, 42, 26), start=1):
            line = QFrame()
            line.setObjectName("FiberLine")
            if index == 2:
                line.setProperty("variant", "accent")
            elif index == 4:
                line.setProperty("variant", "signal")
            line.setFixedSize(3, height)
            ornament_layout.addWidget(line)
        return ornament

    def _load_header_logo_pixmap(self, logo_root: Path) -> QPixmap | None:
        svg_path = logo_root / "Logo_Fibionic_4c.svg"
        if svg_path.exists():
            renderer = QSvgRenderer(str(svg_path))
            if renderer.isValid():
                size = renderer.defaultSize()
                target_height = max(64, size.height() or 64)
                target_width = max(160, int((size.width() or 220) * (target_height / max(size.height() or 1, 1))))
                pixmap = QPixmap(target_width, target_height)
                pixmap.fill(Qt.GlobalColor.transparent)
                painter = QPainter(pixmap)
                renderer.render(painter)
                painter.end()
                return pixmap

        png_path = logo_root / "Logo_Fibionic_4c.png"
        if png_path.exists():
            pixmap = QPixmap(str(png_path))
            if not pixmap.isNull():
                return self._trim_logo_pixmap(pixmap)

        return None

    def _trim_logo_pixmap(self, pixmap: QPixmap) -> QPixmap:
        image = pixmap.toImage()
        min_x = image.width()
        min_y = image.height()
        max_x = -1
        max_y = -1

        for y in range(image.height()):
            for x in range(image.width()):
                color = image.pixelColor(x, y)
                if color.red() > 245 and color.green() > 245 and color.blue() > 245:
                    continue
                min_x = min(min_x, x)
                min_y = min(min_y, y)
                max_x = max(max_x, x)
                max_y = max(max_y, y)

        if max_x < min_x or max_y < min_y:
            return pixmap

        return pixmap.copy(min_x, min_y, (max_x - min_x) + 1, (max_y - min_y) + 1)

    def _build_status_panel(self) -> QFrame:
        panel = QFrame()
        panel.setObjectName("StatusPanel")
        self.status_panel = panel
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(12)

        self.stage_label = QLabel("Bereit")
        self.stage_label.setObjectName("StageLabel")
        self.stage_copy_label = QLabel("")
        self.stage_copy_label.setObjectName("StageCopy")
        self.stage_copy_label.setWordWrap(True)
        layout.addWidget(self.stage_label)
        layout.addWidget(self.stage_copy_label)

        stats_row = QHBoxLayout()
        stats_row.setSpacing(10)
        self.live_weight_value, self.live_weight_detail, self.live_card = self._metric_card("Live-Wert")
        self.pending_value_label, self.pending_detail_label, self.pending_card = self._metric_card("Stabiler Messwert")
        self.next_cell_value, self.next_cell_detail_label, self.next_cell_card = self._metric_card("Nächste Zelle")
        self._flash_widgets = (self.pending_card,)
        stats_row.addWidget(self.live_card, 1)
        stats_row.addWidget(self.pending_card, 1)
        stats_row.addWidget(self.next_cell_card, 1)
        layout.addLayout(stats_row)

        meta = QGridLayout()
        meta.setHorizontalSpacing(16)
        meta.setVerticalSpacing(6)
        meta.addWidget(self._field_label("Zielbereich"), 0, 0)
        self.target_range_value = QLabel("--")
        self.target_range_value.setObjectName("BodyCopy")
        meta.addWidget(self.target_range_value, 0, 1)

        meta.addWidget(self._field_label("Logging-Format"), 0, 2)
        self.logging_format_value = QLabel("--")
        self.logging_format_value.setObjectName("BodyCopy")
        meta.addWidget(self.logging_format_value, 0, 3)
        layout.addLayout(meta)

        self.backend_value = QLabel("Auto")
        self.backend_value.hide()
        self.raw_value_label = QLabel("-")
        self.raw_value_label.hide()

        return panel

    def _build_monitor_box(self) -> QGroupBox:
        box = QGroupBox("Verlauf")
        layout = QVBoxLayout(box)
        layout.setContentsMargins(12, 14, 12, 12)
        layout.setSpacing(6)

        header_row = QHBoxLayout()
        header_row.setContentsMargins(0, 0, 0, 0)
        header_row.setSpacing(12)
        header_row.addStretch(1)

        self.clear_log_button = self._utility_button("Verlauf löschen", self._clear_log_history)
        self.clear_log_button.setEnabled(False)
        header_row.addWidget(self.clear_log_button, 0, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)
        layout.addLayout(header_row)

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
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(5)

        title_label = QLabel(title)
        title_label.setObjectName("MetricTitle")
        value_label = QLabel("--")
        value_label.setObjectName("MetricValue")
        detail_label = QLabel("-")
        detail_label.setObjectName("MetricDetail")
        detail_label.setWordWrap(True)

        layout.addWidget(title_label)
        layout.addWidget(value_label)
        layout.addWidget(detail_label)
        return value_label, detail_label, card

    def _field_label(self, text: str) -> QLabel:
        label = QLabel(text)
        label.setObjectName("FieldLabel")
        return label

    def _line_edit(self, value: str) -> QLineEdit:
        return QLineEdit(value)

    def _combo_field(self, editable: bool = False) -> tuple[QComboBox, QFrame]:
        combo = QComboBox()
        combo.setObjectName("FormCombo")
        combo.setEditable(editable)
        if editable and combo.lineEdit() is not None:
            combo.lineEdit().setPlaceholderText("")

        shell = QFrame()
        shell.setObjectName("FieldShell")
        shell_layout = QHBoxLayout(shell)
        shell_layout.setContentsMargins(2, 1, 2, 1)
        shell_layout.setSpacing(0)
        arrow_button = QToolButton(shell)
        arrow_button.setObjectName("ComboArrowButton")
        arrow_button.setCursor(Qt.CursorShape.PointingHandCursor)
        arrow_button.setText("\u2304")
        arrow_button.clicked.connect(combo.showPopup)
        shell_layout.addWidget(combo, 1)
        shell_layout.addWidget(arrow_button, 0)
        return combo, shell

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

    def _danger_button(self, text: str, handler) -> QPushButton:
        button = QPushButton(text)
        button.setObjectName("DangerButton")
        button.clicked.connect(handler)
        return button

    def _utility_button(self, text: str, handler) -> QPushButton:
        button = QPushButton(text)
        button.setObjectName("SoftButton")
        button.setFixedHeight(30)
        button.setMinimumWidth(148)
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
                font-family: "Avenir Next LT Pro", "Avenir Next", "Segoe UI", "Helvetica Neue", sans-serif;
                font-size: 10pt;
            }}
            QScrollArea#LeftScroll {{
                background: transparent;
                border: none;
            }}
            QScrollArea#LeftScroll > QWidget > QWidget {{
                background: transparent;
                border: none;
            }}
            QFrame#LeftRail {{
                background: {COLORS["rail"]};
                border: 1px solid {COLORS["line"]};
                border-radius: 14px;
            }}
            QFrame, QGroupBox {{
                background: {COLORS["card"]};
                border: 1px solid {COLORS["line"]};
                border-radius: 14px;
            }}
            QFrame#HeaderPanel {{
                background: {COLORS["card"]};
            }}
            QFrame#StatusPanel {{
                background: #D9EEF4;
            }}
            QGroupBox {{
                margin-top: 14px;
                padding-top: 16px;
                font-size: 11.5pt;
                font-weight: 600;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 4px;
                color: {COLORS["ink"]};
            }}
            QLabel#BrandLabel {{
                font-size: 10pt;
                font-weight: 500;
                letter-spacing: 0.5px;
                color: {COLORS["muted"]};
                border: none;
                background: transparent;
            }}
            QLabel#HeaderTitle {{
                font-size: 22pt;
                font-weight: 500;
                color: {COLORS["ink"]};
                border: none;
                background: transparent;
            }}
            QLabel#HeaderLogo {{
                border: none;
                background: transparent;
                padding: 0;
                margin: 0;
            }}
            QLabel#HeaderCopy {{
                font-size: 10.5pt;
                color: {COLORS["muted"]};
                border: none;
                background: transparent;
            }}
            QLabel#StageLabel {{
                font-size: 20pt;
                font-weight: 500;
                color: {COLORS["ink"]};
                border: none;
                background: transparent;
            }}
            QLabel#StageCopy {{
                font-size: 10pt;
                color: {COLORS["muted"]};
                border: none;
                background: transparent;
            }}
            QLabel#FieldLabel {{
                color: {COLORS["muted"]};
                font-size: 8.5pt;
                font-weight: 500;
                border: none;
                background: transparent;
            }}
            QLabel#MetricTitle {{
                color: {COLORS["muted"]};
                font-size: 8.5pt;
                font-weight: 500;
                border: none;
                background: transparent;
            }}
            QLabel#MetricValue {{
                font-size: 17pt;
                font-weight: 600;
                color: {COLORS["ink"]};
                border: none;
                background: transparent;
            }}
            QLabel#MetricDetail, QLabel#BodyCopy, QLabel#InlineValue {{
                color: {COLORS["ink"]};
                border: none;
                background: transparent;
            }}
            QLabel#InlineValue {{
                font-size: 12pt;
                font-weight: 600;
            }}
            QFrame#MetricCard {{
                background: {COLORS["card"]};
                border-radius: 12px;
                min-height: 104px;
                border: 1px solid {COLORS["line"]};
            }}
            QFrame#MetricCard[flashState="success"] {{
                background: #CFEED8;
                border: 1px solid #4FA36B;
            }}
            QFrame#FieldShell {{
                background: {COLORS["surface"]};
                border: 1px solid {COLORS["line"]};
                border-radius: 10px;
            }}
            QLineEdit, QPlainTextEdit {{
                background: {COLORS["surface"]};
                border: 1px solid {COLORS["line"]};
                border-radius: 10px;
                padding: 6px 9px;
                selection-background-color: {COLORS["accent"]};
                font-size: 10pt;
            }}
            QComboBox#FormCombo {{
                background: transparent;
                border: none;
                padding: 1px 2px 1px 8px;
                min-height: 30px;
                selection-background-color: {COLORS["accent"]};
            }}
            QComboBox#FormCombo::drop-down {{
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 0px;
                border: none;
                background: transparent;
            }}
            QComboBox#FormCombo::down-arrow {{
                image: none;
                width: 0px;
                height: 0px;
            }}
            QComboBox#FormCombo QAbstractItemView {{
                background: {COLORS["surface"]};
                color: {COLORS["ink"]};
                border: 1px solid {COLORS["line"]};
                selection-background-color: {COLORS["accent_soft"]};
                selection-color: {COLORS["ink"]};
                outline: none;
            }}
            QComboBox#FormCombo QLineEdit {{
                background: transparent;
                border: none;
                padding: 0;
                margin: 0;
            }}
            QToolButton#ComboArrowButton {{
                background: transparent;
                border: none;
                min-width: 24px;
                max-width: 24px;
                min-height: 30px;
                padding: 0 6px 0 0;
                color: {COLORS["muted"]};
                font-size: 12pt;
                font-weight: 600;
            }}
            QToolButton#ComboArrowButton:hover {{
                color: {COLORS["accent_dark"]};
            }}
            QPlainTextEdit#LogView {{
                font-family: "SF Mono", "Cascadia Code", monospace;
                background: #FCFDFE;
                font-size: 9.8pt;
                padding-top: 8px;
            }}
            QPushButton {{
                border-radius: 10px;
                min-height: 30px;
                padding: 3px 12px;
                font-weight: 600;
                font-size: 9.5pt;
            }}
            QPushButton#AccentButton {{
                background: {COLORS["signal"]};
                color: {COLORS["ink"]};
                border: 1px solid {COLORS["signal_dark"]};
            }}
            QPushButton#AccentButton:hover {{
                background: #FFF066;
            }}
            QPushButton#SoftButton {{
                background: transparent;
                color: {COLORS["accent_dark"]};
                border: 1px solid {COLORS["accent"]};
            }}
            QPushButton#SoftButton:hover {{
                background: {COLORS["accent_soft"]};
            }}
            QPushButton#DangerButton {{
                background: {COLORS["danger"]};
                color: #FFFFFF;
                border: 1px solid {COLORS["danger_dark"]};
            }}
            QPushButton#DangerButton:hover {{
                background: #E35D5D;
            }}
            QPushButton#UtilityButton {{
                background: transparent;
                color: {COLORS["muted"]};
                border: 1px solid {COLORS["line"]};
                border-radius: 9px;
                min-height: 28px;
                padding: 3px 12px;
                font-size: 9pt;
                font-weight: 500;
            }}
            QPushButton#UtilityButton:hover {{
                background: {COLORS["accent_soft"]};
                color: {COLORS["accent_dark"]};
                border: 1px solid {COLORS["accent"]};
            }}
            QPushButton#UtilityButton:disabled {{
                color: #9AA2AB;
                border: 1px solid #D7DCE2;
            }}
            QFrame#FiberLine {{
                background: {COLORS["muted"]};
                border: none;
                border-radius: 1px;
            }}
            QFrame#FiberLine[variant="accent"] {{
                background: {COLORS["accent"]};
            }}
            QFrame#FiberLine[variant="signal"] {{
                background: {COLORS["signal"]};
            }}
            """
        )

    def refresh_ports(self) -> None:
        manual_text = self._saved_manual_port or self.manual_port_combo.currentText().strip()
        system_ports = sorted(list_serial_ports())
        auto_ports = auto_detectable_serial_ports(system_ports)

        self.available_ports = system_ports.copy()
        self.verified_port = verified_serial_port(auto_ports) or ""
        self.detected_port = self.verified_port or preferred_serial_port(auto_ports) or ""
        if self.verified_port:
            self.detected_port_label.setText(f"{self.verified_port} (verifiziert)")
        elif self.detected_port:
            self.detected_port_label.setText(f"{self.detected_port} (Vorschlag)")
        else:
            self.detected_port_label.setText("Keine Waage gefunden")

        self.manual_port_combo.clear()
        self.manual_port_combo.addItems(self.available_ports)
        if manual_text:
            self.manual_port_combo.setCurrentText(manual_text)
        elif self.detected_port:
            self.manual_port_combo.setCurrentText(self.detected_port)

        self.auto_port_button.setVisible(True)
        self.manual_port_button.setVisible(True)
        self.manual_port_shell.setVisible(self.manual_port_override)

        if self.scale_source is None:
            self.active_source_value.setText(self._active_source_preview())

        if self._selected_source_mode() != SOURCE_MODE_SERIAL or self.scale_source is not None:
            return

        if self.verified_port and not self.manual_port_override:
            self.connection_note_label.setText("Waage automatisch erkannt und am Datenformat verifiziert.")
        elif self.detected_port and not self.manual_port_override:
            self.connection_note_label.setText(
                "Port automatisch vorgeschlagen. Wenn nötig kannst du ihn manuell überschreiben."
            )
        elif not self.detected_port and not self.manual_port_override:
            self.connection_note_label.setText("Keine Waage gefunden. Bitte Waage anschließen oder den Port manuell auswählen.")

    def toggle_manual_port_selection(self) -> None:
        self.manual_port_override = True
        self._saved_manual_port = self.manual_port_combo.currentText().strip()
        self.refresh_ports()

    def use_auto_port_selection(self) -> None:
        self.detected_port_label.setText("Port der Waage wird gesucht ...")
        self.connection_note_label.setText("Suche nach einer angeschlossenen Waage ...")
        self.auto_port_button.setEnabled(False)
        self.manual_port_button.setEnabled(False)
        self.manual_port_override = False
        QApplication.processEvents()
        QTimer.singleShot(1000, self._finish_auto_port_selection)

    def _finish_auto_port_selection(self) -> None:
        try:
            self.refresh_ports()
        finally:
            self.auto_port_button.setEnabled(True)
            self.manual_port_button.setEnabled(True)

    def browse_excel_file(self) -> None:
        start_path = self.excel_path_edit.text().strip() or str(Path.cwd())
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Excel-Datei auswählen",
            start_path,
            "Excel-Datei (*.xlsx)",
        )
        if not path:
            return

        blocked_reason = workbook_path_block_reason(Path(path))
        if blocked_reason:
            QMessageBox.warning(self, "Excel-Datei", blocked_reason)
            return

        self.excel_path_edit.setText(path)
        self._handle_excel_settings_changed()

    def open_excel_file(self) -> None:
        path_text = self.excel_path_edit.text().strip()
        if not path_text:
            QMessageBox.information(self, "Excel-Datei", "Bitte zuerst eine Excel-Datei auswählen.")
            return

        path = Path(path_text).expanduser()
        if not path.exists():
            QMessageBox.warning(self, "Excel-Datei", "Die ausgewählte Excel-Datei wurde nicht gefunden.")
            return

        if not QDesktopServices.openUrl(QUrl.fromLocalFile(str(path))):
            QMessageBox.warning(self, "Excel-Datei", "Die Excel-Datei konnte nicht geöffnet werden.")

    def _handle_primary_source_action(self) -> None:
        if self.source_control_state == SOURCE_CONTROL_IDLE:
            self.start_source()
            return

        if self.source_control_state == SOURCE_CONTROL_RUNNING:
            self.pause_source()
            return

        self.resume_source()

    def _set_source_control_state(self, state: str) -> None:
        self.source_control_state = state
        self._set_running_state(self.scale_source is not None)

    def _refresh_source_controls(self) -> None:
        state = self.source_control_state
        primary_label = {
            SOURCE_CONTROL_IDLE: "Quelle starten",
            SOURCE_CONTROL_RUNNING: "Logging pausieren",
            SOURCE_CONTROL_PAUSED: "Logging fortsetzen",
        }[state]

        self.primary_source_button.setText(primary_label)
        self.primary_source_button.setVisible(True)
        self.stop_button.setVisible(state != SOURCE_CONTROL_IDLE)

        start_available = state != SOURCE_CONTROL_IDLE or self.scale_source is None
        self.primary_source_button.setEnabled(start_available)
        self.stop_button.setEnabled(state != SOURCE_CONTROL_IDLE and self.scale_source is not None)

    def start_source(self) -> None:
        if self.scale_source and self.scale_source.is_alive():
            return

        if self._selected_source_mode() == SOURCE_MODE_SERIAL:
            self.refresh_ports()

        try:
            capture_settings = self._collect_capture_settings()
            self.capture_engine.update_settings(capture_settings)
            session = self._sync_excel_session()
            column, row = session.detect_current_cell()
            source = self._build_scale_source(capture_settings.target_weight)
        except Exception as exc:
            QMessageBox.critical(self, "Konfiguration", str(exc))
            return

        self.scale_source = source
        self.paused = False
        self.last_excel_error = None
        self._last_logged_value = None
        self._set_logged_feedback_active(False)
        self._set_source_control_state(SOURCE_CONTROL_RUNNING)
        self._set_pending_value("--")
        self.pending_detail_label.setText("Warte auf Treffer")
        self._set_next_cell_position(column, row)
        self._set_backend(session.backend_display_name())
        self.active_source_value.setText(self._source_runtime_name(source))
        self._set_running_state(True)
        self._set_stage("Verbinde Quelle", f"Starte {self._source_runtime_name(source)} und warte auf den Datenstrom.")
        self.connection_note_label.setText("Quelle wird gestartet...")
        source.start()
        self._save_settings()

    def connect_source(self) -> None:
        self.start_source()

    def pause_source(self) -> None:
        if self.scale_source is None:
            return

        self.paused = True
        self.capture_engine.reset()
        self._set_pending_value("--")
        self.pending_detail_label.setText("Logging aus")
        self._set_stage("Pausiert", "Die Quelle läuft weiter, aber es wird nichts in Excel geschrieben.")
        self.connection_note_label.setText("Pausiert. Live-Werte laufen weiter, Logging ist aus.")
        self._set_source_control_state(SOURCE_CONTROL_PAUSED)
        self._refresh_runtime_inputs()

    def resume_source(self) -> None:
        if self.scale_source is None:
            return

        self.paused = False
        self._apply_runtime_target_changes()
        self.pending_detail_label.setText("Warte auf Treffer")
        self._set_stage("Warte auf neues Bauteil", self._target_instruction_text())
        self.connection_note_label.setText("Logging wieder aktiv.")
        self._set_source_control_state(SOURCE_CONTROL_RUNNING)
        self._refresh_runtime_inputs()

    def toggle_pause_logging(self) -> None:
        if self.source_control_state == SOURCE_CONTROL_RUNNING:
            self.pause_source()
            return

        if self.source_control_state == SOURCE_CONTROL_PAUSED:
            self.resume_source()

    def stop_source(self) -> None:
        if self.scale_source is not None:
            self.paused = False
            self._set_source_control_state(SOURCE_CONTROL_IDLE)
            self._last_logged_value = None
            self._set_logged_feedback_active(False)
            self._set_pending_value("--")
            self.scale_source.stop()
            self.connection_note_label.setText("Quelle wird gestoppt...")

    def disconnect_source(self) -> None:
        self.stop_source()

    def refresh_excel_target(self) -> None:
        self._refresh_excel_target(silent=False)

    def _refresh_excel_target(self, silent: bool) -> None:
        try:
            session = self._sync_excel_session()
            column, row = session.detect_current_cell()
        except Exception as exc:
            self._set_backend("Auto")
            if silent:
                return
            QMessageBox.critical(self, "Excel-Ziel", str(exc))
            return

        self._set_next_cell_position(column, row)
        self._set_backend(session.backend_display_name())
        self._refresh_logging_format_display()
        self._save_settings()

    def _handle_excel_settings_changed(self) -> None:
        if self._syncing_excel_cursor:
            return

        self.excel_session = None
        if not self.excel_path_edit.text().strip():
            self._refresh_excel_file_ui()
            self._set_next_cell("--")
            self._set_backend("Auto")
            self._refresh_logging_format_display()
            self._save_settings()
            return

        self._refresh_excel_file_ui()
        self._refresh_excel_target(silent=True)
        self._save_settings()

    def _refresh_excel_file_ui(self) -> None:
        path_text = self.excel_path_edit.text().strip()
        has_file = bool(path_text)
        allow_file_change = self.source_control_state in {SOURCE_CONTROL_IDLE, SOURCE_CONTROL_PAUSED}

        if has_file:
            file_name = Path(path_text).name
            self.excel_file_name_label.setText(file_name)
            self.excel_file_name_label.setToolTip(path_text)
        else:
            self.excel_file_name_label.setText("")
            self.excel_file_name_label.setToolTip("")

        self.excel_file_name_label.setVisible(has_file)
        self.browse_excel_button.setVisible((not has_file) or allow_file_change)
        self.browse_excel_button.setText("Datei ändern" if has_file else "Datei auswählen")
        self.open_excel_button.setVisible(has_file)
        self.open_excel_button.setEnabled(has_file)

    def _apply_runtime_target_changes(self) -> None:
        try:
            capture_settings = self._collect_capture_settings()
        except ValueError:
            return

        self.capture_engine.update_settings(capture_settings)
        if self.scale_source is not None:
            self.capture_engine.reset()
            self._set_pending_value("--")
            if self.paused:
                self.pending_detail_label.setText("Ziel aktualisiert")
                self.connection_note_label.setText("Zielgewicht aktualisiert. Mit Logging fortsetzen startet die Erkennung neu.")
            else:
                self.pending_detail_label.setText("Warte auf Treffer")

        if isinstance(self.scale_source, SimulatedScaleSource):
            self.scale_source.update_target_weight(capture_settings.target_weight)
            if self.paused:
                self.scale_source.reset_cycle()

        self._save_settings()

    def _set_logged_feedback_active(self, active: bool) -> None:
        for widget in self._flash_widgets:
            widget.setProperty("flashState", "success" if active else "")
            widget.style().unpolish(widget)
            widget.style().polish(widget)
            widget.update()

    def _poll_source_events(self) -> None:
        if self.scale_source is None:
            return

        source = self.scale_source
        while True:
            try:
                event = source.events.get_nowait()
            except queue.Empty:
                break
            self._handle_source_event(event)
            if self.scale_source is not source:
                break

    def _handle_source_event(self, event: StreamEvent) -> None:
        if event.kind == "connected":
            self.connection_note_label.setText(event.message)
            self.active_source_value.setText(
                self._source_runtime_name(self.scale_source) if self.scale_source else "--"
            )
            self._set_stage("Warte auf neues Bauteil", self._target_instruction_text())
            self._log(f"Quelle verbunden: {event.message}")
            return

        if event.kind == "measurement" and event.measurement is not None:
            self._handle_measurement(event)
            return

        if event.kind == "raw":
            self.raw_value_label.setText(event.raw_text or "-")
            return

        if event.kind == "error":
            self._set_stage("Fehler", event.message)
            self.connection_note_label.setText(event.message)
            self._log(f"Fehlermeldung: {event.message}")
            if self._selected_source_mode() == SOURCE_MODE_SERIAL:
                QMessageBox.critical(self, "Serielle Verbindung", event.message)
            return

        if event.kind == "stopped":
            was_simulation = isinstance(self.scale_source, SimulatedScaleSource)
            self.scale_source = None
            self.paused = False
            self._set_source_control_state(SOURCE_CONTROL_IDLE)
            self._set_running_state(False)
            self.active_source_value.setText(self._active_source_preview())
            self.connection_note_label.setText("Simulation gestoppt." if was_simulation else "Quelle gestoppt.")
            self._set_stage(
                "Bereit zum Start",
                "Die Einstellungen sind wieder sichtbar. Für die nächste Serie einfach erneut anschalten.",
            )
            self._log("Quelle wurde gestoppt.")
            self.refresh_ports()

    def _handle_measurement(self, event: StreamEvent) -> None:
        measurement = event.measurement
        if measurement is None:
            return

        if not self._ensure_gram_unit(measurement.unit):
            return

        self._unit_error = None
        self._set_live_weight(f"{measurement.value:.3f}", measurement.unit or "g")
        self._clear_logged_feedback_if_value_changed(measurement.value)

        if self.paused:
            return

        state = self.capture_engine.process(measurement)
        self._update_capture_dashboard(state)

        if state.rearmed:
            self._log("Bauteil entfernt. Nächste Wägung ist bereit.")

        if state.new_candidate is not None:
            self._log(f"Stabiler Messwert erkannt: {state.new_candidate:.3f} g")

        if self.capture_engine.peek_pending_capture() is not None:
            self._write_pending_capture(auto=True)

    def _update_capture_dashboard(self, state: CaptureState) -> None:
        if state.pending_capture is not None:
            self._set_pending_value(f"{state.pending_capture:.3f}")

        if state.pending_capture is not None:
            self.pending_detail_label.setText("Wird in Excel geschrieben")
            self._set_stage("Stabil erkannt", "Messwert wird gerade übernommen.")
            return

        if not state.armed:
            self.pending_detail_label.setText("Warte auf Entnahme")
            self._set_stage("Bauteil entfernen", "Sobald die Waage wieder frei ist, startet die nächste Messung.")
            return

        if state.within_target and not state.stable:
            if state.spread is None:
                self.pending_detail_label.setText("Sammle Werte")
            else:
                self.pending_detail_label.setText(f"Spread {state.spread:.3f} g")
            self._set_stage(
                "Gewicht stabilisiert sich",
                f"Im Zielbereich. Automatik wartet auf {self.capture_engine.settings.stable_samples} ruhige Werte.",
            )
            return

        if not state.within_target:
            self.pending_detail_label.setText("Außerhalb Zielbereich")
        else:
            self.pending_detail_label.setText("Warte auf Treffer")
        self._set_stage("Warte auf neues Bauteil", self._target_instruction_text())

    def _write_pending_capture(self, auto: bool) -> None:
        pending_value = self.capture_engine.peek_pending_capture()
        if pending_value is None:
            return

        try:
            session = self._sync_excel_session()
            result = session.write_value(pending_value)
            next_column, next_row = session.detect_current_cell()
        except Exception as exc:
            message = str(exc)
            self._set_stage("Excel-Fehler", message)
            self.connection_note_label.setText("Excel-Ziel bitte prüfen.")
            if self.last_excel_error != message:
                self._log(f"Excel-Fehler: {message}")
            self.last_excel_error = message
            if not auto:
                QMessageBox.critical(self, "Excel-Fehler", message)
            return

        self.last_excel_error = None
        committed_value = self.capture_engine.commit_pending_capture()
        self._set_pending_value(f"{pending_value:.3f}")
        self.pending_detail_label.setText("Gespeichert")
        self._set_next_cell_position(next_column, next_row)
        self._set_backend(session.backend_display_name())
        self._set_stage("Messwert gespeichert", f"{result.sheet_name}!{result.cell} wurde beschrieben. Bitte Bauteil entfernen.")
        self.connection_note_label.setText(f"Gespeichert in {result.sheet_name}!{result.cell}")
        self._set_logged_feedback_active(True)
        if committed_value is not None:
            self._last_logged_value = committed_value
            self._log(f"Messwert {committed_value:.3f} g wurde in {result.sheet_name}!{result.cell} geschrieben.")
        self._save_settings()

    def _build_scale_source(self, target_weight: float | None) -> ScaleSource:
        if self._selected_source_mode() == SOURCE_MODE_SIMULATION:
            profile = self.simulation_profile_combo.currentData() or SIM_PROFILE_STABLE
            return SimulatedScaleSource(profile=profile, target_weight=target_weight)

        return SerialScaleSource(self._collect_serial_settings())

    def _collect_serial_settings(self) -> SerialSettings:
        port = self._selected_port()
        if not port:
            raise ValueError("Keine Waage gefunden. Bitte Waage anschließen oder den Port manuell auswählen.")

        return SerialSettings(port=port, baudrate=9600, timeout=1.0)

    def _collect_capture_settings(self):
        target_weight = self._parse_float(self.target_weight_edit.text(), "Zielgewicht")
        target_window = self._parse_float(self.target_window_edit.text(), "Abweichung")

        if target_weight <= 0:
            raise ValueError("Das Zielgewicht muss größer als 0 sein.")
        if target_window <= 0:
            raise ValueError("Die Abweichung muss größer als 0 sein.")

        return build_capture_settings(target_weight, target_window)

    def _collect_excel_settings(self) -> ExcelSettings:
        return ExcelSettings(
            path=self.excel_path_edit.text().strip(),
            sheet_name=self.sheet_name_edit.text().strip() or "Messwerte",
            column=self.column_edit.text().strip().upper() or "A",
            start_row=self._parse_int(self.start_row_edit.text(), "Zeile"),
            direction=self.direction_combo.currentData() or FLOW_DOWN,
            mode=EXCEL_MODE_AUTO,
        )

    def _sync_excel_session(self) -> ExcelSession:
        settings = self._collect_excel_settings()
        if self.excel_session is None:
            self.excel_session = ExcelSession(settings)
        else:
            self.excel_session.update_settings(settings)
        return self.excel_session

    def _refresh_auto_capture_hint(self) -> None:
        try:
            target_weight = self._parse_float(self.target_weight_edit.text(), "Zielgewicht")
            target_window = self._parse_float(self.target_window_edit.text(), "Abweichung")
            settings = build_capture_settings(target_weight, target_window)
            self.auto_capture_hint.setText(
                f"{settings.stable_samples} Wiederholungen, Start-Toleranz {settings.base_stability_tolerance:.3f} g, danach adaptiv."
            )
        except ValueError:
            self.auto_capture_hint.setText("Start-Toleranz wird intern gesetzt.")

    def _refresh_target_range_display(self) -> None:
        try:
            target_weight = self._parse_float(self.target_weight_edit.text(), "Zielgewicht")
            target_window = self._parse_float(self.target_window_edit.text(), "Abweichung")
        except ValueError:
            self.target_range_value.setText("--")
            return

        lower = target_weight - target_window
        upper = target_weight + target_window
        self.target_range_value.setText(f"{lower:.2f} bis {upper:.2f} g")

    def _update_source_mode_ui(self) -> None:
        mode = self._selected_source_mode()
        is_simulation = mode == SOURCE_MODE_SIMULATION
        self.serial_config_panel.setVisible(not is_simulation)
        self.simulation_config_panel.setVisible(is_simulation)

        if not is_simulation and self.scale_source is None:
            self.refresh_ports()

        if self.scale_source is None:
            self.connection_note_label.setText(self._idle_connection_text())
            self.active_source_value.setText(self._active_source_preview())

        if is_simulation:
            profile = simulation_profile_label(self.simulation_profile_combo.currentData() or SIM_PROFILE_STABLE)
            self.simulation_hint_label.setText(
                f"Virtuelle Waage aktiv. Profil: {profile}. Ideal für Live-Wert, Stabilität und Excel-Tests."
            )
        self._refresh_logging_format_display()
        self._refresh_runtime_inputs()

    def _selected_source_mode(self) -> str:
        return self.source_mode_combo.currentData() or SOURCE_MODE_SERIAL

    def _selected_port(self) -> str:
        if self.manual_port_override:
            return self.manual_port_combo.currentText().strip()
        return self.detected_port or self.manual_port_combo.currentText().strip()

    def _active_source_preview(self) -> str:
        if self._selected_source_mode() == SOURCE_MODE_SIMULATION:
            profile = simulation_profile_label(self.simulation_profile_combo.currentData() or SIM_PROFILE_STABLE)
            return f"SIMULATED_SCALE | {profile}"
        return self._selected_port() or "--"

    def _source_runtime_name(self, source: ScaleSource) -> str:
        if isinstance(source, SimulatedScaleSource):
            return f"Simulation ({simulation_profile_label(source.profile)})"
        return source.current_port_name

    def _idle_connection_text(self) -> str:
        if self._selected_source_mode() == SOURCE_MODE_SIMULATION:
            profile = simulation_profile_label(self.simulation_profile_combo.currentData() or SIM_PROFILE_STABLE)
            return f"Simulation bereit: {profile}."
        if self.detected_port:
            return f"Waage bereit auf {self.detected_port}."
        return "Keine Waage gefunden. Bitte Waage anschließen oder den Port manuell auswählen."

    def _current_target_weight_or_none(self) -> float | None:
        text = self.target_weight_edit.text().strip()
        if not text:
            return None
        try:
            return self._parse_float(text, "Zielgewicht")
        except ValueError:
            return None

    def _target_instruction_text(self) -> str:
        bounds = self.capture_engine.window_bounds()
        if bounds is None:
            return "Lege ein neues Bauteil auf."

        if self._selected_source_mode() == SOURCE_MODE_SIMULATION:
            return (
                f"Simulation arbeitet im Bereich {bounds[0]:.2f} bis {bounds[1]:.2f} g. "
                "Live-Wert und Stabilität können jetzt getestet werden."
            )
        return f"Lege ein Bauteil im Bereich {bounds[0]:.2f} bis {bounds[1]:.2f} g auf."

    def _set_running_state(self, running: bool) -> None:
        show_setup_panel = self.source_control_state in {SOURCE_CONTROL_IDLE, SOURCE_CONTROL_PAUSED}
        self.connection_setup_panel.setVisible(self.source_control_state == SOURCE_CONTROL_IDLE)
        self.active_source_panel.setVisible(self.source_control_state != SOURCE_CONTROL_IDLE)
        self.setup_panel.setVisible(show_setup_panel)
        self.left_scroll.setMaximumWidth(520 if show_setup_panel else 470)
        self._refresh_source_controls()
        self._refresh_runtime_inputs()

    def _refresh_runtime_inputs(self) -> None:
        allow_edit = self.source_control_state in {SOURCE_CONTROL_IDLE, SOURCE_CONTROL_PAUSED}

        for widget in (
            self.sheet_name_edit,
        ):
            widget.setEnabled(allow_edit)

        self._refresh_excel_file_ui()
        self.browse_excel_button.setEnabled(allow_edit)
        self.column_edit.setEnabled(allow_edit)
        self.start_row_edit.setEnabled(allow_edit)
        self.direction_combo.setEnabled(allow_edit)
        self.target_weight_edit.setEnabled(allow_edit)
        self.target_window_edit.setEnabled(allow_edit)

    def _parse_float(self, value: str, label: str) -> float:
        try:
            return float(value.strip().replace(",", "."))
        except ValueError as exc:
            raise ValueError(f"{label} ist keine gültige Zahl.") from exc

    def _parse_int(self, value: str, label: str) -> int:
        try:
            number = int(value.strip())
        except ValueError as exc:
            raise ValueError(f"{label} ist keine gültige ganze Zahl.") from exc

        if number < 1:
            raise ValueError(f"{label} muss größer als 0 sein.")
        return number

    def _set_stage(self, title: str, detail: str) -> None:
        self.stage_label.setText(title)
        self.stage_copy_label.setText(detail)

    def _set_live_weight(self, value_text: str, unit_text: str) -> None:
        self.live_weight_value.setText(value_text)
        self.live_weight_detail.setText(unit_text.upper())

    def _set_pending_value(self, text: str) -> None:
        self.pending_value_label.setText(text)

    def _set_next_cell(self, text: str) -> None:
        self.next_cell_value.setText(text)
        self.next_cell_detail_label.setText("Nächster Excel-Eintrag")

    def _set_next_cell_position(self, column: str, row: int) -> None:
        self._set_next_cell(build_cell_ref(column, row))
        self._sync_excel_cursor_inputs(column, row)

    def _sync_excel_cursor_inputs(self, column: str, row: int) -> None:
        self._syncing_excel_cursor = True
        try:
            self.column_edit.blockSignals(True)
            self.start_row_edit.blockSignals(True)
            self.column_edit.setText(column)
            self.start_row_edit.setText(str(row))
        finally:
            self.column_edit.blockSignals(False)
            self.start_row_edit.blockSignals(False)
            self._syncing_excel_cursor = False

    def _set_backend(self, text: str) -> None:
        self.backend_value.setText(text)

    def _refresh_logging_format_display(self) -> None:
        self.logging_format_value.setText(self.direction_combo.currentText() or "--")

    def _clear_logged_feedback_if_value_changed(self, measured_value: float) -> None:
        if self._last_logged_value is None:
            return

        tolerance = max(
            self.capture_engine.settings.base_stability_tolerance,
            self.capture_engine.effective_tolerance(),
        )
        if abs(measured_value - self._last_logged_value) <= tolerance:
            return

        self._last_logged_value = None
        self._set_logged_feedback_active(False)
        self._set_pending_value("--")

    def _ensure_gram_unit(self, unit_text: str) -> bool:
        normalized = unit_text.strip().upper()
        if normalized in {"G", "GR", "GM"}:
            return True

        unit_display = normalized or "ohne Einheit"

        if self._unit_error == unit_display:
            return False

        self._unit_error = unit_display
        source_unit_text = f"in {unit_display}" if normalized else unit_display
        message = (
            f"Die Waage sendet aktuell {source_unit_text}. "
            "Bitte die Einheit an der Waage auf Gramm (g) umstellen."
        )
        self.paused = True
        if self.scale_source is not None:
            self._set_source_control_state(SOURCE_CONTROL_PAUSED)
            self._set_running_state(True)
        self.pending_detail_label.setText("Einheit prüfen")
        self._set_stage("Einheit prüfen", message)
        self.connection_note_label.setText(message)
        self._log(f"Einheiten-Fehler: {message}")
        QMessageBox.warning(self, "Einheit prüfen", message)
        return False

    def _log(self, message: str) -> None:
        self.log_view.appendPlainText(f"• {message}")
        self.clear_log_button.setEnabled(True)

    def _clear_log_history(self) -> None:
        self.log_view.clear()
        self.clear_log_button.setEnabled(False)

    def _load_settings(self) -> None:
        data = self.settings_store.load()
        if not data:
            direction_index = self.direction_combo.findData(FLOW_DOWN)
            if direction_index >= 0:
                self.direction_combo.setCurrentIndex(direction_index)
            return

        source_mode = str(data.get("source_mode", SOURCE_MODE_SERIAL))
        source_mode_index = self.source_mode_combo.findData(source_mode)
        if source_mode_index >= 0:
            self.source_mode_combo.setCurrentIndex(source_mode_index)

        sim_profile = str(data.get("simulation_profile", SIM_PROFILE_STABLE))
        sim_profile_index = self.simulation_profile_combo.findData(sim_profile)
        if sim_profile_index >= 0:
            self.simulation_profile_combo.setCurrentIndex(sim_profile_index)

        self.manual_port_override = bool(data.get("manual_port_override", False))
        self._saved_manual_port = str(data.get("manual_port", data.get("port", "")))

        self.target_weight_edit.setText(str(data.get("target_weight", "12.50")))
        self.target_window_edit.setText(str(data.get("target_window", "0.50")))

        self.excel_path_edit.setText(str(data.get("excel_path", "")))
        self.sheet_name_edit.setText(str(data.get("sheet_name", "Messwerte")))
        self.column_edit.setText(str(data.get("column", "A")))
        self.start_row_edit.setText(str(data.get("start_row", "2")))

        direction = str(data.get("direction", FLOW_DOWN))
        direction_index = self.direction_combo.findData(direction)
        if direction_index >= 0:
            self.direction_combo.setCurrentIndex(direction_index)

    def _save_settings(self) -> None:
        self.settings_store.save(
            {
                "source_mode": self._selected_source_mode(),
                "simulation_profile": self.simulation_profile_combo.currentData() or SIM_PROFILE_STABLE,
                "manual_port_override": self.manual_port_override,
                "manual_port": self.manual_port_combo.currentText().strip(),
                "target_weight": self.target_weight_edit.text().strip(),
                "target_window": self.target_window_edit.text().strip(),
                "excel_path": self.excel_path_edit.text().strip(),
                "sheet_name": self.sheet_name_edit.text().strip(),
                "column": self.column_edit.text().strip(),
                "start_row": self.start_row_edit.text().strip(),
                "direction": self.direction_combo.currentData() or FLOW_DOWN,
            }
        )

    def closeEvent(self, event: QCloseEvent) -> None:
        self._save_settings()
        if self.scale_source is not None:
            self.scale_source.stop()
        super().closeEvent(event)


def main() -> None:
    app = QApplication.instance() or QApplication(sys.argv)
    window = ScaleLoggerWindow()
    window.show()
    app.exec()


if __name__ == "__main__":
    main()
