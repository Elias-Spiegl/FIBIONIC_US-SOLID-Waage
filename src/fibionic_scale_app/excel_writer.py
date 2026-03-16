from __future__ import annotations

import importlib.util
import os
import sys
from dataclasses import dataclass
from pathlib import Path

from .models import ExcelSettings

EXCEL_MODE_AUTO = "auto"
EXCEL_MODE_FILE = "file"
EXCEL_MODE_LIVE = "live"


@dataclass(slots=True)
class ExcelWriteResult:
    path: Path
    sheet_name: str
    cell: str
    value: float
    row: int
    column: str
    backend: str


class LiveExcelUnavailableError(RuntimeError):
    """Raised when the live Excel backend cannot be used."""


def normalize_column_name(column: str) -> str:
    normalized = column.strip().upper()
    if not normalized or not normalized.isalpha():
        raise ValueError("Die Excel-Spalte muss aus Buchstaben bestehen, z. B. A oder AB.")
    return normalized


def normalize_excel_mode(mode: str) -> str:
    normalized = (mode or EXCEL_MODE_AUTO).strip().lower()
    if normalized not in {EXCEL_MODE_AUTO, EXCEL_MODE_FILE, EXCEL_MODE_LIVE}:
        raise ValueError("Der Excel-Modus muss 'auto', 'file' oder 'live' sein.")
    return normalized


def current_platform_label() -> str:
    if sys.platform == "darwin":
        return "macOS"
    if sys.platform.startswith("win"):
        return "Windows"
    return sys.platform


def excel_mode_options() -> list[tuple[str, str]]:
    platform_name = current_platform_label()
    return [
        (EXCEL_MODE_AUTO, f"Auto ({platform_name})"),
        (EXCEL_MODE_FILE, "Datei-Modus (.xlsx)"),
        (EXCEL_MODE_LIVE, "Live-Modus (offenes Excel)"),
    ]


def mode_label(mode: str) -> str:
    normalized = normalize_excel_mode(mode)
    if normalized == EXCEL_MODE_AUTO:
        return "Auto"
    if normalized == EXCEL_MODE_FILE:
        return "Datei"
    return "Live"


def backend_label(backend: str) -> str:
    normalized = normalize_excel_mode(backend)
    if normalized == EXCEL_MODE_FILE:
        return "Datei-Writer (openpyxl)"
    if normalized == EXCEL_MODE_LIVE:
        return "Live-Writer (Excel/xlwings)"
    return "Auto"


def live_backend_supported() -> bool:
    if sys.platform not in {"darwin", "win32"}:
        return False

    return importlib.util.find_spec("xlwings") is not None


def live_backend_status_text() -> str:
    platform_name = current_platform_label()
    if sys.platform not in {"darwin", "win32"}:
        return f"{platform_name}: Live-Modus wird nur auf macOS und Windows unterstuetzt."
    if live_backend_supported():
        return f"{platform_name}: Live-Modus ueber die lokal installierte Excel-App verfuegbar."
    return f"{platform_name}: Fuer den Live-Modus fehlt derzeit das Paket 'xlwings'."


def normalize_workbook_path(path_text: str) -> Path:
    if not path_text.strip():
        raise ValueError("Bitte zuerst eine Excel-Datei auswählen.")

    path = Path(path_text).expanduser()
    if not path.is_absolute():
        path = Path.cwd() / path

    if path.suffix.lower() != ".xlsx":
        raise ValueError("Bitte eine .xlsx-Datei verwenden.")

    return path


def find_next_empty_row_with_getter(value_getter, column: str, start_row: int) -> int:
    row = max(1, start_row)
    while True:
        if value_getter(f"{column}{row}") in (None, ""):
            return row
        row += 1


class FileExcelBackend:
    key = EXCEL_MODE_FILE

    def detect_current_row(self, settings: ExcelSettings) -> int:
        workbook, worksheet, _ = self._open_workbook(settings)
        column = normalize_column_name(settings.column)
        row = find_next_empty_row_with_getter(lambda cell: worksheet[cell].value, column, settings.start_row)
        workbook.close()
        return row

    def write_value(self, settings: ExcelSettings, row: int, value: float) -> ExcelWriteResult:
        workbook, worksheet, path = self._open_workbook(settings)
        column = normalize_column_name(settings.column)
        cell = f"{column}{row}"

        worksheet[cell] = float(value)
        workbook.save(path)
        workbook.close()

        return ExcelWriteResult(
            path=path,
            sheet_name=worksheet.title,
            cell=cell,
            value=float(value),
            row=row,
            column=column,
            backend=self.key,
        )

    def _open_workbook(self, settings: ExcelSettings):
        path = normalize_workbook_path(settings.path)

        try:
            from openpyxl import Workbook, load_workbook
        except ImportError as exc:
            raise RuntimeError(
                "Das Paket 'openpyxl' fehlt. Bitte installiere zuerst die Projekt-Abhängigkeiten."
            ) from exc

        path.parent.mkdir(parents=True, exist_ok=True)
        existing = path.exists()

        if existing:
            workbook = load_workbook(path)
        else:
            workbook = Workbook()

        sheet_name = settings.sheet_name.strip() or "Messwerte"
        if sheet_name in workbook.sheetnames:
            worksheet = workbook[sheet_name]
        elif len(workbook.sheetnames) == 1 and workbook.active.title == "Sheet" and not existing:
            worksheet = workbook.active
            worksheet.title = sheet_name
        else:
            worksheet = workbook.create_sheet(sheet_name)

        return workbook, worksheet, path


class LiveExcelBackend:
    key = EXCEL_MODE_LIVE

    def detect_current_row(self, settings: ExcelSettings) -> int:
        _, worksheet, _ = self._open_workbook(settings)
        column = normalize_column_name(settings.column)
        return find_next_empty_row_with_getter(lambda cell: worksheet.range(cell).value, column, settings.start_row)

    def write_value(self, settings: ExcelSettings, row: int, value: float) -> ExcelWriteResult:
        workbook, worksheet, path = self._open_workbook(settings)
        column = normalize_column_name(settings.column)
        cell = f"{column}{row}"

        worksheet.range(cell).value = float(value)
        workbook.save()

        return ExcelWriteResult(
            path=path,
            sheet_name=worksheet.name,
            cell=cell,
            value=float(value),
            row=row,
            column=column,
            backend=self.key,
        )

    def _open_workbook(self, settings: ExcelSettings):
        path = normalize_workbook_path(settings.path)
        self._configure_macos_onedrive_env(path)
        xw = self._import_xlwings()

        try:
            workbook = self._resolve_or_open_workbook(xw, path)
        except Exception as exc:  # pragma: no cover - depends on local Excel app
            raise LiveExcelUnavailableError(
                "Live-Modus konnte die lokale Excel-App nicht verbinden. "
                "Bitte pruefe, ob Microsoft Excel lokal installiert ist, die Datei im Desktop-Excel geoeffnet werden kann "
                f"und der Python-Host Excel steuern darf. Originalfehler: {exc}"
            ) from exc

        try:  # pragma: no cover - depends on local Excel app
            workbook.app.visible = True
        except Exception:
            pass

        sheet_name = settings.sheet_name.strip() or "Messwerte"
        try:
            worksheet = workbook.sheets[sheet_name]
        except Exception:
            if len(workbook.sheets) == 1 and str(workbook.sheets[0].name).lower().startswith("sheet"):
                worksheet = workbook.sheets[0]
                worksheet.name = sheet_name
            else:
                worksheet = workbook.sheets.add(sheet_name, after=workbook.sheets[-1])

        return workbook, worksheet, path

    @staticmethod
    def _configure_macos_onedrive_env(path: Path) -> None:
        if sys.platform != "darwin":
            return

        resolved = path.resolve()
        original_chain = (path, *path.parents)
        resolved_chain = (resolved, *resolved.parents)
        for ancestor in (*original_chain, *resolved_chain):
            name = ancestor.name
            if name == "OneDrive":
                consumer_root = str(ancestor.expanduser())
                os.environ.setdefault("ONEDRIVE_CONSUMER_MAC", consumer_root)
                os.environ.setdefault("OneDriveConsumer", consumer_root)
                os.environ.setdefault("OneDrive", consumer_root)
                return
            if name.startswith("OneDrive - ") or name.startswith("OneDrive-"):
                commercial_root = LiveExcelBackend._preferred_macos_onedrive_root(ancestor)
                os.environ.setdefault("ONEDRIVE_COMMERCIAL_MAC", commercial_root)
                os.environ.setdefault("OneDriveCommercial", commercial_root)
                os.environ.setdefault("OneDrive", commercial_root)
                return

    @staticmethod
    def _preferred_macos_onedrive_root(root: Path) -> str:
        expanded = root.expanduser()
        resolved_root = expanded.resolve()
        home = Path.home()

        # xlwings documents the Finder-style root on macOS, e.g. ~/OneDrive - Company.
        # If the selected file lives under Library/CloudStorage, prefer a matching home alias/symlink.
        for candidate in home.glob("OneDrive*"):
            try:
                if candidate.resolve() == resolved_root:
                    return str(candidate)
            except OSError:
                continue

        return str(expanded)

    def _resolve_or_open_workbook(self, xw, path: Path):
        resolved_target = path.resolve()

        for app in self._iter_apps(xw):
            for book in app.books:
                if self._book_matches_path(book, resolved_target):
                    return book

        if path.exists():
            app = self._pick_or_create_app(xw)
            return app.books.open(str(path))

        path.parent.mkdir(parents=True, exist_ok=True)
        app = self._pick_or_create_app(xw)
        workbook = app.books.add()
        workbook.save(str(path))
        return workbook

    @staticmethod
    def _iter_apps(xw):
        try:
            return list(xw.apps)
        except Exception:  # pragma: no cover - depends on local Excel app
            return []

    @staticmethod
    def _pick_or_create_app(xw):
        apps = LiveExcelBackend._iter_apps(xw)
        if apps:
            return apps[0]
        return xw.App(visible=True, add_book=False)

    @staticmethod
    def _book_matches_path(book, target: Path) -> bool:
        fullname = getattr(book, "fullname", "") or ""
        if not fullname:
            return False

        try:
            candidate = Path(fullname).expanduser().resolve()
        except OSError:
            return False

        return candidate == target

    @staticmethod
    def _import_xlwings():
        if sys.platform not in {"darwin", "win32"}:
            raise LiveExcelUnavailableError("Live-Modus wird nur auf macOS und Windows unterstützt.")

        try:
            import xlwings as xw
        except ImportError as exc:
            raise LiveExcelUnavailableError(
                "Das Paket 'xlwings' fehlt. Bitte installiere zuerst die Projekt-Abhängigkeiten."
            ) from exc

        return xw


FILE_BACKEND = FileExcelBackend()
LIVE_BACKEND = LiveExcelBackend()


class ExcelSession:
    def __init__(self, settings: ExcelSettings):
        self.settings = settings
        self.current_row: int | None = None
        self.active_backend = normalize_excel_mode(settings.mode)

    def update_settings(self, settings: ExcelSettings, preserve_row: bool = False) -> None:
        self.settings = settings
        self.active_backend = normalize_excel_mode(settings.mode)
        if not preserve_row:
            self.current_row = None

    def reset_row(self, row: int | None = None) -> None:
        self.current_row = max(1, row or self.settings.start_row)

    def preview_cell(self) -> str:
        column = normalize_column_name(self.settings.column)
        row = self.current_row if self.current_row is not None else max(1, self.settings.start_row)
        return f"{column}{row}"

    def backend_display_name(self) -> str:
        return backend_label(self.active_backend)

    def detect_current_row(self) -> int:
        row = self._run_backend("detect_current_row")
        self.current_row = row
        return row

    def write_value(self, value: float) -> ExcelWriteResult:
        if self.current_row is None:
            self.current_row = self.detect_current_row()

        result = self._run_backend("write_value", self.current_row, value)
        written_row = self.current_row
        if self.settings.auto_advance:
            self.current_row += 1
        else:
            self.current_row = written_row

        return result

    def _run_backend(self, method_name: str, *args):
        mode = normalize_excel_mode(self.settings.mode)
        backends = [FILE_BACKEND] if mode == EXCEL_MODE_FILE else [LIVE_BACKEND] if mode == EXCEL_MODE_LIVE else [LIVE_BACKEND, FILE_BACKEND]
        last_error: Exception | None = None

        for backend in backends:
            try:
                result = getattr(backend, method_name)(self.settings, *args)
                self.active_backend = backend.key
                return result
            except LiveExcelUnavailableError as exc:
                last_error = exc
                if mode == EXCEL_MODE_AUTO:
                    continue
                raise

        if last_error is not None:
            raise last_error

        raise RuntimeError("Kein Excel-Backend verfügbar.")
