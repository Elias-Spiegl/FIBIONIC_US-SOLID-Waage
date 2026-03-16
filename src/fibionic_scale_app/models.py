from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime


@dataclass(slots=True)
class Measurement:
    value: float
    raw_text: str
    unit: str = ""
    timestamp: datetime = field(default_factory=datetime.now)


@dataclass(slots=True)
class SerialSettings:
    port: str = ""
    baudrate: int = 9600
    timeout: float = 1.0


@dataclass(slots=True)
class CaptureSettings:
    target_weight: float | None = None
    target_window: float = 0.50
    stability_tolerance: float = 0.05
    stable_samples: int = 6
    rearm_threshold: float = 0.10
    minimum_weight: float = 0.05
    require_confirmation: bool = False


@dataclass(slots=True)
class ExcelSettings:
    path: str = ""
    sheet_name: str = "Messwerte"
    column: str = "A"
    start_row: int = 2
    auto_advance: bool = True
    mode: str = "auto"
