from __future__ import annotations

from dataclasses import replace
from decimal import ROUND_HALF_UP, Decimal

from .models import Measurement

WEIGHT_DECIMALS = 2
WEIGHT_NUMBER_FORMAT = "0.##"
_WEIGHT_QUANTUM = Decimal("0.01")


def quantize_weight_value(value: float) -> float:
    return float(Decimal(str(value)).quantize(_WEIGHT_QUANTUM, rounding=ROUND_HALF_UP))


def normalize_measurement(measurement: Measurement) -> Measurement:
    return replace(measurement, value=quantize_weight_value(measurement.value))


def format_weight_value(value: float) -> str:
    quantized = quantize_weight_value(value)
    if quantized == 0:
        return "0"

    return f"{quantized:.{WEIGHT_DECIMALS}f}".rstrip("0").rstrip(".")
