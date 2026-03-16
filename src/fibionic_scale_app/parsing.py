from __future__ import annotations

import re
from typing import Final

from .models import Measurement

MEASUREMENT_PATTERN: Final[re.Pattern[str]] = re.compile(
    r"(?P<sign>[+-])?\s*(?P<number>\d[\d\s]*(?:[.,]\d*)?)\s*(?P<unit>[A-Za-z]{0,3})"
)


def clean_raw_text(raw: str | bytes) -> str:
    if isinstance(raw, bytes):
        text = raw.decode("ascii", errors="ignore")
    else:
        text = str(raw)

    return text.replace("\r", "").replace("\n", "").replace("\x00", "").strip()


def parse_scale_output(raw: str | bytes) -> Measurement | None:
    text = clean_raw_text(raw)
    if not text:
        return None

    match = MEASUREMENT_PATTERN.search(text)
    if not match:
        return None

    sign = -1 if match.group("sign") == "-" else 1
    number_text = match.group("number").replace(" ", "").replace(",", ".")
    if number_text.endswith("."):
        number_text += "0"

    try:
        value = sign * float(number_text)
    except ValueError:
        return None

    return Measurement(
        value=value,
        raw_text=text,
        unit=match.group("unit").strip().upper(),
    )
