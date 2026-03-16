from __future__ import annotations

from collections import deque
from dataclasses import dataclass
from statistics import mean

from .models import CaptureSettings, Measurement


@dataclass(slots=True)
class CaptureState:
    measurement: Measurement
    stable: bool
    within_target: bool
    armed: bool
    pending_capture: float | None
    new_candidate: float | None
    spread: float | None
    rearmed: bool = False


class WeightCaptureEngine:
    def __init__(self, settings: CaptureSettings):
        self.settings = settings
        self.history: deque[float] = deque(maxlen=max(2, settings.stable_samples))
        self.armed = True
        self.pending_capture: float | None = None

    def update_settings(self, settings: CaptureSettings) -> None:
        self.settings = settings
        self.history = deque(maxlen=max(2, settings.stable_samples))
        self.armed = True
        self.pending_capture = None

    def reset(self) -> None:
        self.history.clear()
        self.armed = True
        self.pending_capture = None

    def peek_pending_capture(self) -> float | None:
        return self.pending_capture

    def commit_pending_capture(self) -> float | None:
        value = self.pending_capture
        self.pending_capture = None
        return value

    def discard_pending_capture(self) -> None:
        self.pending_capture = None

    def window_bounds(self) -> tuple[float, float] | None:
        if self.settings.target_weight is None:
            return None

        return (
            self.settings.target_weight - self.settings.target_window,
            self.settings.target_weight + self.settings.target_window,
        )

    def process(self, measurement: Measurement) -> CaptureState:
        rearmed = False
        if (
            self.pending_capture is None
            and not self.armed
            and abs(measurement.value) <= self.settings.rearm_threshold
        ):
            self.armed = True
            self.history.clear()
            rearmed = True

        self.history.append(measurement.value)
        within_target = self._matches_target(measurement.value)
        stable = False
        spread: float | None = None
        new_candidate: float | None = None

        if len(self.history) == self.history.maxlen:
            spread = max(self.history) - min(self.history)
            stable = spread <= self.settings.stability_tolerance

        if (
            self.armed
            and self.pending_capture is None
            and stable
            and within_target
        ):
            new_candidate = round(mean(self.history), 3)
            self.pending_capture = new_candidate
            self.armed = False

        return CaptureState(
            measurement=measurement,
            stable=stable,
            within_target=within_target,
            armed=self.armed,
            pending_capture=self.pending_capture,
            new_candidate=new_candidate,
            spread=spread,
            rearmed=rearmed,
        )

    def _matches_target(self, value: float) -> bool:
        if abs(value) < self.settings.minimum_weight:
            return False

        if self.settings.target_weight is None:
            return True

        return abs(value - self.settings.target_weight) <= self.settings.target_window
