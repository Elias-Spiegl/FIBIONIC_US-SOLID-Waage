from __future__ import annotations

from collections import deque
from dataclasses import dataclass
from statistics import mean

from .models import CaptureSettings, Measurement
from .weight_precision import quantize_weight_value


def build_capture_settings(target_weight: float, target_window: float) -> CaptureSettings:
    if target_window <= 0:
        raise ValueError("Die Abweichung muss größer als 0 sein.")

    stable_samples = 10
    base_tolerance = 0.02 + (max(target_weight, 0.0) * 0.001)
    base_tolerance = min(max(0.02, base_tolerance), max(0.02, target_window * 0.60))

    return CaptureSettings(
        target_weight=target_weight,
        target_window=target_window,
        base_stability_tolerance=base_tolerance,
        stable_samples=stable_samples,
        require_confirmation=False,
    )


@dataclass(slots=True)
class CaptureState:
    measurement: Measurement
    stable: bool
    within_target: bool
    armed: bool
    pending_capture: float | None
    new_candidate: float | None
    spread: float | None
    effective_tolerance: float
    rearm_threshold: float
    rearmed: bool = False


class WeightCaptureEngine:
    def __init__(self, settings: CaptureSettings):
        self.settings = settings
        self.history: deque[float] = deque(maxlen=max(3, settings.stable_samples))
        self.armed = True
        self.pending_capture: float | None = None
        self.observed_spread: float | None = None

    def update_settings(self, settings: CaptureSettings) -> None:
        self.settings = settings
        self.history = deque(maxlen=max(3, settings.stable_samples))
        self.armed = True
        self.pending_capture = None
        self.observed_spread = None

    def reset(self) -> None:
        self.history.clear()
        self.armed = True
        self.pending_capture = None
        self.observed_spread = None

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

    def effective_tolerance(self) -> float:
        base = max(0.005, self.settings.base_stability_tolerance)
        window_cap = max(0.02, self.settings.target_window * 0.25)

        if self.observed_spread is None:
            return min(window_cap, base)

        return min(window_cap, max(base, self.observed_spread * 1.35))

    def effective_rearm_threshold(self) -> float:
        target_weight = abs(self.settings.target_weight or 0.0)
        dynamic = max(self.settings.target_window * 2.0, target_weight * 0.15)
        return max(0.10, min(2.00, dynamic))

    def process(self, measurement: Measurement) -> CaptureState:
        value = quantize_weight_value(measurement.value)
        rearm_threshold = self.effective_rearm_threshold()
        rearmed = False
        if self.pending_capture is None and not self.armed and abs(value) <= rearm_threshold:
            self.armed = True
            self.history.clear()
            rearmed = True

        self.history.append(value)
        within_target = self._matches_target(value)
        stable = False
        spread: float | None = None
        new_candidate: float | None = None
        effective_tolerance = self.effective_tolerance()

        if len(self.history) == self.history.maxlen:
            spread = max(self.history) - min(self.history)
            if within_target:
                self._remember_spread(spread)
                effective_tolerance = self.effective_tolerance()
            stable = spread <= effective_tolerance

        if self.armed and self.pending_capture is None and stable and within_target:
            new_candidate = quantize_weight_value(mean(self.history))
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
            effective_tolerance=effective_tolerance,
            rearm_threshold=rearm_threshold,
            rearmed=rearmed,
        )

    def _matches_target(self, value: float) -> bool:
        if self.settings.target_weight is None:
            return True

        return abs(value - self.settings.target_weight) <= self.settings.target_window

    def _remember_spread(self, spread: float) -> None:
        if self.observed_spread is None:
            self.observed_spread = spread
            return

        self.observed_spread = (self.observed_spread * 0.7) + (spread * 0.3)
