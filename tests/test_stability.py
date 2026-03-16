from __future__ import annotations

import unittest

from fibionic_scale_app.models import CaptureSettings, Measurement
from fibionic_scale_app.stability import WeightCaptureEngine


class WeightCaptureEngineTests(unittest.TestCase):
    def measurement(self, value: float) -> Measurement:
        return Measurement(value=value, raw_text=f"{value:.3f}")

    def test_creates_pending_capture_after_stable_values(self) -> None:
        engine = WeightCaptureEngine(
            CaptureSettings(
                target_weight=12.5,
                target_window=0.3,
                stability_tolerance=0.04,
                stable_samples=4,
                rearm_threshold=0.1,
                minimum_weight=0.05,
            )
        )

        state = None
        for value in (12.49, 12.50, 12.51, 12.50):
            state = engine.process(self.measurement(value))

        self.assertIsNotNone(state)
        assert state is not None
        self.assertTrue(state.stable)
        self.assertTrue(state.within_target)
        self.assertAlmostEqual(state.new_candidate or 0.0, 12.5, places=2)
        self.assertAlmostEqual(engine.peek_pending_capture() or 0.0, 12.5, places=2)

    def test_rejects_values_outside_target_window(self) -> None:
        engine = WeightCaptureEngine(
            CaptureSettings(
                target_weight=10.0,
                target_window=0.2,
                stability_tolerance=0.03,
                stable_samples=4,
            )
        )

        for value in (11.00, 11.01, 11.00, 11.02):
            state = engine.process(self.measurement(value))

        self.assertFalse(state.within_target)
        self.assertIsNone(state.new_candidate)
        self.assertIsNone(engine.peek_pending_capture())

    def test_rearms_after_item_is_removed(self) -> None:
        engine = WeightCaptureEngine(
            CaptureSettings(
                target_weight=5.0,
                target_window=0.2,
                stability_tolerance=0.03,
                stable_samples=3,
                rearm_threshold=0.1,
            )
        )

        for value in (4.99, 5.00, 5.01):
            engine.process(self.measurement(value))

        self.assertIsNotNone(engine.commit_pending_capture())
        state = engine.process(self.measurement(0.02))

        self.assertTrue(state.rearmed)
        self.assertTrue(state.armed)


if __name__ == "__main__":
    unittest.main()
