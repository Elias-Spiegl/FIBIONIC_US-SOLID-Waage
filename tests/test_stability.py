from __future__ import annotations

import unittest

from fibionic_scale_app.models import Measurement
from fibionic_scale_app.stability import WeightCaptureEngine, build_capture_settings


class WeightCaptureEngineTests(unittest.TestCase):
    def measurement(self, value: float) -> Measurement:
        return Measurement(value=value, raw_text=f"{value:.3f}")

    def test_creates_pending_capture_after_stable_values(self) -> None:
        settings = build_capture_settings(12.5, 0.2)
        engine = WeightCaptureEngine(settings)

        state = None
        values = [12.49, 12.50, 12.51, 12.50, 12.49, 12.50, 12.50, 12.51, 12.50, 12.49]
        for value in values:
            state = engine.process(self.measurement(value))

        self.assertIsNotNone(state)
        assert state is not None
        self.assertTrue(state.stable)
        self.assertTrue(state.within_target)
        self.assertAlmostEqual(state.new_candidate or 0.0, 12.5, places=2)
        self.assertAlmostEqual(engine.peek_pending_capture() or 0.0, 12.5, places=2)

    def test_rejects_values_outside_target_window(self) -> None:
        settings = build_capture_settings(10.0, 0.2)
        engine = WeightCaptureEngine(settings)

        for value in (11.00, 11.01, 11.00, 11.02, 11.01, 11.00, 11.01, 11.00):
            state = engine.process(self.measurement(value))

        self.assertFalse(state.within_target)
        self.assertIsNone(state.new_candidate)
        self.assertIsNone(engine.peek_pending_capture())

    def test_rearms_after_item_is_removed(self) -> None:
        settings = build_capture_settings(5.0, 0.2)
        engine = WeightCaptureEngine(settings)

        for value in (4.99, 5.00, 5.01, 5.00, 4.99, 5.00, 5.01, 5.00, 4.99, 5.00):
            engine.process(self.measurement(value))

        self.assertIsNotNone(engine.commit_pending_capture())
        state = engine.process(self.measurement(0.02))

        self.assertTrue(state.rearmed)
        self.assertTrue(state.armed)

    def test_build_capture_settings_uses_fixed_sample_count_and_dynamic_tolerance(self) -> None:
        light = build_capture_settings(5.0, 0.5)
        heavy = build_capture_settings(40.0, 0.5)

        self.assertEqual(light.stable_samples, 10)
        self.assertEqual(heavy.stable_samples, 10)
        self.assertAlmostEqual(light.base_stability_tolerance, 0.025)
        self.assertAlmostEqual(heavy.base_stability_tolerance, 0.06)


if __name__ == "__main__":
    unittest.main()
