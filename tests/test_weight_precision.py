from __future__ import annotations

import unittest

from fibionic_scale_app.weight_precision import format_weight_value, quantize_weight_value


class WeightPrecisionTests(unittest.TestCase):
    def test_quantizes_to_two_decimals_with_half_up_rounding(self) -> None:
        self.assertEqual(quantize_weight_value(13.344), 13.34)
        self.assertEqual(quantize_weight_value(13.345), 13.35)

    def test_format_drops_insignificant_trailing_zeros(self) -> None:
        self.assertEqual(format_weight_value(13.3), "13.3")
        self.assertEqual(format_weight_value(13.30), "13.3")
        self.assertEqual(format_weight_value(13.345), "13.35")
        self.assertEqual(format_weight_value(13.0), "13")


if __name__ == "__main__":
    unittest.main()
