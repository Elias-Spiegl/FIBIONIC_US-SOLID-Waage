from __future__ import annotations

import unittest

from fibionic_scale_app.parsing import parse_scale_output


class ParseScaleOutputTests(unittest.TestCase):
    def test_parses_fixed_width_scale_frame(self) -> None:
        measurement = parse_scale_output("+   12.34kg\r\n")

        self.assertIsNotNone(measurement)
        assert measurement is not None
        self.assertEqual(measurement.value, 12.34)
        self.assertEqual(measurement.unit, "KG")

    def test_parses_negative_value_with_decimal_comma(self) -> None:
        measurement = parse_scale_output("-    0,45lb\r\n")

        self.assertIsNotNone(measurement)
        assert measurement is not None
        self.assertAlmostEqual(measurement.value, -0.45)
        self.assertEqual(measurement.unit, "LB")

    def test_ignores_non_measurement_lines(self) -> None:
        self.assertIsNone(parse_scale_output("READY"))


if __name__ == "__main__":
    unittest.main()
