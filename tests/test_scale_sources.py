from __future__ import annotations

import time
import unittest
from unittest.mock import patch

from fibionic_scale_app.serial_io import (
    auto_detectable_serial_ports,
    SIM_PROFILE_BELOW,
    SIM_PROFILE_NOISY,
    SIM_PROFILE_STABLE,
    SimulationConfig,
    SimulatedScaleSource,
    probe_serial_port,
    preferred_serial_port,
    verified_serial_port,
)


class ScaleSourceTests(unittest.TestCase):
    def test_preferred_serial_port_prefers_usb_devices(self) -> None:
        ports = ["/dev/cu.Bluetooth-Incoming-Port", "/dev/cu.usbserial-130", "/dev/tty.debug"]
        self.assertEqual(preferred_serial_port(ports), "/dev/cu.usbserial-130")

    def test_auto_detectable_serial_ports_excludes_bluetooth(self) -> None:
        ports = ["/dev/cu.Bluetooth-Incoming-Port", "/dev/cu.usbserial-130", "/dev/cu.debug"]
        self.assertEqual(auto_detectable_serial_ports(ports), ["/dev/cu.usbserial-130"])

    def test_simulated_source_emits_connected_and_measurements(self) -> None:
        source = SimulatedScaleSource(
            profile=SIM_PROFILE_STABLE,
            target_weight=40.0,
            seed=42,
            config=SimulationConfig(
                update_interval=0.001,
                idle_steps=1,
                approach_steps=2,
                settle_steps=2,
                stable_steps=4,
                removal_steps=2,
            ),
        )

        source.start()
        received_connected = False
        received_measurement = False
        deadline = time.time() + 0.3
        while time.time() < deadline and not (received_connected and received_measurement):
            try:
                event = source.events.get(timeout=0.02)
            except Exception:
                continue
            if event.kind == "connected":
                received_connected = True
            if event.kind == "measurement" and event.measurement is not None:
                received_measurement = True

        source.stop()

        self.assertTrue(received_connected)
        self.assertTrue(received_measurement)

    def test_verified_serial_port_prefers_first_port_with_valid_scale_frame(self) -> None:
        with patch("fibionic_scale_app.serial_io.probe_serial_port") as probe:
            probe.side_effect = lambda device, **_: device == "/dev/cu.usbserial-130"
            verified = verified_serial_port(
                ["/dev/cu.Bluetooth-Incoming-Port", "/dev/cu.usbserial-130", "/dev/cu.other"]
            )

        self.assertEqual(verified, "/dev/cu.usbserial-130")

    def test_probe_serial_port_accepts_expected_gram_frame(self) -> None:
        class DummySerial:
            EIGHTBITS = 8
            PARITY_NONE = "N"
            STOPBITS_ONE = 1

            class Serial:
                def __init__(self, **kwargs):
                    self.lines = [b"+    44.00g\r\n", b"+    44.00g\r\n"]

                def __enter__(self):
                    return self

                def __exit__(self, exc_type, exc, tb):
                    return False

                def reset_input_buffer(self):
                    return None

                def readline(self):
                    return self.lines.pop(0) if self.lines else b""

        with patch.dict("sys.modules", {"serial": DummySerial}):
            self.assertTrue(probe_serial_port("/dev/cu.usbserial-130", probe_window=0.01))

    def test_probe_serial_port_rejects_unexpected_unit(self) -> None:
        class DummySerial:
            EIGHTBITS = 8
            PARITY_NONE = "N"
            STOPBITS_ONE = 1

            class Serial:
                def __init__(self, **kwargs):
                    self.lines = [b"+    44.00kg\r\n", b"+    44.00kg\r\n"]

                def __enter__(self):
                    return self

                def __exit__(self, exc_type, exc, tb):
                    return False

                def reset_input_buffer(self):
                    return None

                def readline(self):
                    return self.lines.pop(0) if self.lines else b""

        with patch.dict("sys.modules", {"serial": DummySerial}):
            self.assertFalse(probe_serial_port("/dev/cu.usbserial-130", probe_window=0.01))

    def test_probe_serial_port_rejects_frames_without_explicit_gram_unit(self) -> None:
        class DummySerial:
            EIGHTBITS = 8
            PARITY_NONE = "N"
            STOPBITS_ONE = 1

            class Serial:
                def __init__(self, **kwargs):
                    self.lines = [b"+    44.00\r\n", b"+    44.00\r\n"]

                def __enter__(self):
                    return self

                def __exit__(self, exc_type, exc, tb):
                    return False

                def reset_input_buffer(self):
                    return None

                def readline(self):
                    return self.lines.pop(0) if self.lines else b""

        with patch.dict("sys.modules", {"serial": DummySerial}):
            self.assertFalse(probe_serial_port("/dev/cu.usbserial-130", probe_window=0.01))

    def test_below_target_profile_stays_below_target(self) -> None:
        source = SimulatedScaleSource(profile=SIM_PROFILE_BELOW, target_weight=40.0, seed=7)
        values = source._build_cycle_values()
        stable_window = values[-(source.config.stable_steps + source.config.removal_steps + source.config.idle_steps) : -(
            source.config.removal_steps + source.config.idle_steps
        )]

        self.assertTrue(stable_window)
        self.assertLess(max(stable_window), 40.0)

    def test_noisy_profile_has_large_spread(self) -> None:
        source = SimulatedScaleSource(profile=SIM_PROFILE_NOISY, target_weight=20.0, seed=12)
        values = source._build_cycle_values()
        unstable_window = values[-(source.config.stable_steps + source.config.removal_steps + source.config.idle_steps) : -(
            source.config.removal_steps + source.config.idle_steps
        )]

        self.assertTrue(unstable_window)
        self.assertGreater(max(unstable_window) - min(unstable_window), 0.15)


if __name__ == "__main__":
    unittest.main()
