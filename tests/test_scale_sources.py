from __future__ import annotations

import time
import unittest

from fibionic_scale_app.serial_io import (
    SIM_PROFILE_BELOW,
    SIM_PROFILE_NOISY,
    SIM_PROFILE_STABLE,
    SimulationConfig,
    SimulatedScaleSource,
    preferred_serial_port,
)


class ScaleSourceTests(unittest.TestCase):
    def test_preferred_serial_port_prefers_usb_devices(self) -> None:
        ports = ["/dev/cu.Bluetooth-Incoming-Port", "/dev/cu.usbserial-130", "/dev/tty.debug"]
        self.assertEqual(preferred_serial_port(ports), "/dev/cu.usbserial-130")

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
