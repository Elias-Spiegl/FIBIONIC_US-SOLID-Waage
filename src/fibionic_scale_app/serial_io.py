from __future__ import annotations

import queue
import random
import threading
import time
from dataclasses import dataclass

from .models import Measurement, SerialSettings
from .parsing import parse_scale_output


@dataclass(slots=True)
class StreamEvent:
    kind: str
    message: str = ""
    raw_text: str = ""
    measurement: Measurement | None = None


def list_serial_ports() -> list[str]:
    try:
        from serial.tools import list_ports
    except ImportError:
        return []

    return [port.device for port in list_ports.comports()]


class ScaleStreamWorker(threading.Thread):
    def __init__(
        self,
        serial_settings: SerialSettings,
        simulate: bool = False,
        simulation_target: float | None = None,
    ) -> None:
        super().__init__(daemon=True)
        self.serial_settings = serial_settings
        self.simulate = simulate
        self.simulation_target = simulation_target
        self.events: queue.Queue[StreamEvent] = queue.Queue()
        self._stop_event = threading.Event()

    def stop(self) -> None:
        self._stop_event.set()

    def run(self) -> None:
        try:
            if self.simulate:
                self._run_simulation()
            else:
                self._run_serial()
        except Exception as exc:
            self.events.put(StreamEvent(kind="error", message=str(exc)))
        finally:
            self.events.put(StreamEvent(kind="stopped", message="Messquelle gestoppt."))

    def _run_serial(self) -> None:
        try:
            import serial
        except ImportError as exc:
            raise RuntimeError(
                "Das Paket 'pyserial' fehlt. Bitte installiere zuerst die Projekt-Abhängigkeiten."
            ) from exc

        with serial.Serial(
            port=self.serial_settings.port,
            baudrate=self.serial_settings.baudrate,
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            timeout=self.serial_settings.timeout,
        ) as connection:
            connection.reset_input_buffer()
            self.events.put(
                StreamEvent(
                    kind="connected",
                    message=f"Verbunden mit {self.serial_settings.port} @ {self.serial_settings.baudrate} Baud.",
                )
            )

            while not self._stop_event.is_set():
                raw = connection.readline()
                if not raw:
                    continue

                measurement = parse_scale_output(raw)
                text = raw.decode("ascii", errors="ignore").strip()
                if measurement is None:
                    self.events.put(StreamEvent(kind="raw", raw_text=text))
                    continue

                self.events.put(
                    StreamEvent(
                        kind="measurement",
                        raw_text=text,
                        measurement=measurement,
                    )
                )

    def _run_simulation(self) -> None:
        rng = random.Random(42)
        target = self.simulation_target if self.simulation_target and self.simulation_target > 0 else 12.5
        phases = [
            *([0.00] * 8),
            0.35 * target,
            0.62 * target,
            0.83 * target,
            0.94 * target,
            0.98 * target,
            1.00 * target,
            *([target + rng.uniform(-0.03, 0.03) for _ in range(18)]),
            *([0.00] * 12),
        ]

        self.events.put(
            StreamEvent(kind="connected", message="Simulationsmodus aktiv. Keine echte Waage verbunden.")
        )

        while not self._stop_event.is_set():
            for value in phases:
                if self._stop_event.is_set():
                    return

                frame = self._format_simulated_frame(value + rng.uniform(-0.01, 0.01))
                measurement = parse_scale_output(frame)
                if measurement is not None:
                    self.events.put(
                        StreamEvent(
                            kind="measurement",
                            raw_text=frame.strip(),
                            measurement=measurement,
                        )
                    )
                time.sleep(0.18)

    @staticmethod
    def _format_simulated_frame(value: float) -> str:
        sign = "-" if value < 0 else "+"
        body = f"{abs(value):>7.2f}"
        return f"{sign} {body}g  \r\n"
