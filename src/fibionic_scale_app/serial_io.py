from __future__ import annotations

import queue
import random
import re
import threading
import time
from abc import ABC, abstractmethod
from dataclasses import dataclass
from time import monotonic

from .models import Measurement, SerialSettings
from .parsing import parse_scale_output

SOURCE_MODE_SERIAL = "serial"
SOURCE_MODE_SIMULATION = "simulation"

SIM_PROFILE_STABLE = "stable_target"
SIM_PROFILE_BELOW = "below_target"
SIM_PROFILE_ABOVE = "above_target"
SIM_PROFILE_NOISY = "noisy_unstable"
SIM_PROFILE_STEP = "step_response"
SIM_PROFILE_RANDOM = "random_batches"

DEFAULT_SIM_INTERVAL = 0.18
DEFAULT_SIM_TARGET = 12.5
DEFAULT_PROBE_TIMEOUT = 0.2
DEFAULT_PROBE_WINDOW = 1.2
EXPECTED_FRAME_PATTERN = re.compile(r"^[+-]\s*\d+(?:\.\d+)?\s*g$", re.IGNORECASE)


@dataclass(slots=True)
class StreamEvent:
    kind: str
    message: str = ""
    raw_text: str = ""
    measurement: Measurement | None = None


@dataclass(slots=True)
class SimulationConfig:
    update_interval: float = DEFAULT_SIM_INTERVAL
    idle_steps: int = 10
    approach_steps: int = 8
    settle_steps: int = 10
    stable_steps: int = 18
    removal_steps: int = 8


def source_mode_options() -> list[tuple[str, str]]:
    return [
        (SOURCE_MODE_SERIAL, "Echte Waage"),
        (SOURCE_MODE_SIMULATION, "Simulation"),
    ]


def simulation_profile_options() -> list[tuple[str, str]]:
    return [
        (SIM_PROFILE_STABLE, "Stable at target"),
        (SIM_PROFILE_BELOW, "Below target"),
        (SIM_PROFILE_ABOVE, "Above target"),
        (SIM_PROFILE_NOISY, "Noisy / unstable"),
        (SIM_PROFILE_STEP, "Step response"),
        (SIM_PROFILE_RANDOM, "Random batches"),
    ]


def simulation_profile_label(profile: str) -> str:
    return dict(simulation_profile_options()).get(profile, "Stable at target")


def list_serial_ports() -> list[str]:
    try:
        from serial.tools import list_ports
    except ImportError:
        return []

    return [port.device for port in list_ports.comports()]


def preferred_serial_port(ports: list[str] | None = None) -> str | None:
    candidates = ports if ports is not None else list_serial_ports()
    if not candidates:
        return None

    return sorted(candidates, key=_serial_port_rank, reverse=True)[0]


def auto_detectable_serial_ports(ports: list[str] | None = None) -> list[str]:
    candidates = ports if ports is not None else list_serial_ports()
    return [device for device in candidates if _looks_like_usb_serial_port(device)]


def verified_serial_port(
    ports: list[str] | None = None,
    baudrate: int = 9600,
    timeout: float = DEFAULT_PROBE_TIMEOUT,
    probe_window: float = DEFAULT_PROBE_WINDOW,
) -> str | None:
    candidates = auto_detectable_serial_ports(ports)
    for device in sorted(candidates, key=_serial_port_rank, reverse=True):
        if probe_serial_port(device, baudrate=baudrate, timeout=timeout, probe_window=probe_window):
            return device
    return None


def probe_serial_port(
    device: str,
    baudrate: int = 9600,
    timeout: float = DEFAULT_PROBE_TIMEOUT,
    probe_window: float = DEFAULT_PROBE_WINDOW,
) -> bool:
    try:
        import serial
    except ImportError:
        return False

    try:
        with serial.Serial(
            port=device,
            baudrate=baudrate,
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            timeout=timeout,
        ) as connection:
            connection.reset_input_buffer()
            deadline = monotonic() + max(probe_window, timeout)
            valid_frames = 0
            while monotonic() < deadline:
                raw = connection.readline()
                if not raw:
                    continue
                measurement = parse_scale_output(raw.decode("ascii", errors="ignore"))
                if measurement is None:
                    continue
                if _matches_expected_scale_format(measurement):
                    valid_frames += 1
                    if valid_frames >= 2:
                        return True
    except Exception:
        return False

    return False


def _matches_expected_scale_format(measurement: Measurement) -> bool:
    raw_text = measurement.raw_text.strip()
    if not raw_text:
        return False

    return EXPECTED_FRAME_PATTERN.fullmatch(raw_text) is not None


def _looks_like_usb_serial_port(device: str) -> bool:
    lowered = device.strip().lower()
    usb_markers = ("usbserial", "usbmodem", "ttyusb", "usb", "acm", "ftdi", "wch", "silab", "ch340")
    return any(marker in lowered for marker in usb_markers)


def _serial_port_rank(device: str) -> tuple[int, int, str]:
    text = device.strip()
    lowered = text.lower()
    score = 0

    if lowered.startswith("/dev/cu."):
        score += 40
    if lowered.startswith("/dev/tty."):
        score += 10
    if lowered.startswith("com"):
        score += 30
    if any(marker in lowered for marker in ("usbserial", "usbmodem", "wch", "silab", "ftdi", "ch340", "ttyusb", "acm")):
        score += 100
    elif "usb" in lowered:
        score += 60

    return score, -len(text), text


class ScaleSource(ABC):
    def __init__(self) -> None:
        self.events: queue.Queue[StreamEvent] = queue.Queue()
        self._thread: threading.Thread | None = None
        self._stop_event = threading.Event()
        self._pause_event = threading.Event()
        self._connected = False
        self.status = "idle"

    @property
    @abstractmethod
    def source_name(self) -> str:
        raise NotImplementedError

    @property
    def current_port_name(self) -> str:
        return self.source_name

    @property
    def is_connected(self) -> bool:
        return self._connected

    def is_alive(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    def start(self) -> None:
        if self.is_alive():
            return

        self._stop_event.clear()
        self._pause_event.clear()
        self.status = "starting"
        self._thread = threading.Thread(target=self._run_wrapper, daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop_event.set()
        self.status = "stopping"

    def pause(self) -> None:
        self._pause_event.set()
        self.status = "paused"

    def resume(self) -> None:
        self._pause_event.clear()
        self.status = "running"

    def update_target_weight(self, target_weight: float | None) -> None:
        return None

    def reset_cycle(self) -> None:
        return None

    def _run_wrapper(self) -> None:
        try:
            self._run_source()
        except Exception as exc:
            self.events.put(StreamEvent(kind="error", message=str(exc)))
        finally:
            self._connected = False
            self.status = "stopped"
            self.events.put(StreamEvent(kind="stopped", message="Messquelle gestoppt."))

    @abstractmethod
    def _run_source(self) -> None:
        raise NotImplementedError

    def _wait_if_paused(self) -> None:
        while self._pause_event.is_set() and not self._stop_event.is_set():
            time.sleep(0.05)

    def _emit_connected(self, message: str) -> None:
        self._connected = True
        self.status = "running"
        self.events.put(StreamEvent(kind="connected", message=message))

    def _emit_measurement_frame(self, frame_text: str) -> None:
        measurement = parse_scale_output(frame_text)
        text = frame_text.strip()
        if measurement is None:
            self.events.put(StreamEvent(kind="raw", raw_text=text))
            return

        self.events.put(
            StreamEvent(
                kind="measurement",
                raw_text=text,
                measurement=measurement,
            )
        )


class SerialScaleSource(ScaleSource):
    def __init__(self, serial_settings: SerialSettings) -> None:
        super().__init__()
        self.serial_settings = serial_settings

    @property
    def source_name(self) -> str:
        return self.serial_settings.port

    def _run_source(self) -> None:
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
            self._emit_connected(
                f"Verbunden mit {self.serial_settings.port} @ {self.serial_settings.baudrate} Baud."
            )

            while not self._stop_event.is_set():
                self._wait_if_paused()
                raw = connection.readline()
                if not raw:
                    continue

                self._emit_measurement_frame(raw.decode("ascii", errors="ignore"))


class SimulatedScaleSource(ScaleSource):
    def __init__(
        self,
        profile: str = SIM_PROFILE_STABLE,
        target_weight: float | None = None,
        config: SimulationConfig | None = None,
        seed: int | None = None,
    ) -> None:
        super().__init__()
        self.profile = profile
        self.target_weight = target_weight
        self.config = config or SimulationConfig()
        self._rng = random.Random(seed)
        self._reset_event = threading.Event()

    @property
    def source_name(self) -> str:
        return "SIMULATED_SCALE"

    def update_target_weight(self, target_weight: float | None) -> None:
        self.target_weight = target_weight

    def reset_cycle(self) -> None:
        self._reset_event.set()

    def _run_source(self) -> None:
        self._emit_connected(f"Simulation aktiv: {simulation_profile_label(self.profile)}.")

        while not self._stop_event.is_set():
            cycle_values = self._build_cycle_values()
            for value in cycle_values:
                if self._stop_event.is_set():
                    return
                if self._reset_event.is_set():
                    self._reset_event.clear()
                    break

                self._wait_if_paused()
                self._emit_measurement_frame(self._format_frame(value))
                time.sleep(self.config.update_interval)

    def _build_cycle_values(self) -> list[float]:
        target = self._resolved_target_weight()
        target_delta = max(0.35, target * 0.05)
        stable_noise = max(0.01, target * 0.0008)
        unstable_noise = max(0.18, target * 0.01)
        profile = self.profile
        if profile == SIM_PROFILE_RANDOM:
            profile = self._rng.choice([SIM_PROFILE_STABLE, SIM_PROFILE_STABLE, SIM_PROFILE_BELOW, SIM_PROFILE_ABOVE, SIM_PROFILE_STEP])

        if profile == SIM_PROFILE_BELOW:
            final_value = max(0.0, target - target_delta)
            return self._compose_cycle(target, final_value, stable_noise, overshoot=0.0)

        if profile == SIM_PROFILE_ABOVE:
            final_value = target + target_delta
            return self._compose_cycle(target, final_value, stable_noise, overshoot=0.0)

        if profile == SIM_PROFILE_NOISY:
            final_value = target + self._rng.uniform(-target_delta * 0.2, target_delta * 0.2)
            return self._compose_cycle(target, final_value, unstable_noise, overshoot=0.0, unstable=True)

        if profile == SIM_PROFILE_STEP:
            final_value = target + self._rng.uniform(-stable_noise, stable_noise)
            overshoot = max(0.25, target * 0.03)
            return self._compose_cycle(target, final_value, stable_noise, overshoot=overshoot)

        final_value = target + self._rng.uniform(-stable_noise * 2.0, stable_noise * 2.0)
        return self._compose_cycle(target, final_value, stable_noise, overshoot=0.0)

    def _compose_cycle(
        self,
        target: float,
        final_value: float,
        noise: float,
        overshoot: float,
        unstable: bool = False,
    ) -> list[float]:
        values: list[float] = []
        values.extend(self._idle_values())

        if overshoot > 0:
            values.extend(self._approach_values(final_value + overshoot))
            values.extend(self._settle_values(final_value, overshoot))
        else:
            values.extend(self._approach_values(final_value))
            values.extend(self._settle_values(final_value, max(noise * 3.0, target * 0.01)))

        if unstable:
            values.extend(self._unstable_values(final_value, noise))
        else:
            values.extend(self._stable_values(final_value, noise))

        values.extend(self._removal_values(final_value))
        return values

    def _idle_values(self) -> list[float]:
        return [self._rng.uniform(-0.01, 0.01) for _ in range(self.config.idle_steps)]

    def _approach_values(self, final_value: float) -> list[float]:
        values: list[float] = []
        for step in range(1, self.config.approach_steps + 1):
            factor = step / self.config.approach_steps
            values.append((final_value * factor) + self._rng.uniform(-0.04, 0.04))
        return values

    def _settle_values(self, final_value: float, amplitude: float) -> list[float]:
        values: list[float] = []
        for step in range(self.config.settle_steps):
            decay = 1.0 - (step / max(1, self.config.settle_steps))
            direction = -1 if step % 2 else 1
            offset = direction * amplitude * decay * 0.8
            values.append(final_value + offset + self._rng.uniform(-0.02, 0.02))
        return values

    def _stable_values(self, final_value: float, noise: float) -> list[float]:
        return [final_value + self._rng.uniform(-noise, noise) for _ in range(self.config.stable_steps)]

    def _unstable_values(self, final_value: float, noise: float) -> list[float]:
        values: list[float] = []
        for step in range(self.config.stable_steps):
            swing = noise * (0.8 + 0.35 * self._rng.random())
            direction = -1 if step % 2 else 1
            values.append(final_value + (direction * swing) + self._rng.uniform(-noise * 0.25, noise * 0.25))
        return values

    def _removal_values(self, final_value: float) -> list[float]:
        values: list[float] = []
        for step in range(self.config.removal_steps, 0, -1):
            factor = step / self.config.removal_steps
            values.append((final_value * factor) + self._rng.uniform(-0.04, 0.04))
        values.extend(self._idle_values())
        return values

    def _resolved_target_weight(self) -> float:
        if self.target_weight is not None and self.target_weight > 0:
            return self.target_weight
        return DEFAULT_SIM_TARGET

    @staticmethod
    def _format_frame(value: float) -> str:
        sign = "+" if value >= 0 else "-"
        body = f"{abs(value):>8.2f}"
        return f"{sign}{body}g\r\n"
