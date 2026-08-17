"""Instrument and simulation layer for the SID bench application.

Nothing in this module opens hardware at import time.  Sessions are explicit,
serialized, and releasable so the physical front panels remain usable.
"""

from __future__ import annotations

import csv
import json
import math
import os
import re
import struct
import threading
import time
import zlib
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Callable, Iterable


VISA_DLL = r"C:\Windows\System32\visa64.dll"
KEYSIGHT_BIN = r"C:\Program Files\Keysight\IO Libraries Suite\bin"

KNOWN_ADDRESSES = {
    "pa": "",
    "load": "",
    "psu": "",
    "scope": "",
}

IDENTITY_TOKENS = {
    "pa": ("PA2201", "INTEGRAVISION"),
    "load": ("63206",),
    "psu": ("E36312",),
    "scope": ("MSOX4024", "MSO-X 4024A"),
}

_DLL_DIRECTORY_HANDLES: list[Any] = []


def prepare_visa_runtime() -> None:
    """Make Keysight VISA dependency DLLs visible before importing PyVISA."""
    if not os.path.isdir(KEYSIGHT_BIN):
        return
    paths = os.environ.get("PATH", "").split(os.pathsep)
    if KEYSIGHT_BIN.lower() not in {path.lower() for path in paths}:
        os.environ["PATH"] = KEYSIGHT_BIN + os.pathsep + os.environ.get("PATH", "")
    if hasattr(os, "add_dll_directory"):
        # The returned handle must stay alive for the directory to remain active.
        if not _DLL_DIRECTORY_HANDLES:
            _DLL_DIRECTORY_HANDLES.append(os.add_dll_directory(KEYSIGHT_BIN))


class InstrumentError(RuntimeError):
    """A concise, operator-facing instrument error."""


class RequiredInstrumentError(InstrumentError):
    """An error that must abort an automated run."""


class InvalidDataError(InstrumentError):
    """An error indicating non-numeric, non-finite, sentinel, or out-of-range measurement data."""


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds")


def finite_float(value: Any, label: str) -> float:
    try:
        result = float(str(value).strip())
    except (TypeError, ValueError) as exc:
        raise InstrumentError(f"{label} returned non-numeric data: {value!r}") from exc
    if not math.isfinite(result):
        raise InstrumentError(f"{label} returned non-finite data: {value!r}")
    return result


@dataclass
class InstrumentSnapshot:
    instrument: str
    values: dict[str, float | str | bool | None]
    timestamp: str = field(default_factory=utc_now)
    valid: bool = True
    warning: str = ""
    status: str = "Connected"


@dataclass
class SupplyChannel:
    channel: int
    role: str
    displayed: bool = True
    enabled: bool = False
    contributes_loss: bool = True
    voltage_set: float | None = None
    voltage: float | None = None
    current_limit: float | None = None
    current: float | None = None


class VisaManager:
    """Lazy shared VISA resource manager with model-based discovery."""

    def __init__(self, library: str = VISA_DLL):
        self.library = library
        self._rm = None
        self._lock = threading.RLock()
        self.discovery_errors: dict[str, str] = {}

    @property
    def available(self) -> bool:
        return os.path.isfile(self.library)

    def open(self):
        with self._lock:
            if self._rm is None:
                prepare_visa_runtime()
                try:
                    import pyvisa
                except ImportError as exc:
                    raise InstrumentError("PyVISA is not installed") from exc
                if not self.available:
                    raise InstrumentError(
                        "Keysight VISA was not found. Install/repair Connection Expert 2026."
                    )
                self._rm = pyvisa.ResourceManager(self.library)
            return self._rm

    def list_resources(self) -> tuple[str, ...]:
        return tuple(self.open().list_resources())

    def discover(self) -> dict[str, dict[str, str]]:
        found: dict[str, dict[str, str]] = {}
        self.discovery_errors = {}
        for address in self.list_resources():
            if address.endswith("000000000001::0::INSTR") or "000000000001" in address:
                continue
            session = None
            try:
                session = self.open().open_resource(address)
                session.timeout = 1500
                identity = str(session.query("*IDN?")).strip()
                upper = identity.upper()
                for kind, tokens in IDENTITY_TOKENS.items():
                    if any(token in upper for token in tokens):
                        found[kind] = {"address": address, "identity": identity}
            except Exception as exc:
                self.discovery_errors[address] = str(exc)
                continue
            finally:
                if session is not None:
                    try:
                        session.close()
                    except Exception:
                        pass
        return found

    def close(self) -> None:
        with self._lock:
            if self._rm is not None:
                try:
                    self._rm.close()
                finally:
                    self._rm = None


class VisaInstrument:
    """Base class with explicit persistent or temporary session ownership."""

    kind = "instrument"

    def __init__(self, manager: VisaManager, address: str, expected: Iterable[str]):
        self.manager = manager
        self.address = address
        self.expected = tuple(x.upper() for x in expected)
        self.identity = ""
        self._session = None
        self._persistent = False
        self._lock = threading.RLock()

    @property
    def connected(self) -> bool:
        return self._session is not None

    def connect(self, persistent: bool = True) -> str:
        with self._lock:
            if self._session is None:
                if not self.address:
                    raise InstrumentError(f"No VISA address is configured for {self.kind}")
                try:
                    self._session = self.manager.open().open_resource(self.address)
                    self._session.timeout = 3000
                    self.identity = str(self._session.query("*IDN?")).strip()
                    upper = self.identity.upper()
                    if self.expected and not any(token in upper for token in self.expected):
                        raise InstrumentError(
                            f"Unexpected {self.kind} identity at {self.address}: {self.identity}"
                        )
                except Exception as exc:
                    self._close_locked()
                    self._raise_clean(exc)
            self._persistent = self._persistent or persistent
            return self.identity

    def _raise_clean(self, exc: Exception):
        text = str(exc)
        if "VI_ERROR_NCIC" in text:
            raise InstrumentError(
                "VISA controller conflict (VI_ERROR_NCIC). Release Interactive IO or other VISA clients, then retry."
            ) from exc
        if isinstance(exc, InstrumentError):
            raise exc
        raise InstrumentError(f"{self.kind} communication failed: {text}") from exc

    @contextmanager
    def session(self):
        temporary = False
        with self._lock:
            if self._session is None:
                self.connect(persistent=False)
                temporary = True
            try:
                yield self._session
            except Exception as exc:
                self._close_locked()
                self._raise_clean(exc)
            finally:
                if temporary and not self._persistent:
                    self._close_locked()

    def _close_locked(self) -> None:
        if self._session is not None:
            try:
                self._session.close()
            except Exception:
                pass
        self._session = None
        self._persistent = False

    def release(self) -> None:
        with self._lock:
            if self._session is not None:
                for command in self.local_commands():
                    try:
                        self._session.write(command)
                        break
                    except Exception:
                        continue
            self._close_locked()

    def local_commands(self) -> tuple[str, ...]:
        return ("SYSTem:LOCal", "SYST:LOC")

    def read_snapshot(self, **_: Any) -> InstrumentSnapshot:
        raise NotImplementedError


class Chroma63206A(VisaInstrument):
    kind = "Chroma 63206A"

    def read_snapshot(self, **_: Any) -> InstrumentSnapshot:
        with self.session() as dev:
            current = finite_float(dev.query("MEASure:CURRent?"), "load current")
            warning = ""
            try:
                voltage = finite_float(dev.query("MEASure:VOLTage?"), "load voltage")
            except InstrumentError as exc:
                voltage = None
                warning = str(exc)
        return InstrumentSnapshot(
            "load", {"current": current, "voltage": voltage, "power": current * voltage if voltage is not None else None}, warning=warning
        )

    def set_current(self, amps: float) -> None:
        amps = finite_float(amps, "requested load current")
        if amps < 0:
            raise InstrumentError("Load current cannot be negative")
        with self.session() as dev:
            rating = self.reported_current_rating()
            if rating is not None and amps > rating:
                raise InstrumentError(f"{amps:g} A exceeds the load model's reported {rating:g} A rating")
            dev.write("MODE CCH")
            dev.write(f"CURRent:STATic:L1 {amps:.9g}")

    def reported_current_rating(self) -> float | None:
        match = re.search(r"63206A-\d+(?:\.\d+)?-(\d+(?:\.\d+)?)", self.identity.upper())
        return float(match.group(1)) if match else None

    def set_input(self, enabled: bool) -> None:
        with self.session() as dev:
            dev.write("LOAD ON" if enabled else "LOAD OFF")

    def safe_off(self) -> None:
        try:
            self.set_input(False)
        except Exception:
            pass


class PA2201A(VisaInstrument):
    kind = "Keysight PA2201A"

    def configure_dc_analysis(self, dev: Any) -> None:
        """Configure PA2201A Channels 1 and 2 for DC Power Quality measurement with 1s window."""
        dev.write("ANALyze:ENABle ON")
        dev.write("ANALyze:WINDow W_1S")
        dev.write("ANALyze:SOURce1:SYNC LINE")
        dev.write("ANALyze:MODE1 DC_MODE")
        dev.write("ANALyze:SOURce2:SYNC LINE")
        dev.write("ANALyze:MODE2 DC_MODE")

    def trigger_dc_measurements(self, dev: Any) -> None:
        """Trigger DC Power Quality measurement calculation on Channels 1 and 2."""
        dev.write("ANALyze:QUALity1:MEASure")
        dev.write("ANALyze:QUALity2:MEASure")

    def _query_syst_err(self, dev: Any) -> str:
        try:
            err = str(dev.query("SYSTem:ERRor?")).strip()
            if err and not err.startswith("+0") and not err.startswith("0"):
                return f" (Instrument error: {err})"
        except Exception:
            pass
        return ""

    def _parse_and_validate(self, raw: Any, label: str, max_abs: float = 2000.0) -> float:
        raw_str = str(raw).strip()
        try:
            val = float(raw_str)
        except (TypeError, ValueError) as exc:
            raise InvalidDataError(f"{label} returned non-numeric data: {raw_str!r}") from exc
        if not math.isfinite(val):
            raise InvalidDataError(f"{label} returned non-finite data: {raw_str!r}")
        # PA2200 sentinel/unready values are large numbers (e.g. 9.91E+37 or > 1e6)
        if abs(val) > max_abs or abs(val) >= 9e36:
            raise InvalidDataError(f"{label} returned sentinel/out-of-range value: {raw_str}")
        return val

    def read_dc_snapshot(self, dev: Any, settle_s: float = 1.2) -> dict[str, float]:
        """Perform verified PA2200 DC Power Quality measurement flow."""
        self.configure_dc_analysis(dev)
        self.trigger_dc_measurements(dev)
        if settle_s > 0:
            time.sleep(settle_s)

        values: dict[str, float] = {}

        # Vin: Channel 1 DC voltage
        try:
            raw_v1 = dev.query("ANALyze:QUALity1:VOLTage:DC?")
            values["vin"] = self._parse_and_validate(raw_v1, "CH1 DC voltage", max_abs=1500.0)
        except InvalidDataError as exc:
            inst_err = self._query_syst_err(dev)
            raise InvalidDataError(f"PA connected, but CH1 DC voltage query returned invalid data: {exc}{inst_err}") from exc
        except Exception as exc:
            inst_err = self._query_syst_err(dev)
            raise InstrumentError(f"PA connected, but CH1 DC voltage query failed: {exc}{inst_err}") from exc

        # Iin: Channel 1 DC current
        try:
            raw_i1 = dev.query("ANALyze:QUALity1:CURRent:DC?")
            values["iin"] = self._parse_and_validate(raw_i1, "CH1 DC current", max_abs=500.0)
        except InvalidDataError as exc:
            inst_err = self._query_syst_err(dev)
            raise InvalidDataError(f"PA connected, but CH1 DC current query returned invalid data: {exc}{inst_err}") from exc
        except Exception as exc:
            inst_err = self._query_syst_err(dev)
            raise InstrumentError(f"PA connected, but CH1 DC current query failed: {exc}{inst_err}") from exc

        # Vout: Channel 2 DC voltage
        try:
            raw_v2 = dev.query("ANALyze:QUALity2:VOLTage:DC?")
            values["vout"] = self._parse_and_validate(raw_v2, "CH2 DC voltage", max_abs=1500.0)
        except InvalidDataError as exc:
            inst_err = self._query_syst_err(dev)
            raise InvalidDataError(f"PA connected, but CH2 DC voltage query returned invalid data: {exc}{inst_err}") from exc
        except Exception as exc:
            inst_err = self._query_syst_err(dev)
            raise InstrumentError(f"PA connected, but CH2 DC voltage query failed: {exc}{inst_err}") from exc

        return values

    def read_snapshot(self, settle_s: float = 1.2, **_: Any) -> InstrumentSnapshot:
        with self.session() as dev:
            try:
                values = self.read_dc_snapshot(dev, settle_s=settle_s)
                return InstrumentSnapshot("pa", values, valid=True, status="Connected")
            except InvalidDataError as exc:
                return InstrumentSnapshot(
                    "pa",
                    {"vin": None, "iin": None, "vout": None},
                    valid=False,
                    warning=str(exc),
                    status="Connected · Invalid Data",
                )
            except InstrumentError as exc:
                return InstrumentSnapshot(
                    "pa",
                    {"vin": None, "iin": None, "vout": None},
                    valid=False,
                    warning=str(exc),
                    status="Connected · Read Error",
                )


class E36312A(VisaInstrument):
    kind = "Keysight E36312A"

    def read_snapshot(
        self, channels: Iterable[int] = (1, 2, 3), **_: Any
    ) -> InstrumentSnapshot:
        result: dict[str, float | bool] = {}
        with self.session() as dev:
            for channel in channels:
                ch = int(channel)
                result[f"ch{ch}_voltage"] = finite_float(
                    dev.query(f"MEASure:VOLTage? (@{ch})"), f"PSU CH{ch} voltage"
                )
                result[f"ch{ch}_current"] = finite_float(
                    dev.query(f"MEASure:CURRent? (@{ch})"), f"PSU CH{ch} current"
                )
                try:
                    outp = str(dev.query(f"OUTPut? (@{ch})")).strip()
                    result[f"ch{ch}_enabled"] = outp in {"1", "ON"}
                except Exception:
                    pass
        return InstrumentSnapshot("psu", result)

    def configure_channel(
        self, channel: int, voltage: float, current_limit: float, enabled: bool
    ) -> None:
        ch = int(channel)
        voltage = finite_float(voltage, "PSU voltage")
        current_limit = finite_float(current_limit, "PSU current limit")
        if ch not in (1, 2, 3) or voltage < 0 or current_limit < 0:
            raise InstrumentError("Invalid E36312A channel configuration")
        with self.session() as dev:
            dev.write(f"VOLTage {voltage:.9g}, (@{ch})")
            dev.write(f"CURRent {current_limit:.9g}, (@{ch})")
            dev.write(f"OUTPut {'ON' if enabled else 'OFF'}, (@{ch})")


def _decode_ieee_block(raw: bytes) -> bytes:
    if not raw.startswith(b"#") or len(raw) < 3:
        return raw.rstrip(b"\r\n")
    digits = int(raw[1:2])
    count = int(raw[2 : 2 + digits])
    start = 2 + digits
    return raw[start : start + count]


class MSOX4024A(VisaInstrument):
    kind = "Keysight MSOX4024A"

    def displayed_channels(self) -> list[int]:
        channels: list[int] = []
        with self.session() as dev:
            for channel in range(1, 5):
                try:
                    if int(float(dev.query(f":CHANnel{channel}:DISPlay?"))):
                        channels.append(channel)
                except Exception:
                    continue
        return channels

    def read_snapshot(self, **_: Any) -> InstrumentSnapshot:
        return InstrumentSnapshot("scope", {"displayed_channels": ",".join(map(str, self.displayed_channels()))})

    def capture(self, png_path: Path, csv_path: Path, timeout_s: float = 15.0) -> tuple[list[int], int]:
        """Freeze acquisition with :STOP, save displayed PNG and channel waveform CSVs, then restore :RUN."""
        png_path.parent.mkdir(parents=True, exist_ok=True)
        channels: list[int] = []
        waveforms: dict[int, tuple[list[float], list[float]]] = {}
        with self.session() as dev:
            previous_timeout = getattr(dev, "timeout", None)
            stage = "reading displayed channels"
            try:
                if previous_timeout is not None:
                    dev.timeout = max(int(timeout_s * 1000), int(previous_timeout))
                for channel in range(1, 5):
                    if int(float(dev.query(f":CHANnel{channel}:DISPlay?"))):
                        channels.append(channel)
                if not channels:
                    raise InstrumentError("No analog scope channels are displayed")

                stage = "freezing scope acquisition"
                dev.write(":STOP")
                time.sleep(0.05)

                stage = "saving screen PNG"
                dev.write(":HARDcopy:INKSaver ON")
                dev.write(":DISPlay:DATA? PNG, COLor")
                png_path.write_bytes(_decode_ieee_block(dev.read_raw()))

                stage = "preparing waveform export"
                dev.write(":WAVeform:POINts:MODE RAW")
                dev.write(":WAVeform:FORMat WORD")
                dev.write(":WAVeform:BYTeorder LSBFirst")
                for channel in channels:
                    stage = f"exporting CH{channel} waveform"
                    dev.write(f":WAVeform:SOURce CHANnel{channel}")
                    pre = [float(x) for x in dev.query(":WAVeform:PREamble?").split(",")]
                    raw_values = dev.query_binary_values(
                        ":WAVeform:DATA?", datatype="h", is_big_endian=False, container=list
                    )
                    xinc, xorg, xref = pre[4], pre[5], pre[6]
                    yinc, yorg, yref = pre[7], pre[8], pre[9]
                    times = [(index - xref) * xinc + xorg for index in range(len(raw_values))]
                    volts = [(raw - yref) * yinc + yorg for raw in raw_values]
                    waveforms[channel] = (times, volts)
            except InstrumentError:
                raise
            except Exception as exc:
                raise InstrumentError(f"Scope capture failed while {stage}: {exc}") from exc
            finally:
                # Always return the front panel to live run acquisition.
                try:
                    dev.write(":RUN")
                except Exception:
                    pass
                if previous_timeout is not None:
                    try:
                        dev.timeout = previous_timeout
                    except Exception:
                        pass


        max_len = max(len(values[0]) for values in waveforms.values())
        with csv_path.open("w", newline="", encoding="utf-8") as handle:
            writer = csv.writer(handle)
            header: list[str] = []
            for channel in channels:
                header.extend([f"CH{channel}_Time_s", f"CH{channel}_Value_V"])
            writer.writerow(header)
            for index in range(max_len):
                row: list[float | str] = []
                for channel in channels:
                    times, volts = waveforms[channel]
                    row.extend([times[index], volts[index]] if index < len(times) else ["", ""])
                writer.writerow(row)
        return channels, max_len


class SimulationEnvironment:
    """Deterministic virtual bench shared by all simulation instruments."""

    def __init__(self, scenario: str = "Nominal"):
        self.scenario = scenario
        self.current_set = 0.0
        self.load_enabled = False
        self.vin = 48.0
        self.sample_index = 0
        self.channels = {
            1: {"voltage": 0.0, "current_limit": 1.0, "enabled": False},
            2: {"voltage": 0.0, "current_limit": 1.0, "enabled": False},
            3: {"voltage": 6.0, "current_limit": 1.0, "enabled": True},
        }

    @property
    def current(self) -> float:
        return self.current_set if self.load_enabled else 0.0

    def advance(self) -> None:
        self.sample_index += 1
        if self.scenario == "Required device failure" and self.sample_index >= 4:
            raise RequiredInstrumentError("Simulated required-device communication failure")


class SimInstrument:
    kind = "Simulation"

    def __init__(self, env: SimulationEnvironment, identity: str):
        self.env = env
        self.identity = identity
        self.connected = False
        self._persistent = False

    def connect(self, persistent: bool = True) -> str:
        self.connected = True
        self._persistent = persistent
        return self.identity

    def release(self) -> None:
        self.connected = False
        self._persistent = False


class SimLoad(SimInstrument):
    def read_snapshot(self, **_: Any) -> InstrumentSnapshot:
        self.env.advance()
        current = self.env.current
        voltage = max(0.0, 12.05 - 0.0012 * current)
        return InstrumentSnapshot("load", {"current": current, "voltage": voltage, "power": current * voltage})

    def set_current(self, amps: float) -> None:
        self.env.current_set = float(amps)

    def set_input(self, enabled: bool) -> None:
        self.env.load_enabled = bool(enabled)

    def safe_off(self) -> None:
        self.env.load_enabled = False


class SimPA(SimInstrument):
    def read_snapshot(self, **_: Any) -> InstrumentSnapshot:
        self.env.advance()
        current = self.env.current
        vout = max(0.0, 12.05 - 0.0012 * current)
        pout = vout * current
        efficiency = max(0.88, 0.982 - 0.00045 * current)
        pin = pout / efficiency if current > 0 else 2.0
        vin = self.env.vin
        iin = pin / vin
        if self.env.scenario == "Stale measurement":
            return InstrumentSnapshot(
                "pa",
                {"vin": vin, "iin": iin, "vout": vout},
                timestamp=(datetime.now(timezone.utc) - timedelta(seconds=30)).isoformat(timespec="milliseconds"),
                valid=False,
                warning="Simulated stale PA reading",
                status="Connected · Read Error",
            )
        return InstrumentSnapshot("pa", {"vin": vin, "iin": iin, "vout": vout}, valid=True, status="Connected")


class SimPSU(SimInstrument):
    def read_snapshot(self, channels: Iterable[int] = (1, 2, 3), **_: Any) -> InstrumentSnapshot:
        values: dict[str, Any] = {}
        for channel in channels:
            ch = int(channel)
            state = self.env.channels.get(ch, {"voltage": 0.0, "current_limit": 1.0, "enabled": False})
            voltage = state["voltage"] if state["enabled"] else 0.0
            current = (0.025 + 0.0007 * self.env.current) if state["enabled"] else 0.0
            values[f"ch{ch}_voltage"] = voltage
            values[f"ch{ch}_current"] = current
            values[f"ch{ch}_enabled"] = state["enabled"]
        return InstrumentSnapshot("psu", values)

    def configure_channel(self, channel: int, voltage: float, current_limit: float, enabled: bool) -> None:
        self.env.channels[int(channel)] = {
            "voltage": float(voltage), "current_limit": float(current_limit), "enabled": bool(enabled)
        }


def _png_chunk(kind: bytes, data: bytes) -> bytes:
    return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)


def make_demo_png(width: int = 640, height: int = 360) -> bytes:
    """Create a dependency-free, valid white-background demo PNG."""
    rows = []
    for y in range(height):
        row = bytearray([0])
        for x in range(width):
            trace = height // 2 + int(55 * math.sin(x / 28.0))
            if abs(y - trace) <= 1:
                row.extend((0, 74, 173))
            elif y % 60 == 0 or x % 80 == 0:
                row.extend((225, 229, 235))
            else:
                row.extend((255, 255, 255))
        rows.append(bytes(row))
    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    return b"\x89PNG\r\n\x1a\n" + _png_chunk(b"IHDR", ihdr) + _png_chunk(b"IDAT", zlib.compress(b"".join(rows), 6)) + _png_chunk(b"IEND", b"")


class SimScope(SimInstrument):
    def read_snapshot(self, **_: Any) -> InstrumentSnapshot:
        if self.env.scenario == "Missing optional device":
            raise InstrumentError("Simulated scope unavailable")
        return InstrumentSnapshot("scope", {"displayed_channels": "1,2,3,4"})

    def capture(self, png_path: Path, csv_path: Path) -> tuple[list[int], int]:
        if self.env.scenario in {"Missing optional device", "Scope capture failure"}:
            raise InstrumentError(f"Simulated: {self.env.scenario.lower()}")
        png_path.parent.mkdir(parents=True, exist_ok=True)
        png_path.write_bytes(make_demo_png())
        with csv_path.open("w", newline="", encoding="utf-8") as handle:
            writer = csv.writer(handle)
            writer.writerow([item for ch in range(1, 5) for item in (f"CH{ch}_Time_s", f"CH{ch}_Value_V")])
            for index in range(1000):
                t = (index - 500) * 1e-8
                row: list[float] = []
                for ch in range(1, 5):
                    row.extend((t, (4.0 / ch) * math.sin(2 * math.pi * 100_000 * t + ch)))
                writer.writerow(row)
        return [1, 2, 3, 4], 1000


class InstrumentHub:
    """Own the real or simulated bench and expose one consistent API."""

    def __init__(self, simulation: bool = False, config: dict[str, Any] | None = None):
        config = config or {}
        self.simulation = simulation
        self.manager: VisaManager | None = None
        if simulation:
            self.environment = SimulationEnvironment(config.get("simulation_scenario", "Nominal"))
            self.instruments = {
                "pa": SimPA(self.environment, "SIM,PA2201A,SIM0001,1.0"),
                "load": SimLoad(self.environment, "SIM,63206A-60-1000,SIM0002,1.0"),
                "psu": SimPSU(self.environment, "SIM,E36312A,SIM0003,1.0"),
                "scope": SimScope(self.environment, "SIM,MSOX4024A,SIM0004,1.0"),
            }
        else:
            self.environment = None
            self.manager = VisaManager(config.get("visa_library", VISA_DLL))
            addresses = {**KNOWN_ADDRESSES, **config.get("addresses", {})}
            self.instruments = {
                "pa": PA2201A(self.manager, addresses["pa"], IDENTITY_TOKENS["pa"]),
                "load": Chroma63206A(self.manager, addresses["load"], IDENTITY_TOKENS["load"]),
                "psu": E36312A(self.manager, addresses["psu"], IDENTITY_TOKENS["psu"]),
                "scope": MSOX4024A(self.manager, addresses["scope"], IDENTITY_TOKENS["scope"]),
            }

    def discover(self) -> dict[str, dict[str, str]]:
        if self.simulation:
            return {kind: {"address": "SIM", "identity": inst.identity} for kind, inst in self.instruments.items()}
        assert self.manager is not None
        found = self.manager.discover()
        for kind, item in found.items():
            self.instruments[kind].address = item["address"]
        return found

    @property
    def discovery_errors(self) -> dict[str, str]:
        return {} if self.manager is None else dict(self.manager.discovery_errors)

    def release_all(self) -> None:
        for instrument in self.instruments.values():
            try:
                instrument.release()
            except Exception:
                pass
        if self.manager is not None:
            self.manager.close()

    def safe_shutdown(self) -> None:
        try:
            self.instruments["load"].safe_off()
        finally:
            self.release_all()
