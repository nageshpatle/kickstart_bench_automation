"""Dependency bootstrap and passive SID bench confidence checks."""

from __future__ import annotations

import argparse
import importlib.util
import os
import subprocess
import sys
from pathlib import Path


KEYSIGHT_BIN = Path(r"C:\Program Files\Keysight\IO Libraries Suite\bin")
VISA_DLL = Path(r"C:\Windows\System32\visa64.dll")
PACKAGES = {"PyQt6": "PyQt6", "pyqtgraph": "pyqtgraph", "openpyxl": "openpyxl", "pyvisa": "pyvisa", "numpy": "numpy", "pytest": "pytest"}


def missing_packages() -> list[str]:
    return [package for module, package in PACKAGES.items() if importlib.util.find_spec(module) is None]


def install_missing() -> bool:
    missing = missing_packages()
    if not missing:
        print("Python dependencies: ready")
        return True
    print("Installing missing Python packages:", ", ".join(missing))
    result = subprocess.run([sys.executable, "-m", "pip", "install", "-r", str(Path(__file__).with_name("requirements.txt"))], check=False)
    if result.returncode:
        print("Package installation failed. Resolve the pip error above and retry.")
        return False
    print("Dependencies installed. Restart this command or launch sid_bench_gui.py.")
    return True


def prepare_keysight_path() -> None:
    if KEYSIGHT_BIN.is_dir():
        paths = os.environ.get("PATH", "").split(os.pathsep)
        if str(KEYSIGHT_BIN).lower() not in {path.lower() for path in paths}:
            os.environ["PATH"] = str(KEYSIGHT_BIN) + os.pathsep + os.environ.get("PATH", "")


def print_driver_status() -> bool:
    ready = VISA_DLL.is_file() and KEYSIGHT_BIN.is_dir()
    print(f"Keysight IO Libraries: {'ready' if ready else 'not detected'}")
    print(f"  VISA DLL: {VISA_DLL}")
    print(f"  IO bin:   {KEYSIGHT_BIN}")
    if not ready:
        print("Install or repair Keysight IO Libraries Suite / Connection Expert 2026, then retry.")
        print("Do not reinstall drivers merely because one VISA session reports VI_ERROR_NCIC.")
    return ready


def passive_discovery() -> int:
    from sid_instruments import InstrumentHub
    print("\nPASSIVE VISA DISCOVERY")
    print("Each resource receives one *IDN? query; every discovery session is then closed.")
    hub = InstrumentHub(False)
    try:
        found = hub.discover()
        if not found:
            print("No recognized SID bench instruments found.")
            return 1
        for kind, item in found.items():
            print(f"{kind:>6}: {item['identity']}")
            print(f"        {item['address']}")
        missing = [name for name in ("pa", "load", "psu", "scope") if name not in found]
        if missing:
            print("Optional/offline or unresolved:", ", ".join(missing))
        if hub.discovery_errors:
            print("\nResources that did not answer *IDN?:")
            for address, error in hub.discovery_errors.items():
                print(f"  {address}: {error}")
            if any("VI_ERROR_NCIC" in error for error in hub.discovery_errors.values()):
                print("  Action: release Interactive IO/other VISA clients, then retry. Do not reinstall drivers first.")
        return 0
    except Exception as exc:
        print("Discovery failed:", exc)
        return 1
    finally:
        hub.release_all()


def simulation_confidence() -> int:
    from sid_instruments import InstrumentHub
    print("\nSIMULATION CONFIDENCE CHECK (no VISA calls)")
    hub = InstrumentHub(True)
    try:
        for kind, instrument in hub.instruments.items():
            identity = instrument.connect(persistent=False)
            snapshot = instrument.read_snapshot()
            instrument.release()
            print(f"{kind:>6}: {identity} -> {snapshot.values}")
        print("Simulation confidence check passed.")
        return 0
    finally:
        hub.release_all()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--install", action="store_true", help="Install missing Python packages")
    parser.add_argument("--simulation", action="store_true", help="Run no-hardware confidence checks")
    parser.add_argument("--discover", action="store_true", help="Passively discover real VISA instruments")
    parser.add_argument("--launch", action="store_true", help="Launch the GUI after checks")
    args = parser.parse_args()
    if args.install and not install_missing():
        return 1
    missing = missing_packages()
    if missing:
        print("Missing Python packages:", ", ".join(missing))
        print("Run: python bench_test.py --install")
        return 1
    prepare_keysight_path()
    print_driver_status()
    result = 0
    if args.simulation or (not args.discover and not args.launch):
        result = simulation_confidence()
    if args.discover:
        result = max(result, passive_discovery()) if print_driver_status() else 1
    if args.launch and result == 0:
        from sid_bench_gui import main as gui_main
        return gui_main()
    return result


if __name__ == "__main__":
    raise SystemExit(main())
