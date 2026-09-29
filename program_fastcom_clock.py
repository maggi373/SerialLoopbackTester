#!/usr/bin/env python3
"""Restricted Fastcom PCI-335 clock programmer used by the Linux GUI."""

from __future__ import annotations

import mmap
import os
from pathlib import Path
import re
import sys
import time


FASTCOM_VENDOR = "0x18f7"
SUPPORTED_DEVICES = {"0x000a", "0x000b"}
EXPECTED_DRIVER = "exar_serial"
MPIO_LEVEL_OFFSET = 0x90
FASTCOM_29MHZ_WORD = 0x100801
MPIO_DATA = 0x01
MPIO_CLOCK = 0x02
MPIO_STROBE = 0x04
PCI_COMMAND_OFFSET = 0x04
PCI_COMMAND_MEMORY = 0x0002


def fail(message: str) -> "None":
    raise SystemExit(message)


def read_lower(path: Path) -> str:
    return path.read_text(encoding="ascii").strip().lower()


def find_fastcom_pci_device(tty_name: str) -> Path:
    device_link = Path("/sys/class/tty") / tty_name / "device"
    if not device_link.exists():
        fail(f"Linux sysfs has no device for /dev/{tty_name}.")
    cursor = device_link.resolve()
    for candidate in (cursor, *cursor.parents):
        vendor_path = candidate / "vendor"
        device_path = candidate / "device"
        if not vendor_path.is_file() or not device_path.is_file():
            continue
        vendor = read_lower(vendor_path)
        device = read_lower(device_path)
        if vendor == FASTCOM_VENDOR and device in SUPPORTED_DEVICES:
            return candidate
        fail(f"/dev/{tty_name} belongs to unsupported PCI device {vendor}:{device}.")
    fail(f"/dev/{tty_name} is not below a PCI device in sysfs.")


def find_all_fastcom_pci_devices() -> list[Path]:
    pci_root = Path("/sys/bus/pci/devices")
    cards: list[Path] = []
    try:
        candidates = sorted(pci_root.iterdir(), key=lambda path: path.name)
    except (FileNotFoundError, PermissionError, OSError) as exc:
        fail(f"Could not enumerate PCI devices below {pci_root}: {exc}")
    for candidate in candidates:
        vendor_path = candidate / "vendor"
        device_path = candidate / "device"
        try:
            if read_lower(vendor_path) == FASTCOM_VENDOR and read_lower(device_path) in SUPPORTED_DEVICES:
                cards.append(candidate.resolve())
        except (FileNotFoundError, PermissionError, OSError):
            continue
    if not cards:
        fail("No supported Fastcom PCI-335 cards (18f7:000a/000b) were detected.")
    return cards


def card_tty_names(pci_device: Path) -> set[str]:
    names: set[str] = set()
    for candidate in pci_device.rglob("ttyS*"):
        if re.fullmatch(r"ttyS\d+", candidate.name):
            names.add(candidate.name)
    return names


def open_users(tty_names: set[str]) -> list[str]:
    device_targets = {f"/dev/{name}" for name in tty_names}
    users: list[str] = []
    proc = Path("/proc")
    for process in proc.iterdir():
        if not process.name.isdigit():
            continue
        fd_dir = process / "fd"
        try:
            descriptors = tuple(fd_dir.iterdir())
        except (FileNotFoundError, PermissionError):
            continue
        for descriptor in descriptors:
            try:
                target = os.readlink(descriptor)
            except (FileNotFoundError, PermissionError, OSError):
                continue
            if target in device_targets:
                users.append(f"PID {process.name} fd {descriptor.name} -> {target}")
    return users


def ensure_ports_idle(tty_names: set[str]) -> None:
    try:
        active_consoles = set(Path("/sys/class/tty/console/active").read_text().split())
    except (FileNotFoundError, PermissionError):
        active_consoles = set()
    console_ports = sorted(tty_names & active_consoles)
    if console_ports:
        fail(f"Refusing to unbind a card containing active console port(s): {', '.join(console_ports)}")
    users = open_users(tty_names)
    if users:
        fail("Close every port on this Fastcom card first: " + "; ".join(users[:8]))


def write_mmio_byte(mapping: mmap.mmap, value: int) -> None:
    mapping[MPIO_LEVEL_OFFSET] = value & 0xFF
    _ = mapping[MPIO_LEVEL_OFFSET]


def clock_programming_values(saved: int) -> list[int]:
    """Return the byte writes used by Fastcom's official 24-bit MPIO sequence."""
    values: list[int] = []
    data = saved
    clock_word = FASTCOM_29MHZ_WORD & 0xFFFFFF
    for _bit in range(24):
        if clock_word & 0x800000:
            data |= MPIO_DATA
        else:
            data &= ~MPIO_DATA
        values.append(data)
        data |= MPIO_CLOCK
        values.append(data)
        data &= MPIO_DATA
        values.append(data)
        clock_word = (clock_word << 1) & 0xFFFFFF

    data &= 0xF8
    data |= MPIO_STROBE
    values.append(data)
    data &= ~MPIO_STROBE
    values.append(data)
    values.append(saved)
    return values


def program_clock(resource_path: Path) -> None:
    resource_size = resource_path.stat().st_size
    if resource_size and resource_size <= MPIO_LEVEL_OFFSET:
        fail(f"BAR0 is only {resource_size} bytes; MPIO offset 0x{MPIO_LEVEL_OFFSET:x} is unavailable.")
    mapping_length = min(resource_size, mmap.PAGESIZE) if resource_size else mmap.PAGESIZE
    mapping_length = max(mapping_length, MPIO_LEVEL_OFFSET + 1)
    descriptor = os.open(resource_path, os.O_RDWR | os.O_SYNC)
    try:
        with mmap.mmap(
            descriptor,
            mapping_length,
            flags=mmap.MAP_SHARED,
            prot=mmap.PROT_READ | mmap.PROT_WRITE,
        ) as mapping:
            saved = mapping[MPIO_LEVEL_OFFSET]
            for value in clock_programming_values(saved):
                write_mmio_byte(mapping, value)
    finally:
        os.close(descriptor)


def set_pci_memory_decode(config_path: Path, enabled: bool, original: int | None = None) -> int:
    """Enable BAR memory decoding while unbound, or restore the saved PCI command."""
    descriptor = os.open(config_path, os.O_RDWR | os.O_SYNC)
    try:
        command_bytes = os.pread(descriptor, 2, PCI_COMMAND_OFFSET)
        if len(command_bytes) != 2:
            fail(f"Could not read the PCI command register from {config_path}.")
        current = int.from_bytes(command_bytes, "little")
        desired = (current | PCI_COMMAND_MEMORY) if enabled else original
        if desired is None:
            fail("The original PCI command register value was not supplied.")
        if current != desired:
            written = os.pwrite(descriptor, int(desired).to_bytes(2, "little"), PCI_COMMAND_OFFSET)
            if written != 2:
                fail(f"Could not update the PCI command register in {config_path}.")
        return current
    finally:
        os.close(descriptor)


def program_clock_while_unbound(pci_device: Path) -> None:
    config_path = pci_device / "config"
    original_command = set_pci_memory_decode(config_path, True)
    try:
        program_clock(pci_device / "resource0")
    finally:
        set_pci_memory_decode(config_path, False, original_command)


def bind(driver_path: Path, pci_address: str) -> None:
    (driver_path / "bind").write_text(pci_address, encoding="ascii")


def prepare_card(pci_device: Path, required_tty_name: str | None = None) -> tuple[Path, set[str]]:
    pci_address = pci_device.name
    vendor = read_lower(pci_device / "vendor")
    device = read_lower(pci_device / "device")
    if vendor != FASTCOM_VENDOR or device not in SUPPORTED_DEVICES:
        fail(f"Refusing unsupported PCI device {vendor}:{device} at {pci_address}.")
    driver_link = pci_device / "driver"
    if not driver_link.exists():
        fail(f"Fastcom card {pci_address} is not bound to a driver.")
    driver_path = driver_link.resolve()
    if driver_path.name != EXPECTED_DRIVER:
        fail(
            f"Fastcom card {pci_address} uses {driver_path.name}, not {EXPECTED_DRIVER}; "
            "the direct 8250_exar procedure was not run."
        )

    tty_names = card_tty_names(pci_device)
    if not tty_names:
        fail(f"Fastcom card {pci_address} has no ttyS ports below it.")
    if required_tty_name is not None and required_tty_name not in tty_names:
        fail(f"Could not associate /dev/{required_tty_name} with Fastcom card {pci_address}.")
    ensure_ports_idle(tty_names)
    return driver_path, tty_names


def program_card(pci_device: Path, driver_path: Path, tty_names: set[str]) -> str:
    pci_address = pci_device.name
    unbound = False
    primary_error: BaseException | None = None
    try:
        (driver_path / "unbind").write_text(pci_address, encoding="ascii")
        unbound = True
        program_clock_while_unbound(pci_device)
    except BaseException as exc:
        primary_error = exc
    finally:
        if unbound:
            try:
                bind(driver_path, pci_address)
            except BaseException as bind_error:
                if primary_error is not None:
                    fail(f"{primary_error}; additionally failed to rebind {pci_address}: {bind_error}")
                fail(
                    f"Clock word was sent, but {pci_address} failed to rebind: {bind_error}. "
                    "Reboot to restore the ports."
                )

    if primary_error is not None:
        fail(f"Clock programming failed for {pci_address}; the driver was rebound: {primary_error}")

    expected_ttys = sorted(tty_names)
    deadline = time.monotonic() + 5.0
    while time.monotonic() < deadline:
        returned = card_tty_names(pci_device)
        if len(returned) >= len(expected_ttys):
            break
        time.sleep(0.05)

    return (
        f"Programmed Fastcom {read_lower(pci_device / 'device')} at {pci_address} "
        f"to 29.4912 MHz with word 0x{FASTCOM_29MHZ_WORD:06X}; "
        f"rebound {EXPECTED_DRIVER} ({', '.join(expected_ttys)})."
    )


def main() -> None:
    if os.geteuid() != 0:
        fail("Fastcom clock programming requires root privileges.")
    if len(sys.argv) != 2 or not (
        sys.argv[1] == "--all" or re.fullmatch(r"/dev/ttyS\d+", sys.argv[1])
    ):
        fail(f"Usage: {sys.argv[0]} --all | /dev/ttyS<N>")

    required_tty_names: dict[Path, str | None] = {}
    if sys.argv[1] == "--all":
        for pci_device in find_all_fastcom_pci_devices():
            required_tty_names[pci_device] = None
    else:
        tty_path = Path(sys.argv[1])
        if not tty_path.exists() or not tty_path.is_char_device():
            fail(f"{tty_path} is not a character device.")
        pci_device = find_fastcom_pci_device(tty_path.name)
        required_tty_names[pci_device] = tty_path.name

    # Validate every card and every port before changing the first card.
    prepared = [
        (pci_device, *prepare_card(pci_device, required_tty_name))
        for pci_device, required_tty_name in required_tty_names.items()
    ]
    results = [
        program_card(pci_device, driver_path, tty_names)
        for pci_device, driver_path, tty_names in prepared
    ]
    print("\n".join(results))
    if len(results) > 1:
        print(f"Programmed all {len(results)} detected Fastcom PCI-335 cards.")


if __name__ == "__main__":
    main()
