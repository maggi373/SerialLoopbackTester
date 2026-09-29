from __future__ import annotations

import argparse
import ctypes
from dataclasses import dataclass
import errno
import os
from pathlib import Path
import struct
import sys


MOXA_VENDOR_ID = 0x110A
SUPPORTED_PRODUCT_IDS = {
    0x1150: "UPort 1150",
    0x1151: "UPort 1150I",
}

MODE_CODES = {
    "rs232": 0,
    "rs485-2w": 1,
    "rs422": 2,
    "rs485-4w": 2,
}

MODE_ALIASES = {
    "0": "rs232",
    "rs232": "rs232",
    "rs-232": "rs232",
    "1": "rs485-2w",
    "rs485": "rs485-2w",
    "rs485-2w": "rs485-2w",
    "rs-485": "rs485-2w",
    "rs-485-2w": "rs485-2w",
    "2": "rs422",
    "rs422": "rs422",
    "rs-422": "rs422",
    "3": "rs485-4w",
    "rs485-4w": "rs485-4w",
    "rs-485-4w": "rs485-4w",
}

MODE_LABELS = {
    "rs232": "RS-232",
    "rs485-2w": "RS-485 two-wire",
    "rs422": "RS-422",
    "rs485-4w": "RS-485 four-wire",
}

TI_SET_CONFIG = 0x05
TI_UART1_PORT = 0x03
USB_VENDOR_DEVICE_OUT = 0x40
USB_CONTROL_TIMEOUT_MS = 1000

UART_ENABLE_RTS_IN = 0x0001
UART_ENABLE_PARITY_CHECKING = 0x0008
UART_ENABLE_CTS_OUT = 0x0020
UART_ENABLE_X_OUT = 0x0040
UART_ENABLE_X_IN = 0x0100
UART_ENABLE_MS_INTS = 0x2000
UART_ENABLE_AUTO_START_DMA = 0x4000


@dataclass(frozen=True)
class MoxaUsbDevice:
    tty_path: str
    usb_path: Path
    product_id: int
    product_name: str


class UsbdevfsCtrlTransfer(ctypes.Structure):
    _fields_ = (
        ("bRequestType", ctypes.c_ubyte),
        ("bRequest", ctypes.c_ubyte),
        ("wValue", ctypes.c_ushort),
        ("wIndex", ctypes.c_ushort),
        ("wLength", ctypes.c_ushort),
        ("timeout", ctypes.c_uint),
        ("data", ctypes.c_void_p),
    )


def normalize_mode(value: object) -> str:
    normalized = str(value if value is not None else "").strip().casefold()
    try:
        return MODE_ALIASES[normalized]
    except KeyError as exc:
        choices = ", ".join(MODE_CODES)
        raise ValueError(f"Unsupported Moxa interface mode '{value}'. Choose: {choices}.") from exc


def _read_hex(path: Path) -> int:
    return int(path.read_text(encoding="ascii").strip(), 16)


def find_moxa_uport(
    tty_path: object,
    *,
    sys_class_tty: Path = Path("/sys/class/tty"),
    usbfs_root: Path = Path("/dev/bus/usb"),
) -> MoxaUsbDevice:
    if not sys.platform.startswith("linux"):
        raise OSError("The Moxa UPort USB helper is available only on Linux.")

    endpoint = str(tty_path if tty_path is not None else "").strip()
    if not endpoint.startswith("/dev/") or "://" in endpoint:
        raise ValueError("The Moxa helper requires a local /dev/... serial device.")

    try:
        tty_name = Path(endpoint).resolve(strict=True).name
        sys_device = (sys_class_tty / tty_name / "device").resolve(strict=True)
    except OSError as exc:
        raise OSError(f"Cannot resolve Linux serial device {endpoint}: {exc}") from exc

    usb_device_path: Path | None = None
    vendor_id = product_id = None
    for candidate in (sys_device, *sys_device.parents):
        vendor_file = candidate / "idVendor"
        product_file = candidate / "idProduct"
        if not vendor_file.is_file() or not product_file.is_file():
            continue
        try:
            vendor_id = _read_hex(vendor_file)
            product_id = _read_hex(product_file)
        except (OSError, ValueError) as exc:
            raise OSError(f"Cannot read USB identity for {endpoint}: {exc}") from exc
        usb_device_path = candidate
        break

    if usb_device_path is None or vendor_id is None or product_id is None:
        raise OSError(f"Could not find the parent USB device for {endpoint} in sysfs.")
    if vendor_id != MOXA_VENDOR_ID or product_id not in SUPPORTED_PRODUCT_IDS:
        identity = f"{vendor_id:04x}:{product_id:04x}"
        raise ValueError(
            f"{endpoint} is USB device {identity}, not a supported Moxa UPort 1150/1150I."
        )

    try:
        bus_number = int((usb_device_path / "busnum").read_text(encoding="ascii").strip())
        device_number = int((usb_device_path / "devnum").read_text(encoding="ascii").strip())
    except (OSError, ValueError) as exc:
        raise OSError(f"Cannot locate the USB device node for {endpoint}: {exc}") from exc

    usb_path = usbfs_root / f"{bus_number:03d}" / f"{device_number:03d}"
    if not usb_path.exists():
        raise OSError(f"USB device node {usb_path} does not exist; reconnect the UPort and try again.")

    return MoxaUsbDevice(
        tty_path=endpoint,
        usb_path=usb_path,
        product_id=product_id,
        product_name=SUPPORTED_PRODUCT_IDS[product_id],
    )


def build_uart_config(
    mode: object,
    *,
    baudrate: int,
    bytesize: int,
    parity: str,
    stopbits: float,
    xonxoff: bool = False,
    rtscts: bool = False,
) -> bytes:
    normalized_mode = normalize_mode(mode)
    baud = int(baudrate)
    if baud <= 0:
        raise ValueError("Baudrate must be greater than zero.")
    if baud > 923077:
        raise ValueError("Baudrate is too high for the UPort 1150/1150I divisor.")

    data_bits = int(bytesize)
    if data_bits not in {5, 6, 7, 8}:
        raise ValueError("Bytesize must be 5, 6, 7, or 8.")

    parity_name = str(parity).strip().upper()
    parity_codes = {"N": 0, "O": 1, "E": 2, "M": 3, "S": 4}
    if parity_name not in parity_codes:
        raise ValueError("Parity must be N, O, E, M, or S.")

    stop_bits = float(stopbits)
    if stop_bits == 1.0:
        stop_code = 0
    elif stop_bits == 1.5:
        stop_code = 1
    elif stop_bits == 2.0:
        stop_code = 2
    else:
        raise ValueError("Stop bits must be 1, 1.5, or 2.")

    flags = UART_ENABLE_MS_INTS | UART_ENABLE_AUTO_START_DMA
    if parity_name != "N":
        flags |= UART_ENABLE_PARITY_CHECKING
    if rtscts:
        flags |= UART_ENABLE_RTS_IN | UART_ENABLE_CTS_OUT
    if xonxoff:
        flags |= UART_ENABLE_X_IN | UART_ENABLE_X_OUT

    # Match Moxa's mxu11x0 driver exactly. In particular, it truncates the
    # divisor and leaves XON/XOFF bytes zero unless software flow control is
    # enabled.
    divisor = 923077 // baud
    return struct.pack(
        ">HHBBBBBB",
        divisor,
        flags,
        data_bits - 5,
        parity_codes[parity_name],
        stop_code,
        0x11 if xonxoff else 0,
        0x13 if xonxoff else 0,
        MODE_CODES[normalized_mode],
    )


def _iowr(type_number: int, command_number: int, size: int) -> int:
    # Linux generic _IOWR layout used by Fedora's supported architectures.
    return (3 << 30) | (size << 16) | (type_number << 8) | command_number


USBDEVFS_CONTROL = _iowr(ord("U"), 0, ctypes.sizeof(UsbdevfsCtrlTransfer))


def _usb_control_transfer(
    usb_path: Path,
    payload: bytes,
) -> None:
    buffer = (ctypes.c_ubyte * len(payload)).from_buffer_copy(payload)
    transfer = UsbdevfsCtrlTransfer(
        bRequestType=USB_VENDOR_DEVICE_OUT,
        bRequest=TI_SET_CONFIG,
        wValue=0,
        wIndex=TI_UART1_PORT,
        wLength=len(payload),
        timeout=USB_CONTROL_TIMEOUT_MS,
        data=ctypes.cast(buffer, ctypes.c_void_p),
    )

    flags = os.O_RDWR | getattr(os, "O_CLOEXEC", 0)
    try:
        descriptor = os.open(usb_path, flags)
    except PermissionError as exc:
        raise PermissionError(
            f"Permission denied opening {usb_path}. Run install_serial_access.sh once, reconnect the adapter, "
            "and log out/in; or run this command as root."
        ) from exc

    try:
        libc = ctypes.CDLL(None, use_errno=True)
        libc.ioctl.argtypes = (ctypes.c_int, ctypes.c_ulong, ctypes.c_void_p)
        libc.ioctl.restype = ctypes.c_int
        result = libc.ioctl(descriptor, USBDEVFS_CONTROL, ctypes.byref(transfer))
        if result < 0:
            error_number = ctypes.get_errno()
            if error_number in {errno.EACCES, errno.EPERM}:
                raise PermissionError(
                    error_number,
                    "Moxa USB control access was denied. Run install_serial_access.sh once, reconnect the "
                    "adapter, and log out/in; or run this command as root.",
                    str(usb_path),
                )
            raise OSError(error_number, os.strerror(error_number), str(usb_path))
        if result != len(payload):
            raise OSError(
                f"Moxa SET_CONFIG transferred {result} of {len(payload)} bytes through {usb_path}."
            )
    finally:
        os.close(descriptor)


def apply_moxa_uport_mode(
    tty_path: object,
    mode: object,
    *,
    baudrate: int,
    bytesize: int,
    parity: str,
    stopbits: float,
    xonxoff: bool = False,
    rtscts: bool = False,
) -> MoxaUsbDevice:
    device = find_moxa_uport(tty_path)
    payload = build_uart_config(
        mode,
        baudrate=baudrate,
        bytesize=bytesize,
        parity=parity,
        stopbits=stopbits,
        xonxoff=xonxoff,
        rtscts=rtscts,
    )
    _usb_control_transfer(device.usb_path, payload)
    return device


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Set the electrical interface on a Moxa UPort 1150/1150I through Linux's built-in "
            "ti_usb_3410_5052 driver. No Moxa kernel module is required."
        )
    )
    parser.add_argument("device", help="Local serial device, for example /dev/ttyUSB0")
    parser.add_argument(
        "mode",
        metavar="MODE",
        help="rs232, rs485-2w, rs422, or rs485-4w (numeric aliases 0..3 are also accepted)",
    )
    parser.add_argument("--baudrate", type=int, default=19200)
    parser.add_argument("--bytesize", type=int, choices=(5, 6, 7, 8), default=8)
    parser.add_argument("--parity", choices=("N", "E", "O", "M", "S", "n", "e", "o", "m", "s"), default="N")
    parser.add_argument("--stopbits", type=float, choices=(1.0, 1.5, 2.0), default=1.0)
    parser.add_argument("--xonxoff", action="store_true")
    parser.add_argument("--rtscts", action="store_true")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_argument_parser().parse_args(argv)
    try:
        device = apply_moxa_uport_mode(
            args.device,
            args.mode,
            baudrate=args.baudrate,
            bytesize=args.bytesize,
            parity=args.parity,
            stopbits=args.stopbits,
            xonxoff=args.xonxoff,
            rtscts=args.rtscts,
        )
    except (OSError, ValueError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    print(
        f"Applied {MODE_LABELS[normalize_mode(args.mode)]} to {device.product_name} "
        f"on {device.tty_path} ({device.usb_path})."
    )
    print("Reapply after opening the TTY or changing baud/parity because the kernel driver resets its configuration.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
