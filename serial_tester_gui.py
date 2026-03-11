from __future__ import annotations

from collections import deque
import json
import os
import queue
import sys
import threading
import time
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk

try:
    import serial
    from serial import SerialException
    from serial.tools import list_ports
except ImportError as exc:  # pragma: no cover
    raise SystemExit("Missing dependency: pyserial. Install with `pip install pyserial`.") from exc

DEFAULT_RS232_COUNT = 40
DEFAULT_RS485_PAIR_COUNT = 8
DEFAULT_PRESET_COUNT = 5
FAILURE_WINDOW_SECONDS = 3600
FAILURE_WINDOW_LABEL = "1h"
APP_VERSION = "1.0.0"
APP_PUBLISHER = "PoldenTEK"
SETTINGS_FILENAME = "serial_tester_settings.json"
PARITY_OPTIONS = ("N", "E", "O", "M", "S")
BYTESIZE_OPTIONS = ("5", "6", "7", "8")
STOPBITS_OPTIONS = ("1", "1.5", "2")
APP_FOLDER_NAME = "SerialLoopbackTester"
STOPBITS_ALLOWED_VALUES = (1.0, 1.5, 2.0)


def as_bool(value: object, default: bool) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    if isinstance(value, str):
        lowered = value.strip().lower()
        if lowered in {"1", "true", "yes", "on", "enabled"}:
            return True
        if lowered in {"0", "false", "no", "off", "disabled"}:
            return False
    return default


def as_int(value: object, default: int) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def as_float(value: object, default: float) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def sanitize_hex_payload(payload: object, default: str) -> str:
    payload_text = str(payload if payload is not None else "").strip()
    cleaned = "".join(payload_text.split()).upper()
    if not cleaned:
        cleaned = default
    if len(cleaned) % 2 != 0:
        return default
    try:
        bytes.fromhex(cleaned)
    except ValueError:
        return default
    return cleaned


def validate_hex_payload(payload: str) -> str:
    cleaned = "".join(payload.strip().split()).upper()
    if not cleaned:
        raise ValueError("Payload cannot be empty.")
    if len(cleaned) % 2 != 0:
        raise ValueError("Payload hex must use an even number of characters.")
    try:
        bytes.fromhex(cleaned)
    except ValueError as exc:
        raise ValueError("Payload must be valid hexadecimal bytes.") from exc
    return cleaned


def parse_stopbits(value: object, default: float | None = None) -> float:
    text = str(value if value is not None else "").strip()
    if not text:
        if default is not None:
            return default
        raise ValueError("Stopbits must be 1, 1.5, or 2.")

    if text in STOPBITS_OPTIONS:
        return float(text)

    try:
        numeric = float(text)
    except (TypeError, ValueError):
        if default is not None:
            return default
        raise ValueError("Stopbits must be 1, 1.5, or 2.")

    if numeric in STOPBITS_ALLOWED_VALUES:
        return numeric

    if default is not None:
        return default
    raise ValueError("Stopbits must be 1, 1.5, or 2.")


def stopbits_to_text(value: object) -> str:
    numeric = parse_stopbits(value, default=1.0)
    if numeric == 1.0:
        return "1"
    if numeric == 2.0:
        return "2"
    return "1.5"


def normalize_port_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip().upper()


def normalize_port_list(values: object) -> list[str]:
    if isinstance(values, list):
        source_values = values
    elif isinstance(values, str):
        source_values = values.split(",")
    else:
        source_values = []

    normalized: list[str] = []
    seen: set[str] = set()
    for value in source_values:
        port = normalize_port_text(value)
        if port and port not in seen:
            seen.add(port)
            normalized.append(port)
    return normalized


def default_rs232_item(index: int) -> dict:
    number = index + 1
    return {
        "enabled": True,
        "name": f"RS232 {number}",
        "port": f"COM{number}",
        "baudrate": 9600,
        "bytesize": 8,
        "parity": "N",
        "stopbits": 1,
        "timeout_s": 0.5,
        "payload_hex": "55AA",
        "interval_ms": 100,
    }


def default_rs485_item(index: int) -> dict:
    number = index + 1
    sender = 41 + (index * 2)
    echo = sender + 1
    return {
        "enabled": True,
        "name": f"RS485 Pair {number}",
        "sender_port": f"COM{sender}",
        "echo_port": f"COM{echo}",
        "baudrate": 9600,
        "bytesize": 8,
        "parity": "N",
        "stopbits": 1,
        "timeout_s": 0.5,
        "payload_hex": "A55A",
        "interval_ms": 100,
    }


def normalize_rs232(item: object, index: int) -> dict:
    base = default_rs232_item(index)
    source = item if isinstance(item, dict) else {}

    parity = str(source.get("parity", base["parity"])).upper()
    if parity not in PARITY_OPTIONS:
        parity = base["parity"]

    bytesize = as_int(source.get("bytesize", base["bytesize"]), base["bytesize"])
    if str(bytesize) not in BYTESIZE_OPTIONS:
        bytesize = base["bytesize"]

    stopbits = parse_stopbits(source.get("stopbits", base["stopbits"]), default=float(base["stopbits"]))

    baudrate = as_int(source.get("baudrate", base["baudrate"]), base["baudrate"])
    interval_ms = max(as_int(source.get("interval_ms", base["interval_ms"]), base["interval_ms"]), 50)
    timeout_s = max(as_float(source.get("timeout_s", base["timeout_s"]), base["timeout_s"]), 0.05)

    name = str(source.get("name", base["name"])) or base["name"]
    if "port" in source:
        raw_port = source.get("port")
        port = "" if raw_port is None else str(raw_port)
    else:
        port = base["port"]
    enabled = as_bool(source.get("enabled", base["enabled"]), base["enabled"])
    payload_hex = sanitize_hex_payload(source.get("payload_hex"), base["payload_hex"])

    return {
        "enabled": enabled,
        "name": name.strip(),
        "port": port.strip(),
        "baudrate": baudrate,
        "bytesize": bytesize,
        "parity": parity,
        "stopbits": stopbits,
        "timeout_s": timeout_s,
        "payload_hex": payload_hex,
        "interval_ms": interval_ms,
    }


def normalize_rs485(item: object, index: int) -> dict:
    base = default_rs485_item(index)
    source = item if isinstance(item, dict) else {}

    parity = str(source.get("parity", base["parity"])).upper()
    if parity not in PARITY_OPTIONS:
        parity = base["parity"]

    bytesize = as_int(source.get("bytesize", base["bytesize"]), base["bytesize"])
    if str(bytesize) not in BYTESIZE_OPTIONS:
        bytesize = base["bytesize"]

    stopbits = parse_stopbits(source.get("stopbits", base["stopbits"]), default=float(base["stopbits"]))

    baudrate = as_int(source.get("baudrate", base["baudrate"]), base["baudrate"])
    interval_ms = max(as_int(source.get("interval_ms", base["interval_ms"]), base["interval_ms"]), 50)
    timeout_s = max(as_float(source.get("timeout_s", base["timeout_s"]), base["timeout_s"]), 0.05)

    name = str(source.get("name", base["name"])) or base["name"]
    if "sender_port" in source:
        raw_sender = source.get("sender_port")
        sender_port = "" if raw_sender is None else str(raw_sender)
    else:
        sender_port = base["sender_port"]
    if "echo_port" in source:
        raw_echo = source.get("echo_port")
        echo_port = "" if raw_echo is None else str(raw_echo)
    else:
        echo_port = base["echo_port"]
    enabled = as_bool(source.get("enabled", base["enabled"]), base["enabled"])
    payload_hex = sanitize_hex_payload(source.get("payload_hex"), base["payload_hex"])

    return {
        "enabled": enabled,
        "name": name.strip(),
        "sender_port": sender_port.strip(),
        "echo_port": echo_port.strip(),
        "baudrate": baudrate,
        "bytesize": bytesize,
        "parity": parity,
        "stopbits": stopbits,
        "timeout_s": timeout_s,
        "payload_hex": payload_hex,
        "interval_ms": interval_ms,
    }


def default_ui_settings() -> dict:
    return {
        "start_fullscreen": False,
        "auto_start_after_launch_2s": True,
        "delay_comm_start_2s": True,
        "overview_compact_view": True,
        "presets": [default_preset_item(i) for i in range(DEFAULT_PRESET_COUNT)],
    }


def default_preset_item(index: int) -> dict:
    return {
        "name": f"Preset {index + 1}",
        "ports": [],
    }


def normalize_preset_item(item: object, index: int) -> dict:
    base = default_preset_item(index)
    source = item if isinstance(item, dict) else {}
    name = str(source.get("name", base["name"])).strip() or base["name"]
    ports = normalize_port_list(source.get("ports", base["ports"]))
    return {
        "name": name,
        "ports": ports,
    }


def normalize_ui_settings(item: object) -> dict:
    base = default_ui_settings()
    source = item if isinstance(item, dict) else {}
    raw_presets = source.get("presets", base["presets"])
    if not isinstance(raw_presets, list):
        raw_presets = []
    presets = [normalize_preset_item(raw_presets[i] if i < len(raw_presets) else {}, i) for i in range(DEFAULT_PRESET_COUNT)]
    return {
        "start_fullscreen": as_bool(source.get("start_fullscreen", base["start_fullscreen"]), base["start_fullscreen"]),
        "auto_start_after_launch_2s": as_bool(
            source.get("auto_start_after_launch_2s", base["auto_start_after_launch_2s"]),
            base["auto_start_after_launch_2s"],
        ),
        "delay_comm_start_2s": as_bool(source.get("delay_comm_start_2s", base["delay_comm_start_2s"]), base["delay_comm_start_2s"]),
        "overview_compact_view": as_bool(
            source.get("overview_compact_view", base["overview_compact_view"]),
            base["overview_compact_view"],
        ),
        "presets": presets,
    }


def default_settings() -> dict:
    return {
        "rs232_ports": [default_rs232_item(i) for i in range(DEFAULT_RS232_COUNT)],
        "rs485_pairs": [default_rs485_item(i) for i in range(DEFAULT_RS485_PAIR_COUNT)],
        "ui": default_ui_settings(),
    }


def normalize_settings(raw: object) -> dict:
    source = raw if isinstance(raw, dict) else {}
    raw_rs232 = source.get("rs232_ports", [])
    raw_rs485 = source.get("rs485_pairs", [])
    raw_ui = source.get("ui", {})

    if not isinstance(raw_rs232, list):
        raw_rs232 = []
    if not isinstance(raw_rs485, list):
        raw_rs485 = []

    rs232_ports = [normalize_rs232(raw_rs232[i] if i < len(raw_rs232) else {}, i) for i in range(DEFAULT_RS232_COUNT)]
    rs485_pairs = [normalize_rs485(raw_rs485[i] if i < len(raw_rs485) else {}, i) for i in range(DEFAULT_RS485_PAIR_COUNT)]
    ui = normalize_ui_settings(raw_ui)

    return {"rs232_ports": rs232_ports, "rs485_pairs": rs485_pairs, "ui": ui}


def load_settings_file(path: Path) -> dict:
    if not path.exists():
        settings = default_settings()
        path.write_text(json.dumps(settings, indent=2), encoding="utf-8")
        return settings

    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        settings = default_settings()
        path.write_text(json.dumps(settings, indent=2), encoding="utf-8")
        return settings

    settings = normalize_settings(raw)
    path.write_text(json.dumps(settings, indent=2), encoding="utf-8")
    return settings


def save_settings_file(path: Path, settings: dict) -> None:
    path.write_text(json.dumps(settings, indent=2), encoding="utf-8")


def resolve_settings_path() -> Path:
    if getattr(sys, "frozen", False):
        base = Path(os.environ.get("APPDATA", str(Path.home()))) / APP_FOLDER_NAME
    else:
        base = Path(__file__).resolve().parent
    base.mkdir(parents=True, exist_ok=True)
    return base / SETTINGS_FILENAME


def read_exact(port: serial.Serial, length: int, timeout_s: float, stop_event: threading.Event) -> bytes:
    deadline = time.monotonic() + timeout_s
    data = bytearray()

    while len(data) < length and not stop_event.is_set():
        if time.monotonic() >= deadline:
            break
        chunk = port.read(length - len(data))
        if chunk:
            data.extend(chunk)
        else:
            time.sleep(0.01)

    return bytes(data)


class RS232Worker(threading.Thread):
    def __init__(self, index: int, config: dict, event_queue: queue.Queue):
        super().__init__(daemon=True)
        self.index = index
        self.config = config
        self.event_queue = event_queue
        self.stop_event = threading.Event()

    def stop(self) -> None:
        self.stop_event.set()

    def emit(
        self,
        status: str,
        last: str,
        pass_inc: int = 0,
        fail_inc: int = 0,
        log: bool = False,
    ) -> None:
        self.event_queue.put(
            {
                "group": "rs232",
                "index": self.index,
                "status": status,
                "last": last,
                "pass_inc": pass_inc,
                "fail_inc": fail_inc,
                "log": log,
            }
        )

    def open_port(self) -> serial.Serial:
        return serial.Serial(
            port=self.config["port"],
            baudrate=int(self.config["baudrate"]),
            bytesize=int(self.config["bytesize"]),
            parity=str(self.config["parity"]).upper(),
            stopbits=float(self.config["stopbits"]),
            timeout=float(self.config["timeout_s"]),
            write_timeout=float(self.config["timeout_s"]),
        )

    def run(self) -> None:
        payload = bytes.fromhex(self.config["payload_hex"])
        timeout_s = max(float(self.config["timeout_s"]), 0.05)
        interval_s = max(int(self.config["interval_ms"]) / 1000.0, 0.05)
        startup_delay_s = max(as_float(self.config.get("startup_delay_s", 0.0), 0.0), 0.0)
        payload_hex = payload.hex(" ").upper()

        while not self.stop_event.is_set():
            try:
                with self.open_port() as port:
                    self.emit("Running", "Port open", log=True)
                    if startup_delay_s > 0:
                        self.emit("Standby", f"Startup delay {startup_delay_s:.1f}s")
                        if self.stop_event.wait(startup_delay_s):
                            break

                    while not self.stop_event.is_set():
                        port.reset_input_buffer()
                        port.reset_output_buffer()
                        written = port.write(payload)
                        port.flush()
                        rx = read_exact(port, len(payload), timeout_s, self.stop_event)

                        if written == len(payload) and rx == payload:
                            self.emit("PASS", f"TX/RX {payload_hex}", pass_inc=1)
                        else:
                            rx_hex = rx.hex(" ").upper() if rx else "<none>"
                            self.emit(
                                "FAIL",
                                f"TX {payload_hex} RX {rx_hex}",
                                fail_inc=1,
                                log=True,
                            )

                        if self.stop_event.wait(interval_s):
                            break

            except (SerialException, OSError) as exc:
                if self.stop_event.is_set():
                    break
                self.emit("ERROR", str(exc), fail_inc=1, log=True)
                self.stop_event.wait(2.0)

        self.emit("Stopped", "Worker stopped", log=True)


class RS485PairWorker(threading.Thread):
    def __init__(self, index: int, config: dict, event_queue: queue.Queue):
        super().__init__(daemon=True)
        self.index = index
        self.config = config
        self.event_queue = event_queue
        self.stop_event = threading.Event()

    def stop(self) -> None:
        self.stop_event.set()

    def emit(
        self,
        status: str,
        last: str,
        pass_inc: int = 0,
        fail_inc: int = 0,
        log: bool = False,
    ) -> None:
        self.event_queue.put(
            {
                "group": "rs485",
                "index": self.index,
                "status": status,
                "last": last,
                "pass_inc": pass_inc,
                "fail_inc": fail_inc,
                "log": log,
            }
        )

    def open_port(self, port_name: str) -> serial.Serial:
        return serial.Serial(
            port=port_name,
            baudrate=int(self.config["baudrate"]),
            bytesize=int(self.config["bytesize"]),
            parity=str(self.config["parity"]).upper(),
            stopbits=float(self.config["stopbits"]),
            timeout=float(self.config["timeout_s"]),
            write_timeout=float(self.config["timeout_s"]),
        )

    def run(self) -> None:
        payload = bytes.fromhex(self.config["payload_hex"])
        timeout_s = max(float(self.config["timeout_s"]), 0.05)
        interval_s = max(int(self.config["interval_ms"]) / 1000.0, 0.05)
        startup_delay_s = max(as_float(self.config.get("startup_delay_s", 0.0), 0.0), 0.0)
        payload_hex = payload.hex(" ").upper()

        while not self.stop_event.is_set():
            try:
                with self.open_port(self.config["sender_port"]) as sender:
                    with self.open_port(self.config["echo_port"]) as echo:
                        self.emit("Running", "Ports open", log=True)
                        if startup_delay_s > 0:
                            self.emit("Standby", f"Startup delay {startup_delay_s:.1f}s")
                            if self.stop_event.wait(startup_delay_s):
                                break

                        while not self.stop_event.is_set():
                            sender.reset_input_buffer()
                            sender.reset_output_buffer()
                            echo.reset_input_buffer()
                            echo.reset_output_buffer()

                            sender.write(payload)
                            sender.flush()

                            seen = read_exact(echo, len(payload), timeout_s, self.stop_event)
                            if seen != payload:
                                seen_hex = seen.hex(" ").upper() if seen else "<none>"
                                self.emit(
                                    "FAIL",
                                    f"Echo RX {seen_hex}, expected {payload_hex}",
                                    fail_inc=1,
                                    log=True,
                                )
                                if self.stop_event.wait(interval_s):
                                    break
                                continue

                            echo.write(seen)
                            echo.flush()

                            bounced = read_exact(sender, len(payload), timeout_s, self.stop_event)
                            if bounced == payload:
                                self.emit("PASS", f"TX/RX {payload_hex}", pass_inc=1)
                            else:
                                bounced_hex = bounced.hex(" ").upper() if bounced else "<none>"
                                self.emit(
                                    "FAIL",
                                    f"Sender RX {bounced_hex}, expected {payload_hex}",
                                    fail_inc=1,
                                    log=True,
                                )

                            if self.stop_event.wait(interval_s):
                                break

            except (SerialException, OSError) as exc:
                if self.stop_event.is_set():
                    break
                self.emit("ERROR", str(exc), fail_inc=1, log=True)
                self.stop_event.wait(2.0)

        self.emit("Stopped", "Worker stopped", log=True)


class SerialTesterApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(f"Serial Loopback Tester v{APP_VERSION} (40x RS232 + 8x RS485 Pairs)")
        self.geometry("1500x900")
        self.minsize(1200, 720)

        self.settings_path = resolve_settings_path()
        self.settings = load_settings_file(self.settings_path)

        self.rs232_configs = self.settings["rs232_ports"]
        self.rs485_configs = self.settings["rs485_pairs"]
        self.ui_settings = self.settings["ui"]
        self.preset_configs = self.ui_settings["presets"]

        self.rs232_state = [self.new_state() for _ in range(DEFAULT_RS232_COUNT)]
        self.rs485_state = [self.new_state() for _ in range(DEFAULT_RS485_PAIR_COUNT)]
        self.overview_compact_var = tk.BooleanVar(value=bool(self.ui_settings.get("overview_compact_view", True)))

        self.event_queue: queue.Queue = queue.Queue()
        self.rs232_workers: dict[int, RS232Worker] = {}
        self.rs485_workers: dict[int, RS485PairWorker] = {}
        self.overview_rows: dict[tuple[str, int], dict] = {}
        self.overview_rows_canvas: tk.Canvas | None = None
        self.overview_rows_frame: ttk.Frame | None = None
        self.overview_rows_canvas_window: int | None = None
        self.com_port_values = [""]
        self.failure_counts: deque[tuple[int, int]] = deque()
        self.fault_rows_limit = 2000
        self.channel_fault_history: set[tuple[str, int]] = set()
        self.launch_autostart_scheduled = False
        self.app_start_monotonic = time.monotonic()
        self.preset_name_vars: list[tk.StringVar] = []
        self.preset_panels: list[ttk.LabelFrame] = []
        self.preset_port_listboxes: list[tk.Listbox] = []
        self.preset_port_options: list[str] = []
        self.preset_buttons: list[ttk.Button] = []
        self.is_fullscreen = False
        self.log_line_count = 0
        self.health_total_pass_var = tk.StringVar(value="Total Pass: 0")
        self.health_total_fail_var = tk.StringVar(value="Total Fail: 0")
        self.health_total_errors_var = tk.StringVar(value="Total Errors: 0")
        self.health_active_workers_var = tk.StringVar(value="Active Workers: 0")
        self.health_recent_fail_var = tk.StringVar(value=f"Fails Last {FAILURE_WINDOW_LABEL}: 0")
        self.health_runtime_var = tk.StringVar(value="Run Time: 00:00:00")
        self.health_fault_review_count_var = tk.StringVar(value="Faults Logged: 0")
        self.health_alarm_status_var = tk.StringVar(value="STANDBY")
        self.health_alarm_canvas: tk.Canvas | None = None
        self.health_alarm_rect: int | None = None
        self.overview_alarm_canvas: tk.Canvas | None = None
        self.overview_alarm_rect: int | None = None
        self.health_issue_tree: ttk.Treeview | None = None

        self._build_ui()
        self.refresh_com_port_options(show_message=False)
        self._populate_tables()
        self._select_first_rows()

        self.set_fullscreen(bool(self.ui_settings.get("start_fullscreen", False)))
        if bool(self.ui_settings.get("auto_start_after_launch_2s", True)):
            self.launch_autostart_scheduled = True
            self.after(2000, self._auto_start_after_launch)
        self._refresh_health_panel()
        self.after(100, self._process_worker_events)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.bind("<F11>", self._on_toggle_fullscreen)
        self.bind("<Escape>", self._on_exit_fullscreen)

    @staticmethod
    def new_state() -> dict:
        return {"status": "Idle", "pass_count": 0, "fail_count": 0, "last": ""}

    @staticmethod
    def _format_duration(total_seconds: float) -> str:
        seconds = max(int(total_seconds), 0)
        hours, remainder = divmod(seconds, 3600)
        minutes, secs = divmod(remainder, 60)
        return f"{hours:02d}:{minutes:02d}:{secs:02d}"

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        toolbar = ttk.Frame(self, padding=(8, 8))
        toolbar.grid(row=0, column=0, sticky="ew")

        ttk.Button(toolbar, text="Start All", command=self.start_all_tests).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(toolbar, text="Stop All", command=self.stop_all_tests).pack(side=tk.LEFT, padx=(0, 16))

        ttk.Button(toolbar, text="Start RS232", command=self.start_rs232_tests).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(toolbar, text="Stop RS232", command=self.stop_rs232_tests).pack(side=tk.LEFT, padx=(0, 16))

        ttk.Button(toolbar, text="Start RS485", command=self.start_rs485_tests).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(toolbar, text="Stop RS485", command=self.stop_rs485_tests).pack(side=tk.LEFT, padx=(0, 16))

        ttk.Button(toolbar, text="Save Settings", command=self.save_settings).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(toolbar, text="Reload Settings", command=self.reload_settings).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(toolbar, text="Refresh COM List", command=lambda: self.refresh_com_port_options(show_message=True)).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        self.fullscreen_button = ttk.Button(toolbar, text="Fullscreen", command=self.toggle_fullscreen)
        self.fullscreen_button.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(toolbar, text="Presets:").pack(side=tk.LEFT, padx=(6, 4))
        for idx in range(DEFAULT_PRESET_COUNT):
            button = ttk.Button(
                toolbar,
                text=self.preset_configs[idx]["name"],
                width=11,
                command=lambda i=idx: self.apply_preset(i),
            )
            button.pack(side=tk.LEFT, padx=(0, 4))
            self.preset_buttons.append(button)

        self.status_label = ttk.Label(
            toolbar,
            text=f"v{APP_VERSION} | Made by {APP_PUBLISHER} | Settings file: {self.settings_path}",
        )
        self.status_label.pack(side=tk.RIGHT)

        notebook = ttk.Notebook(self)
        notebook.grid(row=1, column=0, sticky="nsew")

        overview_tab = ttk.Frame(notebook, padding=(8, 8))
        health_tab = ttk.Frame(notebook, padding=(8, 8))
        faults_tab = ttk.Frame(notebook, padding=(8, 8))
        presets_tab = ttk.Frame(notebook, padding=(8, 8))
        rs232_tab = ttk.Frame(notebook, padding=(8, 8))
        rs485_tab = ttk.Frame(notebook, padding=(8, 8))
        settings_tab = ttk.Frame(notebook, padding=(8, 8))
        log_tab = ttk.Frame(notebook, padding=(8, 8))

        notebook.add(overview_tab, text="Overview")
        notebook.add(health_tab, text="Health")
        notebook.add(faults_tab, text="Fault Review")
        notebook.add(presets_tab, text="Presets")
        notebook.add(rs232_tab, text="RS232 Monitor")
        notebook.add(rs485_tab, text="RS485 Monitor")
        notebook.add(settings_tab, text="Settings")
        notebook.add(log_tab, text="Log")

        self._build_overview_tab(overview_tab)
        self._build_health_tab(health_tab)
        self._build_fault_review_tab(faults_tab)
        self._build_presets_tab(presets_tab)
        self.rs232_tree = self._build_rs232_monitor_tree(rs232_tab)
        self.rs485_tree = self._build_rs485_monitor_tree(rs485_tab)
        self._build_settings_tab(settings_tab)
        self._build_log_tab(log_tab)

    def set_fullscreen(self, enabled: bool) -> None:
        self.is_fullscreen = bool(enabled)
        self.attributes("-fullscreen", self.is_fullscreen)
        self.fullscreen_button.configure(text="Exit Fullscreen" if self.is_fullscreen else "Fullscreen")

    def toggle_fullscreen(self) -> None:
        self.set_fullscreen(not self.is_fullscreen)

    def exit_fullscreen(self) -> None:
        if not self.is_fullscreen:
            return
        self.set_fullscreen(False)

    def _on_start_fullscreen_preference_changed(self) -> None:
        self.ui_settings["start_fullscreen"] = bool(self.start_fullscreen_var.get())

    def _on_auto_start_launch_preference_changed(self) -> None:
        self.ui_settings["auto_start_after_launch_2s"] = bool(self.auto_start_launch_var.get())

    def _on_delay_comm_start_preference_changed(self) -> None:
        self.ui_settings["delay_comm_start_2s"] = bool(self.delay_comm_start_var.get())

    def _on_overview_layout_changed(self) -> None:
        self.ui_settings["overview_compact_view"] = bool(self.overview_compact_var.get())
        self._rebuild_overview_rows()

    def _on_toggle_fullscreen(self, _event: tk.Event) -> str:
        self.toggle_fullscreen()
        return "break"

    def _on_exit_fullscreen(self, _event: tk.Event) -> str:
        self.exit_fullscreen()
        return "break"

    def _auto_start_after_launch(self) -> None:
        self.launch_autostart_scheduled = False
        if not bool(self.ui_settings.get("auto_start_after_launch_2s", True)):
            return
        if self.rs232_workers or self.rs485_workers:
            return
        self.start_all_tests(startup_delay_s=0.0)
        self.append_log("Auto-started tests 2 seconds after launch.")

    def refresh_com_port_options(self, show_message: bool = False) -> None:
        ports: list[str] = []
        error_text = ""
        try:
            ports = sorted({str(item.device).strip().upper() for item in list_ports.comports() if item.device})
        except Exception as exc:  # pragma: no cover - defensive for platform/driver edge cases
            error_text = str(exc)

        self.com_port_values = [""] + ports
        for attr in ("rs232_port_combo", "rs485_sender_combo", "rs485_echo_combo"):
            combo = getattr(self, attr, None)
            if combo is not None:
                combo.configure(values=self.com_port_values)
        self._refresh_preset_port_options(show_message=False)

        if show_message:
            if error_text:
                messagebox.showwarning(
                    "COM list refresh",
                    f"Could not query COM ports from system.\n\n{error_text}\n\nYou can still type port names manually.",
                )
            else:
                messagebox.showinfo("COM list refresh", f"Detected {len(ports)} COM port(s).")

    @staticmethod
    def _com_port_sort_key(port: str) -> tuple[int, int, str]:
        text = port.strip().upper()
        if text.startswith("COM") and text[3:].isdigit():
            return (0, int(text[3:]), text)
        return (1, 0, text)

    def _collect_preset_port_options(self) -> list[str]:
        options: set[str] = set()
        for port in self.com_port_values:
            normalized = normalize_port_text(port)
            if normalized:
                options.add(normalized)
        for cfg in self.rs232_configs:
            normalized = normalize_port_text(cfg["port"])
            if normalized:
                options.add(normalized)
        for cfg in self.rs485_configs:
            sender = normalize_port_text(cfg["sender_port"])
            echo = normalize_port_text(cfg["echo_port"])
            if sender:
                options.add(sender)
            if echo:
                options.add(echo)
        return sorted(options, key=self._com_port_sort_key)

    def _refresh_preset_button_labels(self) -> None:
        for idx, button in enumerate(self.preset_buttons):
            if idx < len(self.preset_configs):
                preset_name = self.preset_configs[idx]["name"]
                button.configure(text=preset_name)
                if idx < len(self.preset_panels):
                    self.preset_panels[idx].configure(text=f"Preset {idx + 1}: {preset_name}")

    def _get_selected_ports_from_listbox(self, listbox: tk.Listbox) -> list[str]:
        selected_indices = listbox.curselection()
        ports: list[str] = []
        for sel_idx in selected_indices:
            port = normalize_port_text(listbox.get(sel_idx))
            if port:
                ports.append(port)
        return normalize_port_list(ports)

    def _on_preset_name_changed(self, idx: int) -> None:
        if not (0 <= idx < len(self.preset_configs)):
            return
        new_name = self.preset_name_vars[idx].get().strip() or f"Preset {idx + 1}"
        self.preset_configs[idx]["name"] = new_name
        self._refresh_preset_button_labels()

    def _on_preset_ports_selected(self, idx: int) -> None:
        if not (0 <= idx < len(self.preset_configs)) or idx >= len(self.preset_port_listboxes):
            return
        self.preset_configs[idx]["ports"] = self._get_selected_ports_from_listbox(self.preset_port_listboxes[idx])

    def _select_all_preset_ports(self, idx: int) -> None:
        if not (0 <= idx < len(self.preset_port_listboxes)):
            return
        listbox = self.preset_port_listboxes[idx]
        if listbox.size() > 0:
            listbox.selection_set(0, tk.END)
        self._on_preset_ports_selected(idx)

    def _clear_preset_ports(self, idx: int) -> None:
        if not (0 <= idx < len(self.preset_port_listboxes)):
            return
        listbox = self.preset_port_listboxes[idx]
        listbox.selection_clear(0, tk.END)
        self._on_preset_ports_selected(idx)

    def _refresh_preset_port_options(self, show_message: bool = False) -> None:
        self.preset_port_options = self._collect_preset_port_options()
        if not self.preset_port_listboxes:
            return

        for idx, listbox in enumerate(self.preset_port_listboxes):
            current_selected = set(self.preset_configs[idx]["ports"])
            listbox.delete(0, tk.END)
            for port in self.preset_port_options:
                listbox.insert(tk.END, port)
            for item_idx, port in enumerate(self.preset_port_options):
                if port in current_selected:
                    listbox.selection_set(item_idx)

        if show_message:
            messagebox.showinfo("Preset ports", f"Loaded {len(self.preset_port_options)} selectable COM port(s).")

    def _build_presets_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        top = ttk.Frame(parent)
        top.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        top.columnconfigure(0, weight=1)
        ttk.Label(top, text="Configure 5 named presets and select which COM ports each preset enables/disables.").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Button(top, text="Refresh Port Choices", command=lambda: self._refresh_preset_port_options(show_message=True)).grid(
            row=0, column=1, sticky="e", padx=(8, 0)
        )
        ttk.Button(top, text="Save Presets", command=lambda: self.save_settings(show_message=True)).grid(
            row=0, column=2, sticky="e", padx=(8, 0)
        )

        grid = ttk.Frame(parent)
        grid.grid(row=1, column=0, sticky="nsew")
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(1, weight=1)
        grid.rowconfigure(0, weight=1)
        grid.rowconfigure(1, weight=1)
        grid.rowconfigure(2, weight=1)

        self.preset_name_vars = []
        self.preset_panels = []
        self.preset_port_listboxes = []
        for idx in range(DEFAULT_PRESET_COUNT):
            preset_name = self.preset_configs[idx]["name"]
            pane = ttk.LabelFrame(grid, text=f"Preset {idx + 1}: {preset_name}", padding=8)
            pane.grid(row=idx // 2, column=idx % 2, sticky="nsew", padx=4, pady=4)
            pane.columnconfigure(1, weight=1)
            pane.rowconfigure(1, weight=1)
            self.preset_panels.append(pane)

            name_var = tk.StringVar(value=self.preset_configs[idx]["name"])
            self.preset_name_vars.append(name_var)
            name_var.trace_add("write", lambda *_args, i=idx: self._on_preset_name_changed(i))

            ttk.Label(pane, text="Name").grid(row=0, column=0, sticky="w")
            ttk.Entry(pane, textvariable=name_var).grid(row=0, column=1, sticky="ew", padx=(6, 0))

            listbox = tk.Listbox(pane, selectmode=tk.MULTIPLE, exportselection=False, height=7)
            listbox.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(6, 6))
            listbox.bind("<<ListboxSelect>>", lambda _event, i=idx: self._on_preset_ports_selected(i))
            self.preset_port_listboxes.append(listbox)

            actions = ttk.Frame(pane)
            actions.grid(row=2, column=0, columnspan=2, sticky="ew")
            ttk.Button(actions, text="Select All", command=lambda i=idx: self._select_all_preset_ports(i)).pack(
                side=tk.LEFT, padx=(0, 6)
            )
            ttk.Button(actions, text="Clear", command=lambda i=idx: self._clear_preset_ports(i)).pack(side=tk.LEFT)

        self._refresh_preset_button_labels()
        self._refresh_preset_port_options(show_message=False)

    def _sync_presets_from_ui(self) -> None:
        for idx in range(DEFAULT_PRESET_COUNT):
            if idx < len(self.preset_name_vars):
                self.preset_configs[idx]["name"] = self.preset_name_vars[idx].get().strip() or f"Preset {idx + 1}"
            if idx < len(self.preset_port_listboxes):
                self.preset_configs[idx]["ports"] = self._get_selected_ports_from_listbox(self.preset_port_listboxes[idx])
            else:
                self.preset_configs[idx]["ports"] = normalize_port_list(self.preset_configs[idx].get("ports", []))
        self._refresh_preset_button_labels()

    def _rs232_in_port_set(self, idx: int, selected_ports: set[str]) -> bool:
        cfg_port = normalize_port_text(self.rs232_configs[idx]["port"])
        return bool(cfg_port) and cfg_port in selected_ports

    def _rs485_in_port_set(self, idx: int, selected_ports: set[str]) -> bool:
        cfg = self.rs485_configs[idx]
        sender = normalize_port_text(cfg["sender_port"])
        echo = normalize_port_text(cfg["echo_port"])
        return bool(sender and echo) and sender in selected_ports and echo in selected_ports

    def apply_preset(self, idx: int) -> None:
        if not (0 <= idx < len(self.preset_configs)):
            return

        self._sync_presets_from_ui()
        preset = self.preset_configs[idx]
        selected_ports = set(normalize_port_list(preset.get("ports", [])))
        enabled_count = 0
        disabled_count = 0
        stopped = 0

        for rs232_idx in range(DEFAULT_RS232_COUNT):
            cfg = self.rs232_configs[rs232_idx]
            should_enable = self._rs232_in_port_set(rs232_idx, selected_ports)
            cfg["enabled"] = should_enable
            if should_enable:
                enabled_count += 1
            else:
                disabled_count += 1
            if (not should_enable) and (rs232_idx in self.rs232_workers):
                self.stop_single_test("rs232", rs232_idx, log_event=False)
                stopped += 1
            else:
                self.refresh_rs232_row(rs232_idx)

        for rs485_idx in range(DEFAULT_RS485_PAIR_COUNT):
            cfg = self.rs485_configs[rs485_idx]
            should_enable = self._rs485_in_port_set(rs485_idx, selected_ports)
            cfg["enabled"] = should_enable
            if should_enable:
                enabled_count += 1
            else:
                disabled_count += 1
            if (not should_enable) and (rs485_idx in self.rs485_workers):
                self.stop_single_test("rs485", rs485_idx, log_event=False)
                stopped += 1
            else:
                self.refresh_rs485_row(rs485_idx)

        self.save_settings(show_message=False)
        self.append_log(
            f'Preset "{preset["name"]}" applied ({enabled_count} enabled, {disabled_count} disabled, {stopped} stopped).'
        )

    def _build_health_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        summary = ttk.Frame(parent)
        summary.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        for col in range(3):
            summary.columnconfigure(col, weight=1)

        ttk.Label(summary, textvariable=self.health_total_pass_var).grid(row=0, column=0, sticky="w")
        ttk.Label(summary, textvariable=self.health_total_fail_var).grid(row=0, column=1, sticky="w")
        ttk.Label(summary, textvariable=self.health_total_errors_var).grid(row=0, column=2, sticky="w")
        ttk.Label(summary, textvariable=self.health_active_workers_var).grid(row=1, column=0, sticky="w", pady=(2, 0))
        ttk.Label(summary, textvariable=self.health_recent_fail_var).grid(row=1, column=1, sticky="w", pady=(2, 0))
        ttk.Label(summary, textvariable=self.health_runtime_var).grid(row=1, column=2, sticky="w", pady=(2, 0))

        alarm_and_issues = ttk.Panedwindow(parent, orient=tk.HORIZONTAL)
        alarm_and_issues.grid(row=1, column=0, sticky="nsew")

        alarm_frame = ttk.LabelFrame(alarm_and_issues, text="Alarm", padding=10)
        issues_frame = ttk.LabelFrame(alarm_and_issues, text="Current Issues (FAIL/ERROR)", padding=10)
        alarm_and_issues.add(alarm_frame, weight=1)
        alarm_and_issues.add(issues_frame, weight=3)

        alarm_frame.columnconfigure(0, weight=0)
        alarm_frame.columnconfigure(1, weight=1)
        alarm_frame.rowconfigure(1, weight=1)

        ttk.Label(alarm_frame, textvariable=self.health_alarm_status_var, font=("Segoe UI", 16, "bold")).grid(
            row=0, column=0, sticky="w", pady=(0, 8)
        )
        self.health_alarm_canvas = tk.Canvas(
            alarm_frame,
            width=220,
            height=44,
            highlightthickness=1,
            highlightbackground="#666666",
            bd=0,
        )
        self.health_alarm_rect = self.health_alarm_canvas.create_rectangle(1, 1, 219, 43, fill="#9CA3AF", outline="")
        self.health_alarm_canvas.grid(row=1, column=0, sticky="w")
        ttk.Label(
            alarm_frame,
            textvariable=self.health_fault_review_count_var,
            font=("Segoe UI", 11, "bold"),
        ).grid(row=1, column=1, sticky="w", padx=(10, 0))

        issues_frame.columnconfigure(0, weight=1)
        issues_frame.rowconfigure(0, weight=1)
        issue_cols = ("type", "name", "ports", "status", "last")
        self.health_issue_tree = ttk.Treeview(issues_frame, columns=issue_cols, show="headings", height=10)
        issue_headers = {
            "type": "Type",
            "name": "Name",
            "ports": "Port(s)",
            "status": "Status",
            "last": "Last",
        }
        issue_widths = {"type": 80, "name": 190, "ports": 180, "status": 90, "last": 420}
        for col in issue_cols:
            self.health_issue_tree.heading(col, text=issue_headers[col])
            self.health_issue_tree.column(col, width=issue_widths[col], anchor=tk.W)
        issue_scroll = ttk.Scrollbar(issues_frame, orient=tk.VERTICAL, command=self.health_issue_tree.yview)
        self.health_issue_tree.configure(yscrollcommand=issue_scroll.set)
        self.health_issue_tree.grid(row=0, column=0, sticky="nsew")
        issue_scroll.grid(row=0, column=1, sticky="ns")

    def _build_fault_review_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        top = ttk.Frame(parent)
        top.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        top.columnconfigure(0, weight=1)
        ttk.Label(top, text="Records transitions where a channel goes from PASS to FAIL/ERROR.").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Button(top, text="Clear Faults", command=self.clear_fault_review).grid(row=0, column=1, sticky="e")

        table_wrap = ttk.Frame(parent)
        table_wrap.grid(row=1, column=0, sticky="nsew")
        table_wrap.columnconfigure(0, weight=1)
        table_wrap.rowconfigure(0, weight=1)

        cols = ("time", "channel", "name", "ports", "from", "to", "message")
        self.fault_tree = ttk.Treeview(table_wrap, columns=cols, show="headings", height=14)
        headers = {
            "time": "Time",
            "channel": "Channel",
            "name": "Name",
            "ports": "Port(s)",
            "from": "From",
            "to": "To",
            "message": "Failure Detail",
        }
        widths = {"time": 145, "channel": 95, "name": 180, "ports": 170, "from": 70, "to": 85, "message": 520}
        for col in cols:
            self.fault_tree.heading(col, text=headers[col])
            self.fault_tree.column(col, width=widths[col], anchor=tk.W)

        scroll = ttk.Scrollbar(table_wrap, orient=tk.VERTICAL, command=self.fault_tree.yview)
        self.fault_tree.configure(yscrollcommand=scroll.set)
        self.fault_tree.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")

    def clear_fault_review(self) -> None:
        self.fault_tree.delete(*self.fault_tree.get_children())
        self.channel_fault_history.clear()
        self._refresh_health_panel()
        for idx in range(DEFAULT_RS232_COUNT):
            self._refresh_overview_row("rs232", idx)
        for idx in range(DEFAULT_RS485_PAIR_COUNT):
            self._refresh_overview_row("rs485", idx)
        self.append_log("Fault review list cleared.")

    def _record_fault_transition(self, group: str, idx: int, previous_status: str, new_status: str, detail: str) -> None:
        if previous_status != "PASS" or new_status not in {"FAIL", "ERROR"}:
            return

        self.channel_fault_history.add((group, idx))

        if group == "rs232":
            cfg = self.rs232_configs[idx]
            channel = f"RS232 #{idx + 1}"
            ports = cfg["port"]
        else:
            cfg = self.rs485_configs[idx]
            channel = f"RS485 #{idx + 1}"
            ports = f'{cfg["sender_port"]} <-> {cfg["echo_port"]}'

        stamp = time.strftime("%Y-%m-%d %H:%M:%S")
        self.fault_tree.insert(
            "",
            0,
            values=(
                stamp,
                channel,
                cfg["name"],
                ports,
                previous_status,
                new_status,
                detail,
            ),
        )

        children = self.fault_tree.get_children()
        if len(children) > self.fault_rows_limit:
            self.fault_tree.delete(*children[self.fault_rows_limit :])

    def _record_failure_event(self) -> None:
        now_sec = int(time.time())
        if self.failure_counts and self.failure_counts[-1][0] == now_sec:
            sec, count = self.failure_counts[-1]
            self.failure_counts[-1] = (sec, count + 1)
        else:
            self.failure_counts.append((now_sec, 1))
        self._trim_failure_counts(now_sec)

    def _trim_failure_counts(self, now_sec: int | None = None) -> None:
        current = now_sec if now_sec is not None else int(time.time())
        cutoff = current - FAILURE_WINDOW_SECONDS
        while self.failure_counts and self.failure_counts[0][0] <= cutoff:
            self.failure_counts.popleft()

    def _recent_failure_count(self) -> int:
        self._trim_failure_counts()
        return sum(count for _sec, count in self.failure_counts)

    def _refresh_health_panel(self) -> None:
        total_pass = sum(item["pass_count"] for item in self.rs232_state) + sum(item["pass_count"] for item in self.rs485_state)
        total_fail = sum(item["fail_count"] for item in self.rs232_state) + sum(item["fail_count"] for item in self.rs485_state)
        total_errors = total_fail
        active_workers = len(self.rs232_workers) + len(self.rs485_workers)
        recent_failures = self._recent_failure_count()
        runtime_seconds = time.monotonic() - self.app_start_monotonic
        runtime_text = self._format_duration(runtime_seconds)

        self.health_total_pass_var.set(f"Total Pass: {total_pass}")
        self.health_total_fail_var.set(f"Total Fail: {total_fail}")
        self.health_total_errors_var.set(f"Total Errors: {total_errors}")
        self.health_active_workers_var.set(f"Active Workers: {active_workers}")
        self.health_recent_fail_var.set(f"Fails Last {FAILURE_WINDOW_LABEL}: {recent_failures}")
        self.health_runtime_var.set(f"Run Time: {runtime_text}")
        self.health_fault_review_count_var.set(f"Faults Logged: {len(self.fault_tree.get_children())}")

        current_issues: list[tuple[str, str, str, str, str]] = []
        for idx, state in enumerate(self.rs232_state):
            if state["status"] in {"FAIL", "ERROR"}:
                cfg = self.rs232_configs[idx]
                current_issues.append(("RS232", cfg["name"], cfg["port"], state["status"], state["last"]))
        for idx, state in enumerate(self.rs485_state):
            if state["status"] in {"FAIL", "ERROR"}:
                cfg = self.rs485_configs[idx]
                current_issues.append(
                    ("RS485", cfg["name"], f'{cfg["sender_port"]} <-> {cfg["echo_port"]}', state["status"], state["last"])
                )

        fault_review_count = len(self.fault_tree.get_children())
        green_ready = runtime_seconds >= FAILURE_WINDOW_SECONDS and recent_failures == 0

        if current_issues:
            alarm_text = "ALARM"
            alarm_color = "#DC2626"
        elif active_workers > 0 and fault_review_count > 0:
            alarm_text = "RECOVERED"
            alarm_color = "#8B5CF6"
        elif active_workers > 0 and green_ready:
            alarm_text = "GOOD"
            alarm_color = "#22C55E"
        else:
            alarm_text = "STANDBY"
            alarm_color = "#9CA3AF"

        self.health_alarm_status_var.set(alarm_text)
        if self.health_alarm_canvas is not None and self.health_alarm_rect is not None:
            self.health_alarm_canvas.itemconfigure(self.health_alarm_rect, fill=alarm_color)
        if self.overview_alarm_canvas is not None and self.overview_alarm_rect is not None:
            self.overview_alarm_canvas.itemconfigure(self.overview_alarm_rect, fill=alarm_color)

        if self.health_issue_tree is not None:
            self.health_issue_tree.delete(*self.health_issue_tree.get_children())
            for issue in current_issues:
                self.health_issue_tree.insert("", tk.END, values=issue)

    def _build_rs232_monitor_tree(self, parent: ttk.Frame) -> ttk.Treeview:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)

        columns = ("idx", "enabled", "name", "port", "status", "pass", "fail", "last")
        tree = ttk.Treeview(parent, columns=columns, show="headings", height=24)

        headings = {
            "idx": "#",
            "enabled": "Enabled",
            "name": "Name",
            "port": "Port",
            "status": "Status",
            "pass": "Pass",
            "fail": "Fail",
            "last": "Last Result",
        }
        widths = {
            "idx": 45,
            "enabled": 70,
            "name": 180,
            "port": 90,
            "status": 90,
            "pass": 70,
            "fail": 70,
            "last": 680,
        }

        for col in columns:
            tree.heading(col, text=headings[col])
            tree.column(col, width=widths[col], anchor=tk.W if col in {"name", "last"} else tk.CENTER)

        yscroll = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=yscroll.set)

        tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        return tree

    def _build_overview_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)

        legend = ttk.Frame(parent)
        legend.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        legend.columnconfigure(0, weight=1)

        ttk.Label(
            legend,
            text="Combined live status: Green = good, Purple = recovered, Yellow = standby, Red = wrong message/error",
        ).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(
            legend,
            text="Compact View (2 Columns)",
            variable=self.overview_compact_var,
            command=self._on_overview_layout_changed,
        ).grid(row=0, column=1, sticky="e")

        health = ttk.LabelFrame(parent, text="Health", padding=(8, 6))
        health.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        for col in range(5):
            health.columnconfigure(col, weight=1)

        status_row = ttk.Frame(health)
        status_row.grid(row=0, column=0, columnspan=5, sticky="ew", pady=(0, 4))
        status_row.columnconfigure(4, weight=1)
        ttk.Label(status_row, textvariable=self.health_alarm_status_var, font=("Segoe UI", 12, "bold")).grid(
            row=0, column=0, sticky="w", padx=(0, 6)
        )
        self.overview_alarm_canvas = tk.Canvas(
            status_row,
            width=130,
            height=20,
            highlightthickness=1,
            highlightbackground="#666666",
            bd=0,
        )
        self.overview_alarm_rect = self.overview_alarm_canvas.create_rectangle(1, 1, 129, 19, fill="#9CA3AF", outline="")
        self.overview_alarm_canvas.grid(row=0, column=1, sticky="w")
        ttk.Label(
            status_row,
            textvariable=self.health_fault_review_count_var,
            font=("Segoe UI", 10, "bold"),
        ).grid(row=0, column=2, sticky="w", padx=(10, 0))

        ttk.Label(health, textvariable=self.health_total_pass_var).grid(row=1, column=0, sticky="w")
        ttk.Label(health, textvariable=self.health_total_fail_var).grid(row=1, column=1, sticky="w")
        ttk.Label(health, textvariable=self.health_active_workers_var).grid(row=1, column=2, sticky="w")
        ttk.Label(health, textvariable=self.health_recent_fail_var).grid(row=1, column=3, sticky="w")

        holder = ttk.Frame(parent)
        holder.grid(row=2, column=0, sticky="nsew")
        holder.columnconfigure(0, weight=1)
        holder.rowconfigure(0, weight=1)

        self.overview_rows_canvas = tk.Canvas(holder, highlightthickness=0)
        vscroll = ttk.Scrollbar(holder, orient=tk.VERTICAL, command=self.overview_rows_canvas.yview)
        self.overview_rows_frame = ttk.Frame(self.overview_rows_canvas)
        self.overview_rows_frame.columnconfigure(0, weight=1)
        self.overview_rows_frame.columnconfigure(1, weight=1)

        self.overview_rows_frame.bind(
            "<Configure>",
            lambda _event: self.overview_rows_canvas.configure(scrollregion=self.overview_rows_canvas.bbox("all")),
        )
        self.overview_rows_canvas_window = self.overview_rows_canvas.create_window((0, 0), window=self.overview_rows_frame, anchor="nw")
        self.overview_rows_canvas.bind(
            "<Configure>",
            lambda event: self.overview_rows_canvas.itemconfigure(self.overview_rows_canvas_window, width=event.width),
        )
        self.overview_rows_canvas.configure(yscrollcommand=vscroll.set)

        self.overview_rows_canvas.grid(row=0, column=0, sticky="nsew")
        vscroll.grid(row=0, column=1, sticky="ns")

        self._rebuild_overview_rows()

    def _rebuild_overview_rows(self) -> None:
        if self.overview_rows_frame is None:
            return

        for child in self.overview_rows_frame.winfo_children():
            child.destroy()
        self.overview_rows.clear()

        overview_idx = 0
        for idx in range(DEFAULT_RS232_COUNT):
            add_row = self._add_overview_compact_row if self.overview_compact_var.get() else self._add_overview_row
            add_row(
                self.overview_rows_frame,
                overview_idx // 2,
                overview_idx % 2,
                "rs232",
                idx,
                "RS232",
                self.rs232_configs[idx]["name"],
                self.rs232_configs[idx]["port"],
            )
            overview_idx += 1

        for idx in range(DEFAULT_RS485_PAIR_COUNT):
            cfg = self.rs485_configs[idx]
            add_row = self._add_overview_compact_row if self.overview_compact_var.get() else self._add_overview_row
            add_row(
                self.overview_rows_frame,
                overview_idx // 2,
                overview_idx % 2,
                "rs485",
                idx,
                "RS485",
                cfg["name"],
                f'{cfg["sender_port"]} <-> {cfg["echo_port"]}',
            )
            overview_idx += 1

        for idx in range(DEFAULT_RS232_COUNT):
            self._refresh_overview_row("rs232", idx)
        for idx in range(DEFAULT_RS485_PAIR_COUNT):
            self._refresh_overview_row("rs485", idx)

    def _add_overview_compact_row(
        self,
        parent: ttk.Frame,
        grid_row: int,
        grid_col: int,
        group: str,
        idx: int,
        row_type: str,
        name: str,
        ports: str,
    ) -> None:
        parent.columnconfigure(grid_col, weight=1)
        row = ttk.Frame(parent, padding=(4, 2))
        row.grid(row=grid_row, column=grid_col, sticky="ew", padx=4, pady=2)
        row.columnconfigure(2, weight=1)

        name_var = tk.StringVar(value=name)
        ports_var = tk.StringVar(value=ports)
        state_var = tk.StringVar(value="Standby")

        ttk.Label(row, text=row_type, width=7).grid(row=0, column=0, sticky="w")
        ttk.Label(row, textvariable=name_var, width=20).grid(row=0, column=1, sticky="w")
        ttk.Label(row, textvariable=ports_var).grid(row=0, column=2, sticky="w")
        ttk.Label(row, textvariable=state_var, width=11).grid(row=0, column=3, sticky="w", padx=(8, 0))

        bar_canvas = tk.Canvas(
            row,
            width=90,
            height=14,
            highlightthickness=1,
            highlightbackground="#777777",
            bd=0,
        )
        bar_rect = bar_canvas.create_rectangle(1, 1, 89, 13, fill="#F0B429", outline="")
        bar_canvas.grid(row=0, column=4, sticky="w", padx=(8, 0))

        start_button = ttk.Button(
            row,
            text="Start",
            width=6,
            command=lambda g=group, i=idx: self.start_single_test(g, i),
        )
        start_button.grid(row=0, column=5, sticky="w", padx=(8, 0))
        stop_button = ttk.Button(
            row,
            text="Stop",
            width=6,
            command=lambda g=group, i=idx: self.stop_single_test(g, i),
        )
        stop_button.grid(row=0, column=6, sticky="w", padx=(4, 0))

        self.overview_rows[(group, idx)] = {
            "name_var": name_var,
            "ports_var": ports_var,
            "state_var": state_var,
            "bar_canvas": bar_canvas,
            "bar_rect": bar_rect,
            "start_button": start_button,
            "stop_button": stop_button,
        }

    def _add_overview_row(
        self,
        parent: ttk.Frame,
        grid_row: int,
        grid_col: int,
        group: str,
        idx: int,
        row_type: str,
        name: str,
        ports: str,
    ) -> None:
        parent.columnconfigure(grid_col, weight=1)
        row = ttk.LabelFrame(parent, text=f"{row_type} #{idx + 1}", padding=(8, 6))
        row.grid(row=grid_row, column=grid_col, sticky="nsew", padx=4, pady=4)
        row.columnconfigure(0, weight=1)

        name_var = tk.StringVar(value=name)
        ports_var = tk.StringVar(value=ports)
        state_var = tk.StringVar(value="Standby")

        ttk.Label(row, textvariable=name_var, width=28).grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(row, textvariable=ports_var, width=28).grid(row=1, column=0, columnspan=3, sticky="w", pady=(1, 4))
        ttk.Label(row, textvariable=state_var, width=12).grid(row=2, column=0, sticky="w")

        bar_canvas = tk.Canvas(
            row,
            width=120,
            height=16,
            highlightthickness=1,
            highlightbackground="#777777",
            bd=0,
        )
        bar_rect = bar_canvas.create_rectangle(1, 1, 119, 15, fill="#F0B429", outline="")
        bar_canvas.grid(row=2, column=1, sticky="w", padx=(6, 6))

        buttons = ttk.Frame(row)
        buttons.grid(row=2, column=2, sticky="e")
        start_button = ttk.Button(
            buttons,
            text="Start",
            width=7,
            command=lambda g=group, i=idx: self.start_single_test(g, i),
        )
        start_button.pack(side=tk.LEFT, padx=(0, 4))
        stop_button = ttk.Button(
            buttons,
            text="Stop",
            width=7,
            command=lambda g=group, i=idx: self.stop_single_test(g, i),
        )
        stop_button.pack(side=tk.LEFT)

        self.overview_rows[(group, idx)] = {
            "name_var": name_var,
            "ports_var": ports_var,
            "state_var": state_var,
            "bar_canvas": bar_canvas,
            "bar_rect": bar_rect,
            "start_button": start_button,
            "stop_button": stop_button,
        }

    def _status_to_overview_state(self, group: str, idx: int, status: str) -> tuple[str, str]:
        upper = status.strip().upper()
        if upper == "PASS":
            if (group, idx) in self.channel_fault_history:
                return ("#8B5CF6", "Recovered")
            return ("#22C55E", "Good")
        if upper in {"FAIL", "ERROR"}:
            return ("#DC2626", "Wrong Message")
        return ("#F0B429", "Standby")

    def _refresh_overview_row(self, group: str, idx: int) -> None:
        row = self.overview_rows.get((group, idx))
        if not row:
            return

        if group == "rs232":
            cfg = self.rs232_configs[idx]
            state = self.rs232_state[idx]
            ports = cfg["port"]
        else:
            cfg = self.rs485_configs[idx]
            state = self.rs485_state[idx]
            ports = f'{cfg["sender_port"]} <-> {cfg["echo_port"]}'

        color, state_text = self._status_to_overview_state(group, idx, state["status"])
        row["name_var"].set(cfg["name"])
        row["ports_var"].set(ports)
        row["state_var"].set(state_text)
        row["bar_canvas"].itemconfigure(row["bar_rect"], fill=color)

        if group == "rs232":
            running = idx in self.rs232_workers
            startable = self._is_rs232_startable(idx)
        else:
            running = idx in self.rs485_workers
            startable = self._is_rs485_startable(idx)

        row["start_button"].configure(state=(tk.DISABLED if running or not startable else tk.NORMAL))
        row["stop_button"].configure(state=(tk.NORMAL if running else tk.DISABLED))

    def _build_rs485_monitor_tree(self, parent: ttk.Frame) -> ttk.Treeview:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)

        columns = ("idx", "enabled", "name", "sender", "echo", "status", "pass", "fail", "last")
        tree = ttk.Treeview(parent, columns=columns, show="headings", height=10)

        headings = {
            "idx": "#",
            "enabled": "Enabled",
            "name": "Name",
            "sender": "Sender Port",
            "echo": "Echo Port",
            "status": "Status",
            "pass": "Pass",
            "fail": "Fail",
            "last": "Last Result",
        }
        widths = {
            "idx": 45,
            "enabled": 70,
            "name": 180,
            "sender": 110,
            "echo": 110,
            "status": 90,
            "pass": 70,
            "fail": 70,
            "last": 720,
        }

        for col in columns:
            tree.heading(col, text=headings[col])
            tree.column(col, width=widths[col], anchor=tk.W if col in {"name", "last"} else tk.CENTER)

        yscroll = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=yscroll.set)

        tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        return tree
    def _build_settings_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        ui_section = ttk.LabelFrame(parent, text="Application", padding=10)
        ui_section.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ui_section.columnconfigure(0, weight=1)

        self.start_fullscreen_var = tk.BooleanVar(value=bool(self.ui_settings.get("start_fullscreen", False)))
        ttk.Checkbutton(
            ui_section,
            text="Start application in fullscreen",
            variable=self.start_fullscreen_var,
            command=self._on_start_fullscreen_preference_changed,
        ).grid(row=0, column=0, sticky="w")
        self.auto_start_launch_var = tk.BooleanVar(value=bool(self.ui_settings.get("auto_start_after_launch_2s", True)))
        ttk.Checkbutton(
            ui_section,
            text="Auto-start tests 2 seconds after launch",
            variable=self.auto_start_launch_var,
            command=self._on_auto_start_launch_preference_changed,
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.delay_comm_start_var = tk.BooleanVar(value=bool(self.ui_settings.get("delay_comm_start_2s", True)))
        ttk.Checkbutton(
            ui_section,
            text="Delay communications startup by 2 seconds",
            variable=self.delay_comm_start_var,
            command=self._on_delay_comm_start_preference_changed,
        ).grid(row=2, column=0, sticky="w", pady=(4, 0))
        ttk.Label(
            ui_section,
            text="Save settings to keep this as default on next launch.",
        ).grid(row=3, column=0, sticky="w", pady=(4, 0))
        all_buttons = ttk.Frame(ui_section)
        all_buttons.grid(row=4, column=0, sticky="w", pady=(8, 0))
        ttk.Button(all_buttons, text="Enable All Ports", command=self.enable_all_ports).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(all_buttons, text="Disable All Ports", command=self.disable_all_ports).pack(side=tk.LEFT)

        panes = ttk.Panedwindow(parent, orient=tk.VERTICAL)
        panes.grid(row=1, column=0, sticky="nsew")

        rs232_section = ttk.Frame(panes, padding=(0, 0, 0, 6))
        rs485_section = ttk.Frame(panes, padding=(0, 6, 0, 0))
        panes.add(rs232_section, weight=3)
        panes.add(rs485_section, weight=2)

        self._build_rs232_settings_section(rs232_section)
        self._build_rs485_settings_section(rs485_section)

    def _set_all_ports_enabled(self, enabled: bool) -> None:
        action_word = "enable" if enabled else "disable"
        answer = messagebox.askyesno(
            f"{action_word.title()} all ports",
            f"{action_word.title()} all RS232 and RS485 channels?",
        )
        if not answer:
            return

        for idx, cfg in enumerate(self.rs232_configs):
            cfg["enabled"] = enabled
            if enabled:
                self.refresh_rs232_row(idx)
            else:
                self.stop_single_test("rs232", idx, log_event=False)

        for idx, cfg in enumerate(self.rs485_configs):
            cfg["enabled"] = enabled
            if enabled:
                self.refresh_rs485_row(idx)
            else:
                self.stop_single_test("rs485", idx, log_event=False)

        self._on_rs232_settings_select(None)
        self._on_rs485_settings_select(None)
        self.save_settings(show_message=False)
        self.append_log(f"{action_word.title()}d all RS232 and RS485 channels from Settings.")

    def enable_all_ports(self) -> None:
        self._set_all_ports_enabled(True)

    def disable_all_ports(self) -> None:
        self._set_all_ports_enabled(False)

    def _build_rs232_settings_section(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=3)
        parent.columnconfigure(1, weight=2)
        parent.rowconfigure(1, weight=1)

        ttk.Label(parent, text="RS232 Port Settings (40 ports)").grid(row=0, column=0, sticky="w", pady=(0, 6))

        table_wrap = ttk.Frame(parent)
        table_wrap.grid(row=1, column=0, sticky="nsew", padx=(0, 8))
        table_wrap.columnconfigure(0, weight=1)
        table_wrap.rowconfigure(0, weight=1)

        columns = ("idx", "enabled", "name", "port", "baud", "payload", "interval", "timeout")
        self.rs232_settings_tree = ttk.Treeview(table_wrap, columns=columns, show="headings", height=10, selectmode="browse")

        headings = {
            "idx": "#",
            "enabled": "Enabled",
            "name": "Name",
            "port": "Port",
            "baud": "Baud",
            "payload": "Payload Hex",
            "interval": "Interval ms",
            "timeout": "Timeout s",
        }
        widths = {
            "idx": 45,
            "enabled": 70,
            "name": 180,
            "port": 90,
            "baud": 90,
            "payload": 130,
            "interval": 95,
            "timeout": 85,
        }

        for col in columns:
            self.rs232_settings_tree.heading(col, text=headings[col])
            self.rs232_settings_tree.column(col, width=widths[col], anchor=tk.CENTER if col != "name" else tk.W)

        rs232_scroll = ttk.Scrollbar(table_wrap, orient=tk.VERTICAL, command=self.rs232_settings_tree.yview)
        self.rs232_settings_tree.configure(yscrollcommand=rs232_scroll.set)

        self.rs232_settings_tree.grid(row=0, column=0, sticky="nsew")
        rs232_scroll.grid(row=0, column=1, sticky="ns")
        self.rs232_settings_tree.bind("<<TreeviewSelect>>", self._on_rs232_settings_select)

        editor = ttk.LabelFrame(parent, text="Edit Selected RS232 Port", padding=10)
        editor.grid(row=1, column=1, sticky="nsew")
        editor.columnconfigure(1, weight=1)

        self.rs232_var_enabled = tk.BooleanVar(value=True)
        self.rs232_var_name = tk.StringVar()
        self.rs232_var_port = tk.StringVar()
        self.rs232_var_baud = tk.StringVar(value="9600")
        self.rs232_var_bytesize = tk.StringVar(value="8")
        self.rs232_var_parity = tk.StringVar(value="N")
        self.rs232_var_stopbits = tk.StringVar(value="1")
        self.rs232_var_timeout = tk.StringVar(value="0.5")
        self.rs232_var_payload = tk.StringVar(value="55AA")
        self.rs232_var_interval = tk.StringVar(value="100")

        row = 0
        ttk.Checkbutton(editor, text="Enabled", variable=self.rs232_var_enabled).grid(row=row, column=0, columnspan=2, sticky="w")
        row += 1
        self._labeled_entry(editor, "Name", self.rs232_var_name, row)
        row += 1
        ttk.Label(editor, text="Port").grid(row=row, column=0, sticky="w", pady=2)
        self.rs232_port_combo = ttk.Combobox(
            editor,
            textvariable=self.rs232_var_port,
            values=self.com_port_values,
            state="normal",
        )
        self.rs232_port_combo.grid(row=row, column=1, sticky="ew", pady=2)
        row += 1
        self._labeled_entry(editor, "Baudrate", self.rs232_var_baud, row)
        row += 1
        self._labeled_combobox(editor, "Bytesize", self.rs232_var_bytesize, BYTESIZE_OPTIONS, row)
        row += 1
        self._labeled_combobox(editor, "Parity", self.rs232_var_parity, PARITY_OPTIONS, row)
        row += 1
        self._labeled_combobox(editor, "Stopbits", self.rs232_var_stopbits, STOPBITS_OPTIONS, row)
        row += 1
        self._labeled_entry(editor, "Timeout (s)", self.rs232_var_timeout, row)
        row += 1
        self._labeled_entry(editor, "Payload Hex", self.rs232_var_payload, row)
        row += 1
        self._labeled_entry(editor, "Interval (ms)", self.rs232_var_interval, row)
        row += 1

        ttk.Button(editor, text="Apply RS232 Changes", command=self.apply_rs232_changes).grid(
            row=row, column=0, columnspan=2, sticky="ew", pady=(8, 0)
        )
        row += 1
        ttk.Button(
            editor,
            text="Apply To All RS232 (Keep Name/Port)",
            command=self.apply_rs232_common_changes_to_all,
        ).grid(row=row, column=0, columnspan=2, sticky="ew", pady=(6, 0))

    def _build_rs485_settings_section(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=3)
        parent.columnconfigure(1, weight=2)
        parent.rowconfigure(1, weight=1)

        ttk.Label(parent, text="RS485 Pair Settings (8 pairs)").grid(row=0, column=0, sticky="w", pady=(0, 6))

        table_wrap = ttk.Frame(parent)
        table_wrap.grid(row=1, column=0, sticky="nsew", padx=(0, 8))
        table_wrap.columnconfigure(0, weight=1)
        table_wrap.rowconfigure(0, weight=1)

        columns = (
            "idx",
            "enabled",
            "name",
            "sender",
            "echo",
            "baud",
            "payload",
            "interval",
            "timeout",
        )
        self.rs485_settings_tree = ttk.Treeview(table_wrap, columns=columns, show="headings", height=6, selectmode="browse")

        headings = {
            "idx": "#",
            "enabled": "Enabled",
            "name": "Name",
            "sender": "Sender",
            "echo": "Echo",
            "baud": "Baud",
            "payload": "Payload Hex",
            "interval": "Interval ms",
            "timeout": "Timeout s",
        }
        widths = {
            "idx": 45,
            "enabled": 70,
            "name": 180,
            "sender": 90,
            "echo": 90,
            "baud": 80,
            "payload": 120,
            "interval": 90,
            "timeout": 80,
        }

        for col in columns:
            self.rs485_settings_tree.heading(col, text=headings[col])
            self.rs485_settings_tree.column(col, width=widths[col], anchor=tk.CENTER if col != "name" else tk.W)

        rs485_scroll = ttk.Scrollbar(table_wrap, orient=tk.VERTICAL, command=self.rs485_settings_tree.yview)
        self.rs485_settings_tree.configure(yscrollcommand=rs485_scroll.set)

        self.rs485_settings_tree.grid(row=0, column=0, sticky="nsew")
        rs485_scroll.grid(row=0, column=1, sticky="ns")
        self.rs485_settings_tree.bind("<<TreeviewSelect>>", self._on_rs485_settings_select)

        editor = ttk.LabelFrame(parent, text="Edit Selected RS485 Pair", padding=10)
        editor.grid(row=1, column=1, sticky="nsew")
        editor.columnconfigure(1, weight=1)

        self.rs485_var_enabled = tk.BooleanVar(value=True)
        self.rs485_var_name = tk.StringVar()
        self.rs485_var_sender = tk.StringVar()
        self.rs485_var_echo = tk.StringVar()
        self.rs485_var_baud = tk.StringVar(value="9600")
        self.rs485_var_bytesize = tk.StringVar(value="8")
        self.rs485_var_parity = tk.StringVar(value="N")
        self.rs485_var_stopbits = tk.StringVar(value="1")
        self.rs485_var_timeout = tk.StringVar(value="0.5")
        self.rs485_var_payload = tk.StringVar(value="A55A")
        self.rs485_var_interval = tk.StringVar(value="100")

        row = 0
        ttk.Checkbutton(editor, text="Enabled", variable=self.rs485_var_enabled).grid(row=row, column=0, columnspan=2, sticky="w")
        row += 1
        self._labeled_entry(editor, "Name", self.rs485_var_name, row)
        row += 1
        ttk.Label(editor, text="Sender Port").grid(row=row, column=0, sticky="w", pady=2)
        self.rs485_sender_combo = ttk.Combobox(
            editor,
            textvariable=self.rs485_var_sender,
            values=self.com_port_values,
            state="normal",
        )
        self.rs485_sender_combo.grid(row=row, column=1, sticky="ew", pady=2)
        row += 1
        ttk.Label(editor, text="Echo Port").grid(row=row, column=0, sticky="w", pady=2)
        self.rs485_echo_combo = ttk.Combobox(
            editor,
            textvariable=self.rs485_var_echo,
            values=self.com_port_values,
            state="normal",
        )
        self.rs485_echo_combo.grid(row=row, column=1, sticky="ew", pady=2)
        row += 1
        self._labeled_entry(editor, "Baudrate", self.rs485_var_baud, row)
        row += 1
        self._labeled_combobox(editor, "Bytesize", self.rs485_var_bytesize, BYTESIZE_OPTIONS, row)
        row += 1
        self._labeled_combobox(editor, "Parity", self.rs485_var_parity, PARITY_OPTIONS, row)
        row += 1
        self._labeled_combobox(editor, "Stopbits", self.rs485_var_stopbits, STOPBITS_OPTIONS, row)
        row += 1
        self._labeled_entry(editor, "Timeout (s)", self.rs485_var_timeout, row)
        row += 1
        self._labeled_entry(editor, "Payload Hex", self.rs485_var_payload, row)
        row += 1
        self._labeled_entry(editor, "Interval (ms)", self.rs485_var_interval, row)
        row += 1

        ttk.Button(editor, text="Apply RS485 Changes", command=self.apply_rs485_changes).grid(
            row=row, column=0, columnspan=2, sticky="ew", pady=(8, 0)
        )
        row += 1
        ttk.Button(
            editor,
            text="Apply To All RS485 (Keep Name/Ports)",
            command=self.apply_rs485_common_changes_to_all,
        ).grid(row=row, column=0, columnspan=2, sticky="ew", pady=(6, 0))

    def _build_log_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)

        self.log_text = scrolledtext.ScrolledText(parent, wrap=tk.WORD, state=tk.DISABLED)
        self.log_text.grid(row=0, column=0, sticky="nsew")

    def _labeled_entry(self, parent: ttk.Frame, label: str, variable: tk.StringVar, row: int) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=2)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", pady=2)

    def _labeled_combobox(
        self,
        parent: ttk.Frame,
        label: str,
        variable: tk.StringVar,
        values: tuple[str, ...],
        row: int,
    ) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=2)
        ttk.Combobox(parent, textvariable=variable, values=values, state="readonly").grid(
            row=row, column=1, sticky="ew", pady=2
        )

    def _populate_tables(self) -> None:
        self.rs232_tree.delete(*self.rs232_tree.get_children())
        self.rs485_tree.delete(*self.rs485_tree.get_children())
        self.rs232_settings_tree.delete(*self.rs232_settings_tree.get_children())
        self.rs485_settings_tree.delete(*self.rs485_settings_tree.get_children())

        for idx in range(DEFAULT_RS232_COUNT):
            self.rs232_tree.insert("", tk.END, iid=f"m232_{idx}")
            self.rs232_settings_tree.insert("", tk.END, iid=f"s232_{idx}")
            self.refresh_rs232_row(idx)

        for idx in range(DEFAULT_RS485_PAIR_COUNT):
            self.rs485_tree.insert("", tk.END, iid=f"m485_{idx}")
            self.rs485_settings_tree.insert("", tk.END, iid=f"s485_{idx}")
            self.refresh_rs485_row(idx)

    def _select_first_rows(self) -> None:
        if self.rs232_settings_tree.get_children():
            self.rs232_settings_tree.selection_set("s232_0")
            self._on_rs232_settings_select(None)
        if self.rs485_settings_tree.get_children():
            self.rs485_settings_tree.selection_set("s485_0")
            self._on_rs485_settings_select(None)

    def refresh_rs232_row(self, idx: int) -> None:
        cfg = self.rs232_configs[idx]
        state = self.rs232_state[idx]

        self.rs232_tree.item(
            f"m232_{idx}",
            values=(
                idx + 1,
                "Yes" if cfg["enabled"] else "No",
                cfg["name"],
                cfg["port"],
                state["status"],
                state["pass_count"],
                state["fail_count"],
                state["last"],
            ),
        )

        self.rs232_settings_tree.item(
            f"s232_{idx}",
            values=(
                idx + 1,
                "Yes" if cfg["enabled"] else "No",
                cfg["name"],
                cfg["port"],
                cfg["baudrate"],
                cfg["payload_hex"],
                cfg["interval_ms"],
                cfg["timeout_s"],
            ),
        )
        self._refresh_overview_row("rs232", idx)

    def refresh_rs485_row(self, idx: int) -> None:
        cfg = self.rs485_configs[idx]
        state = self.rs485_state[idx]

        self.rs485_tree.item(
            f"m485_{idx}",
            values=(
                idx + 1,
                "Yes" if cfg["enabled"] else "No",
                cfg["name"],
                cfg["sender_port"],
                cfg["echo_port"],
                state["status"],
                state["pass_count"],
                state["fail_count"],
                state["last"],
            ),
        )

        self.rs485_settings_tree.item(
            f"s485_{idx}",
            values=(
                idx + 1,
                "Yes" if cfg["enabled"] else "No",
                cfg["name"],
                cfg["sender_port"],
                cfg["echo_port"],
                cfg["baudrate"],
                cfg["payload_hex"],
                cfg["interval_ms"],
                cfg["timeout_s"],
            ),
        )
        self._refresh_overview_row("rs485", idx)

    def _selected_index(self, tree: ttk.Treeview, prefix: str) -> int | None:
        selected = tree.selection()
        if not selected:
            return None
        item_id = selected[0]
        if not item_id.startswith(prefix):
            return None
        return int(item_id.split("_")[1])

    def _on_rs232_settings_select(self, _event: object) -> None:
        idx = self._selected_index(self.rs232_settings_tree, "s232_")
        if idx is None:
            return

        cfg = self.rs232_configs[idx]
        self.rs232_var_enabled.set(bool(cfg["enabled"]))
        self.rs232_var_name.set(str(cfg["name"]))
        self.rs232_var_port.set(str(cfg["port"]))
        self.rs232_var_baud.set(str(cfg["baudrate"]))
        self.rs232_var_bytesize.set(str(cfg["bytesize"]))
        self.rs232_var_parity.set(str(cfg["parity"]))
        self.rs232_var_stopbits.set(stopbits_to_text(cfg["stopbits"]))
        self.rs232_var_timeout.set(str(cfg["timeout_s"]))
        self.rs232_var_payload.set(str(cfg["payload_hex"]))
        self.rs232_var_interval.set(str(cfg["interval_ms"]))

    def _on_rs485_settings_select(self, _event: object) -> None:
        idx = self._selected_index(self.rs485_settings_tree, "s485_")
        if idx is None:
            return

        cfg = self.rs485_configs[idx]
        self.rs485_var_enabled.set(bool(cfg["enabled"]))
        self.rs485_var_name.set(str(cfg["name"]))
        self.rs485_var_sender.set(str(cfg["sender_port"]))
        self.rs485_var_echo.set(str(cfg["echo_port"]))
        self.rs485_var_baud.set(str(cfg["baudrate"]))
        self.rs485_var_bytesize.set(str(cfg["bytesize"]))
        self.rs485_var_parity.set(str(cfg["parity"]))
        self.rs485_var_stopbits.set(stopbits_to_text(cfg["stopbits"]))
        self.rs485_var_timeout.set(str(cfg["timeout_s"]))
        self.rs485_var_payload.set(str(cfg["payload_hex"]))
        self.rs485_var_interval.set(str(cfg["interval_ms"]))
    def _read_rs232_common_editor_values(self) -> dict:
        baudrate = int(self.rs232_var_baud.get().strip())
        if baudrate <= 0:
            raise ValueError("Baudrate must be positive.")

        bytesize = int(self.rs232_var_bytesize.get().strip())
        if str(bytesize) not in BYTESIZE_OPTIONS:
            raise ValueError("Bytesize must be one of 5, 6, 7, 8.")

        parity = self.rs232_var_parity.get().strip().upper()
        if parity not in PARITY_OPTIONS:
            raise ValueError("Parity must be one of N, E, O, M, S.")

        stopbits = parse_stopbits(self.rs232_var_stopbits.get())

        timeout_s = float(self.rs232_var_timeout.get().strip())
        if timeout_s <= 0:
            raise ValueError("Timeout must be greater than 0.")

        interval_ms = int(self.rs232_var_interval.get().strip())
        if interval_ms < 50:
            raise ValueError("Interval must be at least 50 ms.")

        payload_hex = validate_hex_payload(self.rs232_var_payload.get())

        return {
            "enabled": bool(self.rs232_var_enabled.get()),
            "baudrate": baudrate,
            "bytesize": bytesize,
            "parity": parity,
            "stopbits": stopbits,
            "timeout_s": timeout_s,
            "payload_hex": payload_hex,
            "interval_ms": interval_ms,
        }

    def apply_rs232_changes(self) -> None:
        idx = self._selected_index(self.rs232_settings_tree, "s232_")
        if idx is None:
            messagebox.showinfo("Select row", "Select one RS232 row to edit.")
            return

        try:
            name = self.rs232_var_name.get().strip() or f"RS232 {idx + 1}"
            port = self.rs232_var_port.get().strip().upper()

            common_values = self._read_rs232_common_editor_values()
        except ValueError as exc:
            messagebox.showerror("Invalid RS232 settings", str(exc))
            return

        cfg = self.rs232_configs[idx]
        cfg.update(
            {
                "name": name,
                "port": port,
                **common_values,
            }
        )

        self._stop_rs232_worker_if_disabled(idx)
        self.refresh_rs232_row(idx)
        self.save_settings(show_message=False)
        self.append_log(f"RS232 #{idx + 1} settings updated.")

    def apply_rs232_common_changes_to_all(self) -> None:
        try:
            common_values = self._read_rs232_common_editor_values()
        except ValueError as exc:
            messagebox.showerror("Invalid RS232 settings", str(exc))
            return

        answer = messagebox.askyesno(
            "Apply to all RS232",
            "Apply the selected RS232 settings to all ports while keeping each Name and Port unchanged?",
        )
        if not answer:
            return

        for idx, cfg in enumerate(self.rs232_configs):
            cfg.update(common_values)
            self._stop_rs232_worker_if_disabled(idx)
            self.refresh_rs232_row(idx)

        self.save_settings(show_message=False)
        self.append_log("Applied RS232 common settings to all ports (Name/Port kept).")

    def _stop_rs232_worker_if_disabled(self, idx: int) -> None:
        cfg = self.rs232_configs[idx]
        is_disabled = (not cfg["enabled"]) or (not str(cfg["port"]).strip())
        if not is_disabled:
            return

        self.stop_single_test("rs232", idx, log_event=False)

    def _read_rs485_common_editor_values(self) -> dict:
        baudrate = int(self.rs485_var_baud.get().strip())
        if baudrate <= 0:
            raise ValueError("Baudrate must be positive.")

        bytesize = int(self.rs485_var_bytesize.get().strip())
        if str(bytesize) not in BYTESIZE_OPTIONS:
            raise ValueError("Bytesize must be one of 5, 6, 7, 8.")

        parity = self.rs485_var_parity.get().strip().upper()
        if parity not in PARITY_OPTIONS:
            raise ValueError("Parity must be one of N, E, O, M, S.")

        stopbits = parse_stopbits(self.rs485_var_stopbits.get())

        timeout_s = float(self.rs485_var_timeout.get().strip())
        if timeout_s <= 0:
            raise ValueError("Timeout must be greater than 0.")

        interval_ms = int(self.rs485_var_interval.get().strip())
        if interval_ms < 50:
            raise ValueError("Interval must be at least 50 ms.")

        payload_hex = validate_hex_payload(self.rs485_var_payload.get())

        return {
            "enabled": bool(self.rs485_var_enabled.get()),
            "baudrate": baudrate,
            "bytesize": bytesize,
            "parity": parity,
            "stopbits": stopbits,
            "timeout_s": timeout_s,
            "payload_hex": payload_hex,
            "interval_ms": interval_ms,
        }

    def apply_rs485_changes(self) -> None:
        idx = self._selected_index(self.rs485_settings_tree, "s485_")
        if idx is None:
            messagebox.showinfo("Select row", "Select one RS485 row to edit.")
            return

        try:
            name = self.rs485_var_name.get().strip() or f"RS485 Pair {idx + 1}"
            sender_port = self.rs485_var_sender.get().strip().upper()
            echo_port = self.rs485_var_echo.get().strip().upper()

            common_values = self._read_rs485_common_editor_values()
        except ValueError as exc:
            messagebox.showerror("Invalid RS485 settings", str(exc))
            return

        cfg = self.rs485_configs[idx]
        cfg.update(
            {
                "name": name,
                "sender_port": sender_port,
                "echo_port": echo_port,
                **common_values,
            }
        )

        self._stop_rs485_worker_if_disabled(idx)
        self.refresh_rs485_row(idx)
        self.save_settings(show_message=False)
        self.append_log(f"RS485 Pair #{idx + 1} settings updated.")

    def apply_rs485_common_changes_to_all(self) -> None:
        try:
            common_values = self._read_rs485_common_editor_values()
        except ValueError as exc:
            messagebox.showerror("Invalid RS485 settings", str(exc))
            return

        answer = messagebox.askyesno(
            "Apply to all RS485",
            "Apply the selected RS485 settings to all pairs while keeping each Name, Sender Port, and Echo Port unchanged?",
        )
        if not answer:
            return

        for idx, cfg in enumerate(self.rs485_configs):
            cfg.update(common_values)
            self._stop_rs485_worker_if_disabled(idx)
            self.refresh_rs485_row(idx)

        self.save_settings(show_message=False)
        self.append_log("Applied RS485 common settings to all pairs (Name/Ports kept).")

    def _stop_rs485_worker_if_disabled(self, idx: int) -> None:
        cfg = self.rs485_configs[idx]
        has_ports = bool(str(cfg["sender_port"]).strip()) and bool(str(cfg["echo_port"]).strip())
        is_disabled = (not cfg["enabled"]) or (not has_ports)
        if not is_disabled:
            return

        self.stop_single_test("rs485", idx, log_event=False)

    def _is_rs232_startable(self, idx: int) -> bool:
        cfg = self.rs232_configs[idx]
        return bool(cfg["enabled"]) and bool(str(cfg["port"]).strip())

    def _is_rs485_startable(self, idx: int) -> bool:
        cfg = self.rs485_configs[idx]
        return bool(cfg["enabled"]) and bool(str(cfg["sender_port"]).strip()) and bool(str(cfg["echo_port"]).strip())

    def _resolved_startup_delay(self, startup_delay_s: float | None) -> float:
        if startup_delay_s is None:
            return 2.0 if bool(self.ui_settings.get("delay_comm_start_2s", True)) else 0.0
        return max(float(startup_delay_s), 0.0)

    def start_single_test(
        self,
        group: str,
        idx: int,
        startup_delay_s: float | None = None,
        reset_counts: bool = True,
        log_event: bool = True,
    ) -> None:
        delay = self._resolved_startup_delay(startup_delay_s)

        if group == "rs232":
            if idx in self.rs232_workers:
                worker = self.rs232_workers.pop(idx)
                worker.stop()
                worker.join(timeout=0.3)

            cfg = self.rs232_configs[idx]
            state = self.rs232_state[idx]
            if reset_counts:
                state["pass_count"] = 0
                state["fail_count"] = 0

            if self._is_rs232_startable(idx):
                state["status"] = "Starting"
                state["last"] = "Waiting for worker"
                worker_cfg = cfg.copy()
                worker_cfg["startup_delay_s"] = delay
                worker = RS232Worker(idx, worker_cfg, self.event_queue)
                self.rs232_workers[idx] = worker
                worker.start()
                if log_event:
                    self.append_log(f"RS232 #{idx + 1} started.")
            else:
                state["status"] = "Disabled"
                state["last"] = "Disabled in settings" if not cfg["enabled"] else "Disabled (no port selected)"
                if log_event:
                    self.append_log(f"RS232 #{idx + 1} not started ({state['last']}).")

            self.refresh_rs232_row(idx)

        elif group == "rs485":
            if idx in self.rs485_workers:
                worker = self.rs485_workers.pop(idx)
                worker.stop()
                worker.join(timeout=0.3)

            cfg = self.rs485_configs[idx]
            state = self.rs485_state[idx]
            if reset_counts:
                state["pass_count"] = 0
                state["fail_count"] = 0

            if self._is_rs485_startable(idx):
                state["status"] = "Starting"
                state["last"] = "Waiting for worker"
                worker_cfg = cfg.copy()
                worker_cfg["startup_delay_s"] = delay
                worker = RS485PairWorker(idx, worker_cfg, self.event_queue)
                self.rs485_workers[idx] = worker
                worker.start()
                if log_event:
                    self.append_log(f"RS485 Pair #{idx + 1} started.")
            else:
                state["status"] = "Disabled"
                state["last"] = "Disabled in settings" if not cfg["enabled"] else "Disabled (missing sender/echo port)"
                if log_event:
                    self.append_log(f"RS485 Pair #{idx + 1} not started ({state['last']}).")

            self.refresh_rs485_row(idx)
        else:
            raise ValueError(f"Unsupported group: {group}")

        self._refresh_health_panel()

    def stop_single_test(self, group: str, idx: int, log_event: bool = True) -> None:
        if group == "rs232":
            worker = self.rs232_workers.pop(idx, None)
            if worker is not None:
                worker.stop()
                worker.join(timeout=0.3)

            cfg = self.rs232_configs[idx]
            state = self.rs232_state[idx]
            if self._is_rs232_startable(idx):
                state["status"] = "Stopped"
                state["last"] = "Stopped by user"
            else:
                state["status"] = "Disabled"
                state["last"] = "Disabled in settings" if not cfg["enabled"] else "Disabled (no port selected)"
            self.refresh_rs232_row(idx)
            if log_event and worker is not None:
                self.append_log(f"RS232 #{idx + 1} stopped.")

        elif group == "rs485":
            worker = self.rs485_workers.pop(idx, None)
            if worker is not None:
                worker.stop()
                worker.join(timeout=0.3)

            cfg = self.rs485_configs[idx]
            state = self.rs485_state[idx]
            if self._is_rs485_startable(idx):
                state["status"] = "Stopped"
                state["last"] = "Stopped by user"
            else:
                state["status"] = "Disabled"
                state["last"] = "Disabled in settings" if not cfg["enabled"] else "Disabled (missing sender/echo port)"
            self.refresh_rs485_row(idx)
            if log_event and worker is not None:
                self.append_log(f"RS485 Pair #{idx + 1} stopped.")
        else:
            raise ValueError(f"Unsupported group: {group}")

        self._refresh_health_panel()

    def start_all_tests(self, startup_delay_s: float | None = None) -> None:
        self.start_rs232_tests(startup_delay_s=startup_delay_s)
        self.start_rs485_tests(startup_delay_s=startup_delay_s)

    def stop_all_tests(self) -> None:
        self.stop_rs232_tests()
        self.stop_rs485_tests()

    def start_rs232_tests(self, startup_delay_s: float | None = None) -> None:
        self.stop_rs232_tests()
        for idx in range(DEFAULT_RS232_COUNT):
            self.start_single_test("rs232", idx, startup_delay_s=startup_delay_s, reset_counts=True, log_event=False)

        self.append_log("RS232 test workers started.")

    def stop_rs232_tests(self) -> None:
        if not self.rs232_workers:
            return

        for idx in list(self.rs232_workers.keys()):
            self.stop_single_test("rs232", idx, log_event=False)

        self.append_log("RS232 test workers stopped.")

    def start_rs485_tests(self, startup_delay_s: float | None = None) -> None:
        self.stop_rs485_tests()
        for idx in range(DEFAULT_RS485_PAIR_COUNT):
            self.start_single_test("rs485", idx, startup_delay_s=startup_delay_s, reset_counts=True, log_event=False)

        self.append_log("RS485 pair workers started.")

    def stop_rs485_tests(self) -> None:
        if not self.rs485_workers:
            return

        for idx in list(self.rs485_workers.keys()):
            self.stop_single_test("rs485", idx, log_event=False)

        self.append_log("RS485 pair workers stopped.")

    def _process_worker_events(self) -> None:
        while True:
            try:
                event = self.event_queue.get_nowait()
            except queue.Empty:
                break

            group = event.get("group")
            idx = int(event.get("index", -1))
            status = str(event.get("status", ""))
            last = str(event.get("last", ""))
            pass_inc = int(event.get("pass_inc", 0))
            fail_inc = int(event.get("fail_inc", 0))
            log = bool(event.get("log", False))

            if group == "rs232" and 0 <= idx < len(self.rs232_state):
                state = self.rs232_state[idx]
                previous_status = state["status"]
                if status:
                    state["status"] = status
                state["pass_count"] += pass_inc
                state["fail_count"] += fail_inc
                if last:
                    state["last"] = last
                if fail_inc > 0:
                    self._record_failure_event()
                self._record_fault_transition("rs232", idx, previous_status, state["status"], state["last"])
                self.refresh_rs232_row(idx)
                if log:
                    self.append_log(f"RS232 #{idx + 1}: {status} - {last}")

            if group == "rs485" and 0 <= idx < len(self.rs485_state):
                state = self.rs485_state[idx]
                previous_status = state["status"]
                if status:
                    state["status"] = status
                state["pass_count"] += pass_inc
                state["fail_count"] += fail_inc
                if last:
                    state["last"] = last
                if fail_inc > 0:
                    self._record_failure_event()
                self._record_fault_transition("rs485", idx, previous_status, state["status"], state["last"])
                self.refresh_rs485_row(idx)
                if log:
                    self.append_log(f"RS485 Pair #{idx + 1}: {status} - {last}")

        self._refresh_health_panel()
        self.after(100, self._process_worker_events)

    def append_log(self, message: str) -> None:
        stamp = time.strftime("%H:%M:%S")
        line = f"[{stamp}] {message}\n"

        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, line)
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

        self.log_line_count += 1
        if self.log_line_count > 1200:
            self.log_text.configure(state=tk.NORMAL)
            self.log_text.delete("1.0", "101.0")
            self.log_text.configure(state=tk.DISABLED)
            self.log_line_count -= 100

    def save_settings(self, show_message: bool = True) -> bool:
        self._sync_presets_from_ui()
        self.ui_settings["start_fullscreen"] = bool(self.start_fullscreen_var.get())
        self.ui_settings["auto_start_after_launch_2s"] = bool(self.auto_start_launch_var.get())
        self.ui_settings["delay_comm_start_2s"] = bool(self.delay_comm_start_var.get())
        self.ui_settings["overview_compact_view"] = bool(self.overview_compact_var.get())
        self.ui_settings["presets"] = self.preset_configs
        settings = {
            "rs232_ports": self.rs232_configs,
            "rs485_pairs": self.rs485_configs,
            "ui": self.ui_settings,
        }
        try:
            save_settings_file(self.settings_path, settings)
        except OSError as exc:
            messagebox.showerror("Save failed", f"Could not save settings:\n{exc}")
            return False

        if show_message:
            messagebox.showinfo("Settings saved", f"Saved to:\n{self.settings_path}")
        return True

    def reload_settings(self) -> None:
        if self.rs232_workers or self.rs485_workers:
            answer = messagebox.askyesno(
                "Stop tests?",
                "Reloading settings will stop all running tests. Continue?",
            )
            if not answer:
                return
            self.stop_all_tests()

        self.settings = load_settings_file(self.settings_path)
        self.rs232_configs = self.settings["rs232_ports"]
        self.rs485_configs = self.settings["rs485_pairs"]
        self.ui_settings = self.settings["ui"]
        self.preset_configs = self.ui_settings["presets"]
        self.start_fullscreen_var.set(bool(self.ui_settings.get("start_fullscreen", False)))
        self.auto_start_launch_var.set(bool(self.ui_settings.get("auto_start_after_launch_2s", True)))
        self.delay_comm_start_var.set(bool(self.ui_settings.get("delay_comm_start_2s", True)))
        self.overview_compact_var.set(bool(self.ui_settings.get("overview_compact_view", True)))
        self._rebuild_overview_rows()
        for idx in range(DEFAULT_PRESET_COUNT):
            if idx < len(self.preset_name_vars):
                self.preset_name_vars[idx].set(self.preset_configs[idx]["name"])
        self.refresh_com_port_options(show_message=False)
        self._refresh_preset_button_labels()
        self._refresh_preset_port_options(show_message=False)

        self.rs232_state = [self.new_state() for _ in range(DEFAULT_RS232_COUNT)]
        self.rs485_state = [self.new_state() for _ in range(DEFAULT_RS485_PAIR_COUNT)]
        self.failure_counts.clear()
        self.channel_fault_history.clear()

        self._populate_tables()
        self._select_first_rows()
        self._refresh_health_panel()
        self.append_log("Settings reloaded from disk.")

    def on_close(self) -> None:
        self.stop_all_tests()
        self.save_settings(show_message=False)
        self.destroy()


def main() -> None:
    app = SerialTesterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
