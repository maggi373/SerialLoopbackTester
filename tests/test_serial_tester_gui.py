import json
import os
import queue
import tempfile
import types
import unittest
from pathlib import Path
from unittest.mock import patch

import serial_tester_gui as app_module


class StopDuringReadPort:
    def __init__(self, worker):
        self.worker = worker

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False

    def reset_input_buffer(self):
        pass

    def reset_output_buffer(self):
        pass

    def write(self, payload):
        return len(payload)

    def flush(self):
        pass

    def read(self, _length):
        self.worker.stop()
        return b""


class StopDuringReadWorker(app_module.RS232Worker):
    def open_port(self):
        return StopDuringReadPort(self)


class ScriptedPort:
    def __init__(self, reads):
        self.reads = list(reads)
        self.writes = []

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False

    def reset_input_buffer(self):
        pass

    def reset_output_buffer(self):
        pass

    def write(self, payload):
        self.writes.append(payload)
        return len(payload)

    def flush(self):
        pass

    def read(self, _length):
        return self.reads.pop(0) if self.reads else b""


class StopAfterRs232EchoWorker(app_module.RS232Worker):
    def __init__(self, *args, port, **kwargs):
        super().__init__(*args, **kwargs)
        self.port = port

    def open_port(self):
        return self.port

    def emit(self, status, last, **kwargs):
        super().emit(status, last, **kwargs)
        if last.startswith("RS485 reply RX/TX"):
            self.stop_event.set()


class StopAfterRs232FailureWorker(StopAfterRs232EchoWorker):
    def emit(self, status, last, **kwargs):
        app_module.RS232Worker.emit(self, status, last, **kwargs)
        if status == "FAIL":
            self.stop_event.set()


class StopAfterRs232GraceWorker(StopAfterRs232EchoWorker):
    def emit(self, status, last, **kwargs):
        app_module.RS232Worker.emit(self, status, last, **kwargs)
        if last.startswith("Grace period:"):
            self.stop_event.set()


class StopAfterRs485PassWorker(app_module.RS485PairWorker):
    def __init__(self, *args, port, **kwargs):
        super().__init__(*args, **kwargs)
        self.port = port
        self.opened_names = []

    def open_port(self, port_name):
        self.opened_names.append(port_name)
        return self.port

    def emit(self, status, last, **kwargs):
        super().emit(status, last, **kwargs)
        if status == "PASS":
            self.stop_event.set()


class StopAfterRs485GraceWorker(StopAfterRs485PassWorker):
    def emit(self, status, last, **kwargs):
        app_module.RS485PairWorker.emit(self, status, last, **kwargs)
        if last.startswith("Grace period:"):
            self.stop_event.set()


class StopAfterParoResponseWorker(app_module.RS232Worker):
    def __init__(self, *args, port, **kwargs):
        super().__init__(*args, **kwargs)
        self.port = port

    def open_port(self):
        return self.port

    def emit(self, status, last, **kwargs):
        super().emit(status, last, **kwargs)
        if status == "PASS":
            self.stop_event.set()


class FailToOpenWorker(app_module.RS232Worker):
    def open_port(self):
        raise app_module.SerialException("access denied")

    def emit(self, status, last, **kwargs):
        super().emit(status, last, **kwargs)
        if status == "ERROR":
            self.stop_event.set()


class SerialTesterTests(unittest.TestCase):
    def test_port_normalization_preserves_case_sensitive_linux_device_paths(self):
        self.assertEqual(app_module.normalize_port_text("/dev/ttyUSB0"), "/dev/ttyUSB0")
        self.assertEqual(app_module.normalize_port_text("/dev/serial/by-id/MyAdapter"), "/dev/serial/by-id/MyAdapter")
        self.assertEqual(app_module.normalize_port_text("com17"), "COM17")
        self.assertEqual(app_module.normalize_port_text("socket://192.0.2.10:4001"), "socket://192.0.2.10:4001")

    def test_socket_endpoint_uses_pyserial_url_handler(self):
        opened = object()
        options = {
            "baudrate": 9600,
            "bytesize": 8,
            "parity": "N",
            "stopbits": 1.0,
            "timeout": 1.5,
            "write_timeout": 1.5,
        }
        with (
            patch.object(app_module.serial, "serial_for_url", return_value=opened) as url_open,
            patch.object(app_module.serial, "Serial") as local_open,
        ):
            result = app_module.open_serial_endpoint("SOCKET://192.0.2.10:4001", **options)

        self.assertIs(result, opened)
        url_open.assert_called_once_with("socket://192.0.2.10:4001", **options)
        local_open.assert_not_called()

    def test_local_endpoint_uses_native_serial_port(self):
        opened = object()
        options = {
            "baudrate": 19200,
            "bytesize": 8,
            "parity": "E",
            "stopbits": 1.0,
            "timeout": 2.0,
            "write_timeout": 2.0,
        }
        with (
            patch.object(app_module.serial, "serial_for_url") as url_open,
            patch.object(app_module.serial, "Serial", return_value=opened) as local_open,
        ):
            result = app_module.open_serial_endpoint("/dev/ttyUSB0", **options)

        self.assertIs(result, opened)
        local_open.assert_called_once_with(port="/dev/ttyUSB0", **options)
        url_open.assert_not_called()

    def test_unsupported_serial_url_explains_raw_tcp_format(self):
        with self.assertRaisesRegex(ValueError, "socket://host:port"):
            app_module.open_serial_endpoint(
                "rfc2217://192.0.2.10:4001",
                baudrate=9600,
                bytesize=8,
                parity="N",
                stopbits=1.0,
                timeout=1.0,
                write_timeout=1.0,
            )

    def test_linux_config_folder_honors_xdg_config_home(self):
        config_home = str(Path.cwd() / "test-config")
        with patch.dict(os.environ, {"XDG_CONFIG_HOME": config_home}, clear=False):
            self.assertEqual(
                app_module.resolve_linux_config_folder(),
                Path(config_home) / app_module.APP_FOLDER_NAME,
            )

    def test_rs232_paro_role_passively_answers_instead_of_sending_loopback_payload(self):
        events = queue.Queue()
        config = app_module.default_rs232_item(0)
        config.update({"mode": app_module.RS232_MODE_PARO, "paro_device_id": 1, "baudrate": 9600})
        port = ScriptedPort([b"*0100P3\r\n"])
        worker = StopAfterParoResponseWorker(0, config, events, worker_id=7, port=port)

        worker.run()

        self.assertEqual(port.writes, [b"*0001100.00000\r\n"])
        self.assertIn("PARO TX *0001100.00000", [event["last"] for event in events.queue])

    def test_rs232_paro_assignment_is_normalized_and_persisted(self):
        config = app_module.normalize_rs232({"mode": "paro", "paro_device_id": 42}, 0)

        self.assertEqual(config["mode"], app_module.RS232_MODE_PARO)
        self.assertEqual(config["paro_device_id"], 42)
        expected = f'{config["port"]} · PARO ID 42' if config["port"] else "PARO ID 42"
        self.assertEqual(app_module.rs232_port_role_text(config), expected)

    def test_rs485_reply_assignment_is_normalized_and_persisted(self):
        config = app_module.normalize_rs232({"mode": "rs485_reply"}, 0)

        self.assertEqual(config["mode"], app_module.RS232_MODE_RS485_REPLY)
        self.assertEqual(app_module.rs232_mode_label(config["mode"]), "RS485 Reply")

    def test_rs485_reply_role_is_passive_and_echoes_whatever_it_receives(self):
        events = queue.Queue()
        config = app_module.default_rs232_item(0)
        config["mode"] = app_module.RS232_MODE_RS485_REPLY
        rs485_payload = bytes.fromhex(app_module.DEFAULT_RS485_PAYLOAD_HEX)
        port = ScriptedPort([rs485_payload])
        worker = StopAfterRs232EchoWorker(0, config, events, worker_id=3, port=port)

        worker.run()

        self.assertEqual(port.writes, [rs485_payload])
        self.assertTrue(any(event["last"].startswith("RS485 reply RX/TX") for event in events.queue))
        self.assertEqual(sum(event["fail_inc"] for event in events.queue), 0)

    def test_loopback_role_does_not_echo_unexpected_bytes(self):
        events = queue.Queue()
        config = app_module.default_rs232_item(0)
        config["failure_grace_period_s"] = 0
        unexpected = bytes.fromhex("DEADBEEFDEADBEEF")
        port = ScriptedPort([unexpected])
        worker = StopAfterRs232FailureWorker(0, config, events, worker_id=6, port=port)

        worker.run()

        self.assertEqual(port.writes, [bytes.fromhex(config["payload_hex"])])
        self.assertEqual(sum(event["fail_inc"] for event in events.queue), 1)

    def test_rs232_failures_are_ignored_during_two_second_grace_period(self):
        events = queue.Queue()
        config = app_module.default_rs232_item(0)
        unexpected = bytes.fromhex("DEADBEEFDEADBEEF")
        port = ScriptedPort([unexpected])
        worker = StopAfterRs232GraceWorker(0, config, events, worker_id=8, port=port)

        worker.run()

        self.assertEqual(sum(event["fail_inc"] for event in events.queue), 0)
        self.assertTrue(any("failure ignored" in event["last"] for event in events.queue))

    def test_rs485_failures_are_ignored_during_two_second_grace_period(self):
        events = queue.Queue()
        config = app_module.default_rs485_item(0)
        unexpected = bytes.fromhex("DEADBEEFDEADBEEF")
        port = ScriptedPort([unexpected])
        worker = StopAfterRs485GraceWorker(0, config, events, worker_id=9, port=port)

        worker.run()

        self.assertEqual(sum(event["fail_inc"] for event in events.queue), 0)
        self.assertTrue(any("failure ignored" in event["last"] for event in events.queue))

    def test_rs485_uses_one_port_and_accepts_echo_returned_to_it(self):
        events = queue.Queue()
        config = app_module.default_rs485_item(0)
        payload = bytes.fromhex(config["payload_hex"])
        port = ScriptedPort([payload])
        worker = StopAfterRs485PassWorker(0, config, events, worker_id=4, port=port)

        worker.run()

        self.assertEqual(worker.opened_names, [config["sender_port"]])
        self.assertEqual(port.writes, [payload])
        self.assertIn("PASS", [event["status"] for event in events.queue])
        self.assertNotIn("echo_port", config)

    def test_port_open_failure_identifies_port_and_cause(self):
        events = queue.Queue()
        config = app_module.default_rs232_item(0)
        config["port"] = "COM77"
        worker = FailToOpenWorker(0, config, events, worker_id=5)

        worker.run()

        error = next(event for event in events.queue if event["status"] == "ERROR")
        self.assertIn("PORT OPEN FAILED (COM77)", error["last"])
        self.assertIn("access denied", error["last"])
        self.assertEqual(error["error_inc"], 1)

    def test_overview_distinguishes_port_errors_from_message_failures(self):
        app = object.__new__(app_module.SerialTesterApp)
        app.channel_fault_history = set()

        self.assertEqual(app._status_to_overview_state("rs232", 0, "ERROR")[1], "Port Error")
        self.assertEqual(app._status_to_overview_state("rs232", 0, "FAIL")[1], "Wrong Message")

    def test_stop_during_read_does_not_record_failure(self):
        events = queue.Queue()
        worker = StopDuringReadWorker(
            0,
            app_module.default_rs232_item(0),
            events,
            worker_id=17,
        )

        worker.run()

        emitted = list(events.queue)
        self.assertEqual(emitted[0]["status"], "Running")
        self.assertEqual(emitted[-1]["status"], "Stopped")
        self.assertEqual(sum(event["fail_inc"] for event in emitted), 0)
        self.assertTrue(all(event["worker_id"] == 17 for event in emitted))

    def test_invalid_settings_are_backed_up_before_defaults_are_written(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "settings.json"
            path.write_text("{broken", encoding="utf-8")

            settings = app_module.load_settings_file(path)

            self.assertEqual(settings["rs232_ports"][0]["baudrate"], app_module.DEFAULT_BAUDRATE)
            self.assertTrue(list(path.parent.glob("settings.invalid-*.json.bak")))
            json.loads(path.read_text(encoding="utf-8"))

    def test_normalization_restores_last_preset_and_rejects_invalid_serial_values(self):
        settings = app_module.normalize_settings(
            {
                "ui": {
                    "rs232_count": 1,
                    "rs485_pair_count": 0,
                    "active_preset_idx": 3,
                    "overview_hide_non_preset_ports": True,
                },
                "rs232_ports": [{"baudrate": -1, "timeout_s": "nan"}],
            }
        )

        self.assertEqual(settings["ui"]["active_preset_idx"], 3)
        self.assertTrue(settings["ui"]["overview_hide_non_preset_ports"])
        self.assertEqual(settings["rs232_ports"][0]["baudrate"], 1)
        self.assertEqual(settings["rs232_ports"][0]["timeout_s"], 0.5)

    def test_preset_rs232_roles_are_normalized_and_invalid_roles_are_dropped(self):
        preset = app_module.normalize_preset_item(
            {
                "name": "Reply setup",
                "names": ["Reply Port"],
                "rs232_roles": {
                    "Reply Port": "RS485 Reply",
                    "Ignored Port": "invalid-role",
                },
            },
            0,
        )

        self.assertEqual(preset["rs232_roles"], {"Reply Port": app_module.RS232_MODE_RS485_REPLY})

    def test_stale_worker_events_are_ignored_after_restart(self):
        app = object.__new__(app_module.SerialTesterApp)
        app.event_queue = queue.Queue()
        app.event_queue.put(
            {
                "group": "rs232",
                "index": 0,
                "worker_id": 10,
                "status": "FAIL",
                "last": "old worker",
                "pass_inc": 0,
                "fail_inc": 1,
                "error_inc": 0,
                "log": True,
            }
        )
        app.event_queue.put(
            {
                "group": "rs232",
                "index": 0,
                "worker_id": 11,
                "status": "PASS",
                "last": "new worker",
                "pass_inc": 1,
                "fail_inc": 0,
                "error_inc": 0,
                "log": False,
            }
        )
        current_worker = types.SimpleNamespace(worker_id=11)
        app.rs232_workers = {0: current_worker}
        app.rs485_workers = {}
        app.rs232_state = [app_module.SerialTesterApp.new_state()]
        app.rs485_state = []
        app.window_motion_active = False
        app.after = lambda *_args, **_kwargs: None
        app._record_failure_event = lambda: None
        app._record_fault_transition = lambda *_args: None
        app._queue_live_refresh = lambda *_args: None
        app._append_worker_event_log = lambda *_args: self.fail("stale event was logged")

        app_module.SerialTesterApp._process_worker_events(app)

        self.assertEqual(app.rs232_state[0]["status"], "PASS")
        self.assertEqual(app.rs232_state[0]["pass_count"], 1)
        self.assertEqual(app.rs232_state[0]["fail_count"], 0)

    def test_running_preset_switch_stops_old_name_and_starts_new_name(self):
        app = object.__new__(app_module.SerialTesterApp)
        app.rs232_configs = [app_module.default_rs232_item(0), app_module.default_rs232_item(1)]
        app.rs232_configs[0].update({"name": "Profile A", "port": "COM7"})
        app.rs232_configs[1].update({"name": "Profile B", "port": "COM7"})
        app.rs485_configs = []
        app.preset_configs = [app_module.default_preset_item(i) for i in range(app_module.DEFAULT_PRESET_COUNT)]
        app.preset_configs[1]["names"] = ["Profile B"]
        app.preset_configs[1]["rs232_roles"] = {"Profile B": app_module.RS232_MODE_RS485_REPLY}
        app.preset_name_vars = []
        app.preset_name_listboxes = []
        app.preset_buttons = []
        app.preset_panels = []
        app.ui_settings = {}
        app.rs232_workers = {0: object()}
        app.rs485_workers = {}
        calls = []

        def stop_single_test(self, group, idx, **_kwargs):
            calls.append(("stop", group, idx))
            self.rs232_workers.pop(idx, None)

        def start_single_test(self, group, idx, **_kwargs):
            calls.append(("start", group, idx))
            self.rs232_workers[idx] = object()

        app.stop_single_test = types.MethodType(stop_single_test, app)
        app.start_single_test = types.MethodType(start_single_test, app)
        app.refresh_rs232_row = lambda *_args, **_kwargs: None
        app.refresh_rs485_row = lambda *_args, **_kwargs: None
        app.save_settings = lambda **_kwargs: True
        app._rebuild_overview_rows = lambda: None
        app._refresh_health_panel = lambda: None
        app.append_log = lambda _message: None

        app_module.SerialTesterApp.apply_preset(app, 1)

        self.assertFalse(app.rs232_configs[0]["enabled"])
        self.assertTrue(app.rs232_configs[1]["enabled"])
        self.assertEqual(app.rs232_configs[1]["mode"], app_module.RS232_MODE_RS485_REPLY)
        self.assertEqual(calls, [("stop", "rs232", 0), ("start", "rs232", 1)])
        self.assertEqual(app.active_preset_idx, 1)
        self.assertEqual(app.ui_settings["active_preset_idx"], 1)

    def test_idle_preset_switch_does_not_start_workers(self):
        app = object.__new__(app_module.SerialTesterApp)
        app.rs232_configs = [app_module.default_rs232_item(0)]
        app.rs232_configs[0]["name"] = "Profile A"
        app.rs485_configs = []
        app.preset_configs = [app_module.default_preset_item(i) for i in range(app_module.DEFAULT_PRESET_COUNT)]
        app.preset_configs[0]["names"] = ["Profile A"]
        app.preset_name_vars = []
        app.preset_name_listboxes = []
        app.preset_buttons = []
        app.preset_panels = []
        app.ui_settings = {}
        app.rs232_workers = {}
        app.rs485_workers = {}
        starts = []
        app.stop_single_test = lambda *_args, **_kwargs: None
        app.start_single_test = lambda *args, **_kwargs: starts.append(args)
        app.refresh_rs232_row = lambda *_args, **_kwargs: None
        app.refresh_rs485_row = lambda *_args, **_kwargs: None
        app.save_settings = lambda **_kwargs: True
        app._rebuild_overview_rows = lambda: None
        app._refresh_health_panel = lambda: None
        app.append_log = lambda _message: None

        app_module.SerialTesterApp.apply_preset(app, 0)

        self.assertEqual(starts, [])
        self.assertTrue(app.rs232_configs[0]["enabled"])

    def test_preset_role_change_restarts_a_running_rs232_channel(self):
        app = object.__new__(app_module.SerialTesterApp)
        app.rs232_configs = [app_module.default_rs232_item(0)]
        app.rs232_configs[0].update({"name": "Reply Port", "port": "COM7"})
        app.rs485_configs = []
        app.preset_configs = [app_module.default_preset_item(i) for i in range(app_module.DEFAULT_PRESET_COUNT)]
        app.preset_configs[0].update(
            {
                "names": ["Reply Port"],
                "rs232_roles": {"Reply Port": app_module.RS232_MODE_RS485_REPLY},
            }
        )
        app.preset_name_vars = []
        app.preset_name_listboxes = []
        app.preset_buttons = []
        app.preset_panels = []
        app.ui_settings = {}
        app.rs232_workers = {0: object()}
        app.rs485_workers = {}
        calls = []

        def stop_single_test(self, group, idx, **_kwargs):
            calls.append(("stop", group, idx))
            self.rs232_workers.pop(idx, None)

        def start_single_test(self, group, idx, **_kwargs):
            calls.append(("start", group, idx))
            self.rs232_workers[idx] = object()

        app.stop_single_test = types.MethodType(stop_single_test, app)
        app.start_single_test = types.MethodType(start_single_test, app)
        app.refresh_rs232_row = lambda *_args, **_kwargs: None
        app.refresh_rs485_row = lambda *_args, **_kwargs: None
        app.save_settings = lambda **_kwargs: True
        app._rebuild_overview_rows = lambda: None
        app._refresh_health_panel = lambda: None
        app.append_log = lambda _message: None

        app_module.SerialTesterApp.apply_preset(app, 0)

        self.assertEqual(app.rs232_configs[0]["mode"], app_module.RS232_MODE_RS485_REPLY)
        self.assertEqual(calls, [("stop", "rs232", 0), ("start", "rs232", 0)])


if __name__ == "__main__":
    unittest.main()
