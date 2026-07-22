import json
import queue
import tempfile
import types
import unittest
from pathlib import Path

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


class SerialTesterTests(unittest.TestCase):
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
        self.assertEqual([event["status"] for event in emitted], ["Running", "Stopped"])
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


if __name__ == "__main__":
    unittest.main()
