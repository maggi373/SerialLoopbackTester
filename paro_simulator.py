from __future__ import annotations

import re
import time


class ParoSimulator:
    """Stateful Paroscientific Digiquartz protocol simulator.

    This is a Python port of the Arduino implementation in F:\\Github\\parosim.
    Feed it serial bytes and write each returned byte string back to the port.
    """

    FRAME_TIMEOUT_S = 0.250
    MAX_FRAME_BYTES = 223
    PRESSURE_MIN = 100.0
    PRESSURE_MAX = 3000.0
    PRESSURE_STEP = 100.0
    TEMPERATURE_MIN_C = 10.0
    TEMPERATURE_MAX_C = 50.0
    TEMPERATURE_STEP_C = 1.0
    PRESSURE_PERIOD_US = 28.123456
    TEMPERATURE_PERIOD_US = 5.1234567
    FULL_SCALE_PSI = 16.0
    VALID_BAUDRATES = {300, 600, 1200, 2400, 4800, 9600, 19200, 38400, 57600, 115200}
    MEASUREMENT_COMMANDS = {
        "P1", "P2", "P3", "P4", "P5", "P6", "P7",
        "Q1", "Q2", "Q3", "Q4", "Q5", "Q6",
        "E1", "E2", "E3", "E4", "E5", "E6",
        "DB", "DS", "OF", "OF1", "OF2", "OF3", "OFR",
    }
    COEFFICIENT_KEYS = ("C1", "C2", "C3", "D1", "D2", "T1", "T2", "T3", "T4", "T5", "U0", "Y1", "Y2", "Y3")
    INTEGER_DEFAULTS = {
        "PR": 238, "TR": 952, "PI": 666, "TI": 666,
        "UN": 1, "TU": 0, "FM": 0, "OI": 1, "PS": 0, "MD": 1,
        "SL": 0, "ST": 10, "ZE": 0, "ZS": 0, "ZL": 0,
        "US": 0, "SU": 0, "ZI": 0, "DL": 0, "DM": 0, "DO": 0,
        "DP": 6, "IA": 11, "XM": 0, "XN": 0, "BL": 0,
    }

    def __init__(self, device_id: int = 1, link_test_mode: bool = True) -> None:
        self.device_id = max(0, min(int(device_id), 99))
        self.link_test_mode = bool(link_test_mode)
        self.pending_baud: int | None = None
        self.write_enabled = False
        self.ramp_pressure = self.PRESSURE_MIN
        self.ramp_temperature_c = self.TEMPERATURE_MIN_C
        self.pressure_direction = 1
        self.temperature_direction = 1
        self.parameters = self.INTEGER_DEFAULTS.copy()
        self.th = 0
        self.po = 0
        self.uf = 1.0
        self.pa_psi = 0.0
        self.pm = 1.0
        self.tare_psi = 0.0
        self.op_psi = 16.0
        self.coefficients = {key: (228.1234 if key == "C1" else 0.0) for key in self.COEFFICIENT_KEYS}
        self.has_min_max = False
        self.minimum_pressure_psi = 0.0
        self.maximum_pressure_psi = 0.0
        self.held_type = ""
        self.held_value = 0.0
        self.unit_label = "user"
        self.stored_text = " " * 11
        self.direct_text = ""
        self.output_masks = {
            1: "STAR,HA,UA,P,CRLF",
            2: "HEAD,SPC,P,SPC,PU,CRLF",
            3: 'STAR,HA,UA,"Pressure:",P,SPC,PU,";PPeriod:",PPER,";Temp:",ST,SPC,TU,";TPeriod:",TPER,CRLF',
        }
        self.receive_buffer = bytearray()
        self.receive_overflow = False
        self.waiting_for_lf = False
        self.last_received_at: float | None = None
        self.continuous_mode = ""
        self.continuous_host = "00"
        self.last_continuous_at: float | None = None
        self._output: list[bytes] = []

    def feed(self, data: bytes, now: float | None = None) -> list[bytes]:
        current = time.monotonic() if now is None else now
        for value in data:
            self.last_received_at = current
            self._accept_byte(value)
        return self.poll(current)

    def poll(self, now: float | None = None) -> list[bytes]:
        current = time.monotonic() if now is None else now
        if (
            (self.receive_buffer or self.receive_overflow or self.waiting_for_lf)
            and self.last_received_at is not None
            and current - self.last_received_at >= self.FRAME_TIMEOUT_S
        ):
            self._reject_received(97)
        self._update_continuous(current)
        output, self._output = self._output, []
        return output

    def take_pending_baud(self) -> int | None:
        baud, self.pending_baud = self.pending_baud, None
        return baud

    def _accept_byte(self, value: int) -> None:
        process_again = True
        while process_again:
            process_again = False
            if self.waiting_for_lf:
                if value == 0x0A:
                    self._complete_received()
                    return
                self._reject_received(95)
                process_again = True
                continue
            if value == 0x0A:
                self._reject_received(95)
                return
            if value == 0x0D:
                self.waiting_for_lf = True
                return
            if value == ord("*"):
                if self.receive_buffer or self.receive_overflow:
                    if self.receive_overflow:
                        self._send_malformed("", 96)
                    else:
                        frame = self.receive_buffer.decode("latin-1")
                        if self._is_concatenated_ew(frame):
                            self._process_frame(frame)
                        else:
                            self._send_malformed(frame, 95)
                    self._clear_receive()
                self.receive_buffer.append(value)
                return
            if not self.receive_overflow:
                if len(self.receive_buffer) < self.MAX_FRAME_BYTES:
                    self.receive_buffer.append(value)
                else:
                    self.receive_buffer.clear()
                    self.receive_overflow = True

    def _clear_receive(self) -> None:
        self.receive_buffer.clear()
        self.receive_overflow = False
        self.waiting_for_lf = False

    def _reject_received(self, error_code: int) -> None:
        if self.receive_overflow:
            self._send_malformed("", 96)
        else:
            self._send_malformed(self.receive_buffer.decode("latin-1"), error_code)
        self._clear_receive()

    def _complete_received(self) -> None:
        if self.receive_overflow:
            self._send_malformed("", 96)
        else:
            self._process_frame(self.receive_buffer.decode("latin-1"))
        self._clear_receive()

    @staticmethod
    def _is_concatenated_ew(frame: str) -> bool:
        return len(frame) == 7 and frame.startswith("*") and frame[5:] == "EW"

    @staticmethod
    def _two_digits(text: str) -> int | None:
        return int(text) if len(text) == 2 and text.isascii() and text.isdigit() else None

    def _header(self, host_id: str) -> str:
        return f"*{host_id}{self.device_id:02d}"

    def _queue_text(self, text: str) -> None:
        self._output.append(text.encode("latin-1", errors="replace"))

    def _send_payload(self, host_id: str, payload: str) -> None:
        self._queue_text(f"{self._header(host_id)}{payload}\r\n")

    def _send_error(self, host_id: str, error_code: int) -> None:
        self._send_payload(host_id, str(error_code))
        self.ramp_pressure = 0.0
        self.ramp_temperature_c = 0.0
        self.pressure_direction = 1
        self.temperature_direction = 1
        self.has_min_max = False
        self.held_type = ""
        self.held_value = 0.0
        self.continuous_mode = ""
        self.last_continuous_at = None
        self.write_enabled = False
        self.parameters["ZS"] = 0
        self.tare_psi = 0.0

    def _send_malformed(self, frame: str, error_code: int) -> None:
        host_id = "00"
        if len(frame) >= 5 and self._two_digits(frame[3:5]) is not None and int(frame[3:5]) <= 98:
            host_id = frame[3:5]
        self._send_error(host_id, error_code)

    def _process_frame(self, frame: str) -> None:
        if not frame.startswith("*"):
            self._send_malformed(frame, 90)
            return
        if len(frame) < 7:
            self._send_malformed(frame, 91)
            return
        destination = self._two_digits(frame[1:3])
        if destination is None:
            self._send_malformed(frame, 92)
            return
        host_id = frame[3:5]
        source = self._two_digits(host_id)
        if source is None or source > 98:
            self._send_malformed(frame, 93)
            return
        globally_addressed = destination == 99
        if not globally_addressed and destination != self.device_id:
            if self.link_test_mode:
                self._send_error(host_id, 94)
            return
        if globally_addressed:
            self._queue_text(frame + "\r\n")

        command = frame[5:]
        self.continuous_mode = ""
        if command == "EW":
            self.write_enabled = True
        elif command == "MR":
            self.has_min_max = False
            self._send_payload(host_id, "MR>OK")
        elif command == "M1":
            value = self.minimum_pressure_psi if self.has_min_max else 0.0
            self._send_float(host_id, "M1", value * self._unit_factor(), 5)
        elif command == "M3":
            value = self.maximum_pressure_psi if self.has_min_max else 0.0
            self._send_float(host_id, "M3", value * self._unit_factor(), 5)
        elif command in self.MEASUREMENT_COMMANDS:
            self._handle_measurement(host_id, command)
        elif not self._handle_parameter(host_id, command, globally_addressed):
            self._send_error(host_id, 99)

    def _unit_factor(self) -> float:
        return {0: self.uf, 1: 1.0, 2: 68.94757, 3: 0.06894757, 4: 6.894757,
                5: 0.00689476, 6: 2.036021, 7: 51.71493, 8: 0.7030696}.get(self.parameters["UN"], 1.0)

    def _pressure_unit_label(self) -> str:
        if self.parameters["UN"] == 0:
            return self.unit_label
        return {1: "psig" if self.po == 1 else "psia", 2: "hPa", 3: "bar", 4: "kPa",
                5: "MPa", 6: "inHg", 7: "mmHg", 8: "mH2O"}.get(self.parameters["UN"], "")

    def _temperature_unit_label(self) -> str:
        return "F" if self.parameters["TU"] == 1 else "C"

    def _default_decimals(self, value_type: str) -> int:
        if self.parameters["DL"]:
            return {"pressure": 8, "temperature": 7, "pressure_period": 9, "temperature_period": 10}[value_type]
        if self.parameters["XN"] > 0:
            if value_type == "pressure":
                full_scale = self.FULL_SCALE_PSI * self._unit_factor()
                reserved = max(1, min(9, len(str(max(1, int(abs(full_scale)))))))
            else:
                reserved = {"temperature": 3, "pressure_period": 2, "temperature_period": 1}[value_type]
            return max(0, self.parameters["XN"] - reserved)
        return {"pressure": 5, "temperature": 3, "pressure_period": 6, "temperature_period": 7}[value_type]

    def _format_number(self, value: float, value_type: str, decimals: int | None = None) -> str:
        places = self._default_decimals(value_type) if decimals is None else decimals
        text = f"{value:.{places}f}"
        if self.parameters["DL"] and value >= 0:
            text = "+" + text
        return text

    def _send_integer(self, host_id: str, key: str, value: int) -> None:
        self._send_payload(host_id, f"{key}={value}")

    def _send_float(self, host_id: str, key: str, value: float, decimals: int = 6) -> None:
        self._send_payload(host_id, f"{key}={value:.{decimals}f}")

    def _acquire_pressure(self) -> float:
        factor = self._unit_factor()
        raw_psi = self.ramp_pressure / factor if factor else self.ramp_pressure
        value_psi = (raw_psi + self.pa_psi) * self.pm
        if self.parameters["ZS"] == 1:
            self.tare_psi = value_psi
            self.parameters["ZS"] = 2
        if self.parameters["ZS"] == 2:
            value_psi -= self.tare_psi
        if not self.has_min_max:
            self.minimum_pressure_psi = self.maximum_pressure_psi = value_psi
            self.has_min_max = True
        else:
            self.minimum_pressure_psi = min(self.minimum_pressure_psi, value_psi)
            self.maximum_pressure_psi = max(self.maximum_pressure_psi, value_psi)
        return value_psi * factor

    def _acquire_temperature(self) -> float:
        return self.ramp_temperature_c * 1.8 + 32.0 if self.parameters["TU"] == 1 else self.ramp_temperature_c

    @staticmethod
    def _advance(value: float, direction: int, minimum: float, maximum: float, step: float) -> tuple[float, int]:
        if direction > 0 and value >= maximum:
            direction = -1
        if direction < 0 and value <= minimum:
            direction = 1
        return max(minimum, min(maximum, value + direction * step)), direction

    def _advance_pressure(self) -> None:
        self.ramp_pressure, self.pressure_direction = self._advance(
            self.ramp_pressure, self.pressure_direction, self.PRESSURE_MIN, self.PRESSURE_MAX, self.PRESSURE_STEP
        )

    def _advance_temperature(self) -> None:
        self.ramp_temperature_c, self.temperature_direction = self._advance(
            self.ramp_temperature_c, self.temperature_direction,
            self.TEMPERATURE_MIN_C, self.TEMPERATURE_MAX_C, self.TEMPERATURE_STEP_C,
        )

    def _advance_both(self) -> None:
        self._advance_pressure()
        self._advance_temperature()

    def _measurement_text(self, host_id: str, value: float, value_type: str, add_suffix: bool = True) -> str:
        text = self._header(host_id)
        if self.parameters["SU"]:
            text += "_"
        text += self._format_number(value, value_type)
        if value_type == "pressure" and self.parameters["ZI"] and self.parameters["ZS"] == 2:
            text += "T"
        if add_suffix and self.parameters["US"]:
            if self.parameters["SU"]:
                text += "_"
            text += self._pressure_unit_label() if value_type == "pressure" else self._temperature_unit_label()
        return text + "\r\n"

    def _send_measurement(self, host_id: str, value: float, value_type: str, add_suffix: bool = True) -> None:
        self._queue_text(self._measurement_text(host_id, value, value_type, add_suffix))

    def _send_compound(self, host_id: str, kind: int) -> None:
        if kind == 1:
            values = [self._format_number(self.PRESSURE_PERIOD_US, "pressure_period"),
                      self._format_number(self.TEMPERATURE_PERIOD_US, "temperature_period")]
        elif kind == 3:
            values = [self._format_number(self._acquire_pressure(), "pressure"),
                      self._format_number(self._acquire_temperature(), "temperature")]
        else:
            values = [self._format_number(self._acquire_pressure(), "pressure"),
                      self._format_number(self.PRESSURE_PERIOD_US, "pressure_period"),
                      self._format_number(self.TEMPERATURE_PERIOD_US, "temperature_period")]
        self._queue_text(self._header(host_id) + "," + ",".join(values) + "\r\n")

    def _start_continuous(self, mode: str, host_id: str) -> None:
        self.continuous_mode = mode
        self.continuous_host = host_id
        self.last_continuous_at = None

    def _continuous_interval(self) -> float:
        if self.th > 0:
            return max(0.001, 1.0 / self.th)
        if self.continuous_mode in {"Q2", "Q3"}:
            milliseconds = self.parameters["TI"]
        elif self.parameters["OI"] == 1:
            milliseconds = self.parameters["PI"] + self.parameters["TI"]
        else:
            milliseconds = max(self.parameters["PI"], self.parameters["TI"])
        return max(0.001, milliseconds / 1000.0)

    def _update_continuous(self, now: float) -> None:
        if not self.continuous_mode:
            return
        if self.last_continuous_at is not None and now - self.last_continuous_at < self._continuous_interval():
            return
        self.last_continuous_at = now
        host_id, mode = self.continuous_host, self.continuous_mode
        if mode == "P2": self._send_measurement(host_id, self.PRESSURE_PERIOD_US, "pressure_period", False)
        elif mode == "P3": self._send_measurement(host_id, self._acquire_pressure(), "pressure")
        elif mode == "Q2": self._send_measurement(host_id, self.TEMPERATURE_PERIOD_US, "temperature_period", False)
        elif mode == "Q3": self._send_measurement(host_id, self._acquire_temperature(), "temperature")
        elif mode == "E2": self._send_compound(host_id, 1)
        elif mode == "E3": self._send_compound(host_id, 3)
        elif mode == "E5": self._send_compound(host_id, 5)
        elif mode == "OFR": self._queue_text(self._execute_mask(1, host_id))
        if mode in {"P2", "P3"}: self._advance_pressure()
        elif mode in {"Q2", "Q3"}: self._advance_temperature()
        else: self._advance_both()

    def _handle_measurement(self, host_id: str, command: str) -> None:
        if command == "P1": self._send_measurement(host_id, self.PRESSURE_PERIOD_US, "pressure_period", False); self._advance_pressure()
        elif command == "P2": self._start_continuous("P2", host_id)
        elif command == "P3": self._send_measurement(host_id, self._acquire_pressure(), "pressure"); self._advance_pressure()
        elif command in {"P4", "P7"}: self._start_continuous("P3", host_id)
        elif command == "P5": self.held_value, self.held_type = self._acquire_pressure(), "pressure"
        elif command == "P6": self.held_value, self.held_type = self.PRESSURE_PERIOD_US, "pressure_period"
        elif command == "Q1": self._send_measurement(host_id, self.TEMPERATURE_PERIOD_US, "temperature_period", False); self._advance_temperature()
        elif command == "Q2": self._start_continuous("Q2", host_id)
        elif command == "Q3": self._send_measurement(host_id, self._acquire_temperature(), "temperature"); self._advance_temperature()
        elif command == "Q4": self._start_continuous("Q3", host_id)
        elif command == "Q5": self.held_value, self.held_type = self._acquire_temperature(), "temperature"
        elif command == "Q6": self.held_value, self.held_type = self.TEMPERATURE_PERIOD_US, "temperature_period"
        elif command == "E1": self._send_compound(host_id, 1); self._advance_both()
        elif command == "E2": self._start_continuous("E2", host_id)
        elif command == "E3": self._send_compound(host_id, 3); self._advance_both()
        elif command == "E4": self._start_continuous("E3", host_id)
        elif command == "E5": self._send_compound(host_id, 5); self._advance_both()
        elif command == "E6": self._start_continuous("E5", host_id)
        elif command in {"DB", "DS"}:
            held_type = self.held_type
            if held_type:
                self._send_measurement(host_id, self.held_value, held_type, held_type in {"pressure", "temperature"})
            if held_type.startswith("pressure"): self._advance_pressure()
            elif held_type.startswith("temperature"): self._advance_temperature()
        elif command in {"OF", "OF1", "OF2", "OF3"}:
            number = int(command[-1]) if command[-1:].isdigit() else 1
            self._queue_text(self._execute_mask(number, host_id)); self._advance_both()
        elif command == "OFR": self._start_continuous("OFR", host_id)

    @staticmethod
    def _atol(text: str) -> int:
        match = re.match(r"^[\s]*([+-]?\d+)", text)
        return int(match.group(1)) if match else 0

    @staticmethod
    def _atof(text: str) -> float:
        match = re.match(r"^[\s]*([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?)", text)
        return float(match.group(1)) if match else 0.0

    def _handle_parameter(self, host_id: str, command: str, globally_addressed: bool) -> bool:
        if command.startswith("OM"):
            number = int(command[2]) if len(command) > 2 and command[2] in "123" else 1
            key_length = 3 if len(command) > 2 and command[2] in "123" else 2
            key = "OM" + (str(number) if key_length == 3 else "")
            if "=" not in command:
                if len(command) == key_length:
                    self._send_payload(host_id, f"{key}={self.output_masks[number]}")
                return True
            if not self.write_enabled:
                self._send_error(host_id, 98); return True
            self.write_enabled = False
            value = command.split("=", 1)[1]
            if value == "RESET":
                self.output_masks.update({1: "STAR,HA,UA,P,CRLF", 2: "HEAD,SPC,P,SPC,PU,CRLF",
                    3: 'STAR,HA,UA,"Pressure:",P,SPC,PU,";PPeriod:",PPER,";Temp:",ST,SPC,TU,";TPeriod:",TPER,CRLF'})
            else:
                self.output_masks[number] = value[:200]
            self._send_payload(host_id, f"{key}={self.output_masks[number]}")
            return True

        if len(command) < 2:
            return False
        key = command[:2]
        has_equals = "=" in command
        value_text = command.split("=", 1)[1] if has_equals else ""
        if key == "DT" and has_equals:
            self.direct_text = value_text[:16]; self._send_payload(host_id, command); return True
        if key == "BR" and has_equals and globally_addressed and not self.parameters["BL"]:
            requested = self._atol(value_text)
            if requested in self.VALID_BAUDRATES: self.pending_baud = requested
            return True
        if key == "ID" and not has_equals and globally_addressed and command == "ID":
            source = self._two_digits(host_id)
            if source is not None and source < 98: self.device_id = source + 1
            return True
        if has_equals:
            if not self.write_enabled:
                self._send_error(host_id, 98); return True
            self.write_enabled = False
            if key == "UM": self.unit_label = value_text[:4]; self._send_payload(host_id, f"UM={self.unit_label}"); return True
            if key == "UL": self.stored_text = value_text[:11]; self._send_payload(host_id, f"UL={self.stored_text}"); return True
            if key == "UF": self.uf = self._atof(value_text); self._send_float(host_id, key, self.uf); return True
            if key == "PA": self.pa_psi = self._atof(value_text) / self._unit_factor(); self.has_min_max = False; self._send_float(host_id, key, self.pa_psi * self._unit_factor()); return True
            if key == "PM": self.pm = self._atof(value_text); self.has_min_max = False; self._send_float(host_id, key, self.pm); return True
            if key == "ZV": self.tare_psi = self._atof(value_text) / self._unit_factor(); self._send_float(host_id, key, self.tare_psi * self._unit_factor()); return True
            if key == "OP": self.op_psi = self._atof(value_text) / self._unit_factor(); self._send_float(host_id, key, self.op_psi * self._unit_factor(), 5); return True
            if key in self.coefficients: self.coefficients[key] = self._atof(value_text); self._send_float(host_id, key, self.coefficients[key], 7); return True
            if key == "TH":
                self.th = self._atol(value_text)
                payload = f"TH={value_text};>OK" if "," in value_text else f"TH={self.th}"
                self._send_payload(host_id, payload); return True
            if key not in self.parameters: return False
            if key == "ZS" and self.parameters["ZL"]: return True
            self.parameters[key] = self._atol(value_text)
            if key == "PR": self.parameters["TR"] = self.parameters["PR"] * 4
            if key == "PI": self.parameters["TI"] = self.parameters["PI"]
            if key == "XM" and self.parameters["XM"] == 1: self.parameters["OI"] = 0
            if key == "MD": self._begin_md_continuous(host_id)
            self._send_integer(host_id, key, self.parameters[key]); return True

        if command != key:
            return False
        fixed = {"SN": "SN=00001", "VR": "VR=R5.20", "CF": "CF=A1B2",
                 "MN": "MN=DIGIQUARTZ UNO SIM      ", "PO": "PO=0", "TC": "TC=.6666667"}
        if key in fixed: self._send_payload(host_id, fixed[key]); return True
        if key == "PF": self._send_float(host_id, key, self.FULL_SCALE_PSI * self._unit_factor(), 5); return True
        if key == "PL": self._send_float(host_id, key, self.FULL_SCALE_PSI * 1.2 * self._unit_factor(), 4); return True
        if key == "UM": self._send_payload(host_id, f"UM={self.unit_label}"); return True
        if key == "UL": self._send_payload(host_id, f"UL={self.stored_text}"); return True
        if key == "UF": self._send_float(host_id, key, self.uf); return True
        if key == "PA": self._send_float(host_id, key, self.pa_psi * self._unit_factor()); return True
        if key == "PM": self._send_float(host_id, key, self.pm); return True
        if key == "ZV": self._send_float(host_id, key, self.tare_psi * self._unit_factor()); return True
        if key == "OP": self._send_float(host_id, key, self.op_psi * self._unit_factor(), 5); return True
        if key == "TH": self._send_integer(host_id, key, self.th); return True
        if key in self.coefficients: self._send_float(host_id, key, self.coefficients[key], 7); return True
        if key in self.parameters: self._send_integer(host_id, key, self.parameters[key]); return True
        return False

    def _begin_md_continuous(self, host_id: str) -> None:
        md = self.parameters["MD"]
        if md in {2, 3}: self._start_continuous("P3", host_id)
        elif md == 14: self._start_continuous("E3", host_id)
        elif md == 15: self._start_continuous("E5", host_id)
        elif md in {8, 10, 12}: self._start_continuous("OFR", host_id)

    def _mask_number(self, value: float, value_type: str, format_text: str, force_sign: bool = False) -> str:
        decimals = self._default_decimals(value_type)
        minimum_digits = 0
        match = re.match(r"(\d+)", format_text)
        if match: minimum_digits = int(match.group(1))
        decimal_match = re.search(r"\.(\d+)", format_text)
        if decimal_match: decimals = int(decimal_match.group(1))
        text = f"{abs(value):.{decimals}f}"
        integer, dot, fraction = text.partition(".")
        text = integer.zfill(minimum_digits) + (dot + fraction if dot else "")
        return ("-" if value < 0 else "+" if force_sign else "") + text

    def _execute_mask(self, number: int, host_id: str) -> str:
        mask = self.output_masks.get(number, "")
        tokens = re.findall(r'"[^"]*"|[^,]+', mask)
        output = ""
        for raw_token in tokens:
            token = raw_token.strip()
            if token.startswith('"') and token.endswith('"'):
                output += token[1:-1]; continue
            upper = token.upper()
            if upper == "STAR": output += "*"
            elif upper == "UA": output += host_id
            elif upper == "HA": output += f"{self.device_id:02d}"
            elif upper == "HEAD": output += self._header(host_id)
            elif upper == "CR": output += "\r"
            elif upper == "LF": output += "\n"
            elif upper in {"CRLF", "END", "E"}: output += "\r\n"
            elif upper in {"SPACE", "SPC"}: output += " "
            elif upper == "COMMA": output += ","
            elif upper == "PU": output += self._pressure_unit_label()
            elif upper == "TU": output += self._temperature_unit_label()
            elif upper.startswith("MINP"): output += self._mask_number((self.minimum_pressure_psi if self.has_min_max else 0.0) * self._unit_factor(), "pressure", upper[4:])
            elif upper.startswith("MAXP"): output += self._mask_number((self.maximum_pressure_psi if self.has_min_max else 0.0) * self._unit_factor(), "pressure", upper[4:])
            elif upper.startswith("TV"): output += self._mask_number(self.tare_psi * self._unit_factor(), "pressure", upper[2:])
            elif upper.startswith("PPER"): output += self._mask_number(self.PRESSURE_PERIOD_US, "pressure_period", upper[4:])
            elif upper.startswith("TPER"): output += self._mask_number(self.TEMPERATURE_PERIOD_US, "temperature_period", upper[4:])
            elif upper.startswith("ST"): output += self._mask_number(self._acquire_temperature(), "temperature", upper[2:])
            elif upper.startswith("P"):
                plus = upper.startswith("P+")
                output += self._mask_number(self._acquire_pressure(), "pressure", upper[2:] if plus else upper[1:], plus)
        return output.encode("latin-1", errors="replace").decode("latin-1")
