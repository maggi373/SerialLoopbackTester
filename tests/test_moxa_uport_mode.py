import struct
import unittest
from unittest.mock import patch

import moxa_uport_mode as helper


class MoxaUportModeTests(unittest.TestCase):
    def test_mode_aliases_match_moxa_interface_values(self):
        self.assertEqual(helper.normalize_mode("rs232"), "rs232")
        self.assertEqual(helper.normalize_mode("RS-485"), "rs485-2w")
        self.assertEqual(helper.normalize_mode("3"), "rs485-4w")
        self.assertEqual(helper.MODE_CODES["rs485-2w"], 1)
        self.assertEqual(helper.MODE_CODES["rs422"], 2)
        self.assertEqual(helper.MODE_CODES["rs485-4w"], 2)

    def test_uart_config_matches_ti_driver_wire_layout(self):
        payload = helper.build_uart_config(
            "rs485-2w",
            baudrate=19200,
            bytesize=8,
            parity="N",
            stopbits=1,
        )

        self.assertEqual(len(payload), 10)
        divisor, flags, data_bits, parity, stop_bits, xon, xoff, uart_mode = struct.unpack(
            ">HHBBBBBB", payload
        )
        self.assertEqual(divisor, (923077 + 9600) // 19200)
        self.assertEqual(flags, 0x6000)
        self.assertEqual((data_bits, parity, stop_bits), (3, 0, 0))
        self.assertEqual((xon, xoff, uart_mode), (0x11, 0x13, 1))

    def test_mode_config_preserves_every_active_uart_setting(self):
        active = bytes.fromhex("00 60 60 28 03 00 00 11 13 00")

        updated = helper.build_mode_config(active, "rs485-2w")

        self.assertEqual(updated[:-1], active[:-1])
        self.assertEqual(updated[-1], 1)

    def test_mode_config_rejects_incomplete_device_response(self):
        with self.assertRaisesRegex(ValueError, "10-byte"):
            helper.build_mode_config(b"short", "rs485-2w")

    def test_apply_validates_device_then_sends_config(self):
        active = bytes.fromhex("00 60 60 28 03 00 00 11 13 00")
        device = helper.MoxaUsbDevice(
            tty_path="/dev/ttyUSB0",
            usb_path=helper.Path("/dev/bus/usb/001/002"),
            product_id=0x1151,
            product_name="UPort 1150I",
        )
        with (
            patch.object(helper, "find_moxa_uport", return_value=device),
            patch.object(
                helper,
                "_usb_control_transfer",
                side_effect=[active, b""],
            ) as transfer,
        ):
            result = helper.apply_moxa_uport_mode(
                "/dev/ttyUSB0",
                "rs485-2w",
            )

        self.assertIs(result, device)
        self.assertEqual(transfer.call_count, 2)
        read_call, write_call = transfer.call_args_list
        self.assertEqual(read_call.args[0], device.usb_path)
        self.assertEqual(read_call.kwargs["request_type"], helper.USB_VENDOR_DEVICE_IN)
        self.assertEqual(read_call.kwargs["request"], helper.TI_GET_CONFIG)
        self.assertEqual(read_call.kwargs["response_size"], helper.UART_CONFIG_SIZE)
        self.assertEqual(write_call.args[0], device.usb_path)
        self.assertEqual(write_call.kwargs["request_type"], helper.USB_VENDOR_DEVICE_OUT)
        self.assertEqual(write_call.kwargs["request"], helper.TI_SET_CONFIG)
        self.assertEqual(write_call.kwargs["payload"][:-1], active[:-1])
        self.assertEqual(write_call.kwargs["payload"][-1], 1)

    def test_invalid_serial_format_is_rejected_before_usb_access(self):
        with self.assertRaisesRegex(ValueError, "N, O, E, M, or S"):
            helper.build_uart_config(
                "rs485-2w",
                baudrate=19200,
                bytesize=8,
                parity="X",
                stopbits=1,
            )


if __name__ == "__main__":
    unittest.main()
