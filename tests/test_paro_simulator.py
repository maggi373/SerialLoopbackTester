import unittest

from paro_simulator import ParoSimulator


class ParoSimulatorTests(unittest.TestCase):
    def test_pressure_and_temperature_measurements_ramp_independently(self):
        simulator = ParoSimulator(device_id=1)

        self.assertEqual(simulator.feed(b"*0100P3\r\n", now=1.0), [b"*0001100.00000\r\n"])
        self.assertEqual(simulator.feed(b"*0100P3\r\n", now=1.1), [b"*0001200.00000\r\n"])
        self.assertEqual(simulator.feed(b"*0100Q3\r\n", now=1.2), [b"*000110.000\r\n"])

    def test_concatenated_write_enable_updates_one_parameter(self):
        simulator = ParoSimulator()

        self.assertEqual(
            simulator.feed(b"*0100EW*0100PR=18\r\n", now=2.0),
            [b"*0001PR=18\r\n"],
        )
        self.assertEqual(simulator.feed(b"*0100PR\r\n", now=2.1), [b"*0001PR=18\r\n"])
        self.assertEqual(simulator.feed(b"*0100TR\r\n", now=2.2), [b"*0001TR=72\r\n"])

    def test_protocol_errors_and_link_test_reply_match_arduino(self):
        simulator = ParoSimulator()

        self.assertEqual(simulator.feed(b"bad\r\n", now=3.0), [b"*000190\r\n"])
        self.assertEqual(simulator.feed(b"*0200P3\r\n", now=3.1), [b"*000194\r\n"])
        self.assertEqual(simulator.feed(b"*0100PR=20\r\n", now=3.2), [b"*000198\r\n"])

    def test_incomplete_frame_times_out_and_continuous_mode_polls(self):
        simulator = ParoSimulator()
        self.assertEqual(simulator.feed(b"*0100", now=4.0), [])
        self.assertEqual(simulator.poll(now=4.251), [b"*000197\r\n"])

        self.assertEqual(simulator.feed(b"*0100P4\r\n", now=5.0), [b"*00010.00000\r\n"])
        self.assertEqual(simulator.poll(now=6.4), [b"*0001100.00000\r\n"])

    def test_global_baud_command_is_echoed_and_exposes_requested_baud(self):
        simulator = ParoSimulator()

        self.assertEqual(simulator.feed(b"*9900BR=19200\r\n", now=6.0), [b"*9900BR=19200\r\n"])
        self.assertEqual(simulator.take_pending_baud(), 19200)
        self.assertIsNone(simulator.take_pending_baud())


if __name__ == "__main__":
    unittest.main()
