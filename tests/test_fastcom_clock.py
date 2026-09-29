import importlib.util
from pathlib import Path
import unittest
from unittest.mock import patch


REPO_ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location(
    "program_fastcom_clock", REPO_ROOT / "program_fastcom_clock.py"
)
assert SPEC is not None and SPEC.loader is not None
fastcom_clock = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(fastcom_clock)


class FastcomClockSequenceTests(unittest.TestCase):
    def test_clock_sequence_sends_official_word_msb_first(self):
        saved = 0xA8
        values = fastcom_clock.clock_programming_values(saved)

        self.assertEqual(len(values), (24 * 3) + 3)
        shifted_word = 0
        for bit_number in range(24):
            data, clock_high, clock_low = values[bit_number * 3 : (bit_number + 1) * 3]
            bit = data & fastcom_clock.MPIO_DATA
            shifted_word = (shifted_word << 1) | bit
            self.assertEqual(clock_high, data | fastcom_clock.MPIO_CLOCK)
            self.assertEqual(clock_low, bit)

        self.assertEqual(shifted_word, fastcom_clock.FASTCOM_29MHZ_WORD)
        self.assertEqual(values[-3:], [fastcom_clock.MPIO_STROBE, 0, saved])

    def test_all_mode_preflights_three_cards_before_programming_any(self):
        cards = [Path(f"/sys/bus/pci/devices/0000:0{number}:00.0") for number in range(3)]
        events = []

        def prepare(card, required_tty_name=None):
            events.append(("prepare", card, required_tty_name))
            return Path("/sys/bus/pci/drivers/exar_serial"), {f"ttyS{len(events)}"}

        def program(card, _driver, _ttys):
            self.assertEqual(sum(event[0] == "prepare" for event in events), 3)
            events.append(("program", card))
            return f"programmed {card.name}"

        with (
            patch.object(fastcom_clock.os, "geteuid", return_value=0, create=True),
            patch.object(fastcom_clock.sys, "argv", ["program_fastcom_clock.py", "--all"]),
            patch.object(fastcom_clock, "find_all_fastcom_pci_devices", return_value=cards),
            patch.object(fastcom_clock, "prepare_card", side_effect=prepare),
            patch.object(fastcom_clock, "program_card", side_effect=program),
            patch("builtins.print") as output,
        ):
            fastcom_clock.main()

        self.assertEqual([event[0] for event in events], ["prepare"] * 3 + ["program"] * 3)
        output.assert_any_call("Programmed all 3 detected Fastcom PCI-335 cards.")


if __name__ == "__main__":
    unittest.main()
