import importlib.util
from pathlib import Path
import unittest


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


if __name__ == "__main__":
    unittest.main()
