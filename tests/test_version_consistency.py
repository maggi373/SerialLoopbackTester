import re
import unittest
from pathlib import Path

import serial_tester_gui as app_module


REPO_ROOT = Path(__file__).resolve().parents[1]


class VersionConsistencyTests(unittest.TestCase):
    def test_package_versions_match_application_version(self):
        expected = app_module.APP_VERSION
        version_patterns = {
            "build_installer.ps1": r'^\$appVersion = "([^"]+)"$',
            "build_linux.sh": r'^app_version="([^"]+)"$',
            "installer/serial_loopback_tester.iss": r'^#define MyAppVersion "([^"]+)"$',
            "README.md": r'^Version: `([^`]+)`$',
        }

        for relative_path, pattern in version_patterns.items():
            with self.subTest(path=relative_path):
                text = (REPO_ROOT / relative_path).read_text(encoding="utf-8")
                match = re.search(pattern, text, flags=re.MULTILINE)
                self.assertIsNotNone(match, f"Version was not found in {relative_path}")
                self.assertEqual(match.group(1), expected)


if __name__ == "__main__":
    unittest.main()
