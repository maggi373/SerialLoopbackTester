import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]


class LinuxPackagingTests(unittest.TestCase):
    def test_uport_installer_defaults_to_rs485_two_wire_and_verifies_archive(self):
        script = (REPO_ROOT / "drivers" / "moxa" / "install_uport_1150i.sh").read_text(encoding="utf-8")

        self.assertIn('mode="rs485-2w"', script)
        self.assertIn("DEFAULT_UART_MODE=${compile_mode}", script)
        self.assertIn("sha512sum", script)
        self.assertIn("moxa-uport-1100-series-linux-kernel-6.x-driver-v6.0.tgz", script)

    def test_real_tty_installer_maps_four_nport_ports_and_verifies_archive(self):
        script = (REPO_ROOT / "drivers" / "moxa" / "install_nport_real_tty.sh").read_text(encoding="utf-8")

        self.assertIn("ports=4", script)
        self.assertIn("mxaddsvr", script)
        self.assertIn("sha512sum", script)
        self.assertIn("moxa-real-tty-linux-kernel-6.x-driver-v6.2.tar", script)

    def test_linux_build_bundles_moxa_installers_and_socket_handler(self):
        script = (REPO_ROOT / "build_linux.sh").read_text(encoding="utf-8")

        self.assertIn("serial.urlhandler.protocol_socket", script)
        self.assertIn("install_uport_1150i.sh", script)
        self.assertIn("install_nport_real_tty.sh", script)
        self.assertIn("download_and_verify", script)


if __name__ == "__main__":
    unittest.main()
