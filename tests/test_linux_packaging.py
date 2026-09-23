import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]


class LinuxPackagingTests(unittest.TestCase):
    def test_release_workflow_runs_for_published_releases_and_version_tags(self):
        workflow = (REPO_ROOT / ".github" / "workflows" / "release.yml").read_text(encoding="utf-8")

        self.assertIn("release:\n    types:\n      - published", workflow)
        self.assertIn('- "v[0-9]*"', workflow)
        self.assertIn('- "[0-9]*"', workflow)
        self.assertIn("release_tag:", workflow)

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
        self.assertIn("--user-agent", script)
        self.assertIn("start.sh", script)
        self.assertIn("start_as_root.sh", script)
        self.assertIn("run_as_root.sh", script)
        self.assertIn("install_serial_access.sh", script)

    def test_normal_and_root_start_scripts_are_present(self):
        normal_script = (REPO_ROOT / "start.sh").read_text(encoding="utf-8")
        root_script = (REPO_ROOT / "start_as_root.sh").read_text(encoding="utf-8")

        self.assertIn("SerialLoopbackTester-v*-linux-*", normal_script)
        self.assertIn('exec "${script_dir}/run_as_root.sh"', root_script)

    def test_root_launcher_preserves_graphical_session_environment(self):
        script = (REPO_ROOT / "run_as_root.sh").read_text(encoding="utf-8")

        self.assertIn("sudo --preserve-env=DISPLAY,XAUTHORITY,WAYLAND_DISPLAY,XDG_RUNTIME_DIR,DBUS_SESSION_BUS_ADDRESS", script)
        self.assertIn("SerialLoopbackTester-v*-linux-*", script)

    def test_serial_access_installer_covers_usb_and_moxa_tty_devices(self):
        script = (REPO_ROOT / "install_serial_access.sh").read_text(encoding="utf-8")

        self.assertIn('KERNEL=="ttyUSB[0-9]*"', script)
        self.assertIn('KERNEL=="ttyACM[0-9]*"', script)
        self.assertIn('KERNEL=="ttyr[0-9a-fA-F]*"', script)
        self.assertIn("usermod -aG dialout", script)


if __name__ == "__main__":
    unittest.main()
