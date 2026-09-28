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

    def test_release_workflow_flattens_artifacts_before_upload(self):
        workflow = (REPO_ROOT / ".github" / "workflows" / "release.yml").read_text(encoding="utf-8")

        self.assertIn("path: downloaded-assets", workflow)
        self.assertIn("find downloaded-assets -type f -print0", workflow)
        self.assertIn('destination="release-assets/$filename"', workflow)
        self.assertIn('gh release upload "$RELEASE_TAG" ./*', workflow)

    def test_uport_installer_defaults_to_rs485_two_wire_and_verifies_archive(self):
        script = (REPO_ROOT / "drivers" / "moxa" / "install_uport_1150i.sh").read_text(encoding="utf-8")

        self.assertIn('mode="rs485-2w"', script)
        self.assertIn("DEFAULT_UART_MODE=${compile_mode}", script)
        self.assertIn("sha512sum", script)
        self.assertIn("moxa-uport-1100-series-linux-kernel-6.x-driver-v6.0.tgz", script)
        self.assertIn("--force-unsupported-kernel", script)
        self.assertIn("--user-agent", script)
        self.assertIn('log_dir="/var/log/serial-loopback-tester"', script)
        self.assertIn('install_log="${log_dir}/uport-install.log"', script)
        self.assertIn('build_log="${log_dir}/uport-build.log"', script)
        self.assertIn('cp -- "${source_dir}/build.log" "$build_log"', script)
        self.assertIn("Verified the kernel 6.x break-control callback patch.", script)
        self.assertIn("uport-modern-break-ctl.patch", script)

    def test_real_tty_installer_maps_four_nport_ports_and_verifies_archive(self):
        script = (REPO_ROOT / "drivers" / "moxa" / "install_nport_real_tty.sh").read_text(encoding="utf-8")

        self.assertIn("ports=4", script)
        self.assertIn("mxaddsvr", script)
        self.assertIn("sha512sum", script)
        self.assertIn("moxa-real-tty-linux-kernel-6.x-driver-v6.2.tar", script)
        self.assertIn("--force-unsupported-kernel", script)
        self.assertIn("--user-agent", script)
        self.assertIn("serial-loopback-tester-nport-${kernel_release}.log", script)

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
        self.assertIn("install_fedora_driver_dependencies.sh", script)
        self.assertIn("set_moxa_uport_mode.sh", script)
        self.assertIn("moxa_uport_mode.py", script)
        self.assertIn("patches", script)

    def test_fedora_dependency_installer_covers_driver_build_tools(self):
        script = (REPO_ROOT / "install_fedora_driver_dependencies.sh").read_text(encoding="utf-8")

        self.assertIn("kernel-devel-uname-r == ${kernel_release}", script)
        self.assertIn("gcc", script)
        self.assertIn("make", script)
        self.assertIn("patch", script)
        self.assertIn("setserial", script)
        self.assertIn("tar", script)
        self.assertIn("coreutils", script)
        self.assertIn("sha512sum", script)
        self.assertIn("kojipkgs.fedoraproject.org", script)
        self.assertIn("rpmkeys --checksig", script)
        self.assertNotIn('grep -q "signatures OK"', script)
        self.assertNotIn("sudo dnf upgrade --refresh", script)

    def test_uport_compatibility_patch_fixes_modern_break_callback(self):
        patch = (REPO_ROOT / "drivers" / "moxa" / "patches" / "uport-modern-break-ctl.patch").read_text(encoding="utf-8")

        self.assertIn("static int mxu1_break(struct tty_struct *tty, int break_state)", patch)
        self.assertIn("return -ENODEV", patch)
        self.assertIn("return 0", patch)

    def test_normal_and_root_start_scripts_are_present(self):
        normal_script = (REPO_ROOT / "start.sh").read_text(encoding="utf-8")
        root_script = (REPO_ROOT / "start_as_root.sh").read_text(encoding="utf-8")

        self.assertIn("SerialLoopbackTester-v*-linux-*", normal_script)
        self.assertIn('exec "${script_dir}/run_as_root.sh"', root_script)

    def test_repository_production_runner_uses_isolated_runtime_environment(self):
        script = (REPO_ROOT / "run_production.sh").read_text(encoding="utf-8")

        self.assertIn(".venv-production", script)
        self.assertIn("requirements.txt", script)
        self.assertIn("import tkinter", script)
        self.assertIn('if [[ "${1:-}" == "--root" ]]', script)
        self.assertIn('exec "$venv_python" "${repo_dir}/serial_tester_gui.py"', script)

    def test_root_launcher_preserves_graphical_session_environment(self):
        script = (REPO_ROOT / "run_as_root.sh").read_text(encoding="utf-8")

        self.assertIn("sudo --preserve-env=DISPLAY,XAUTHORITY,WAYLAND_DISPLAY,XDG_RUNTIME_DIR,DBUS_SESSION_BUS_ADDRESS", script)
        self.assertIn('SERIAL_LOOPBACK_TESTER_CONFIG_HOME="$config_home"', script)
        self.assertIn('SERIAL_LOOPBACK_TESTER_CONFIG_UID="$config_uid"', script)
        self.assertIn('SERIAL_LOOPBACK_TESTER_CONFIG_GID="$config_gid"', script)
        self.assertIn('${script_dir}/.venv-production', script)
        self.assertIn('${source_venv}/bin/python', script)
        self.assertIn("SerialLoopbackTester-v*-linux-*", script)

    def test_serial_access_installer_covers_usb_and_moxa_tty_devices(self):
        script = (REPO_ROOT / "install_serial_access.sh").read_text(encoding="utf-8")

        self.assertIn('KERNEL=="ttyUSB[0-9]*"', script)
        self.assertIn('KERNEL=="ttyACM[0-9]*"', script)
        self.assertIn('KERNEL=="ttyr[0-9a-fA-F]*"', script)
        self.assertIn('ATTR{idVendor}=="110a"', script)
        self.assertIn('ATTR{idProduct}=="1150"', script)
        self.assertIn('ATTR{idProduct}=="1151"', script)
        self.assertIn("usermod -aG dialout", script)

    def test_manual_moxa_mode_script_is_packaged_and_uses_userspace_helper(self):
        script = (REPO_ROOT / "set_moxa_uport_mode.sh").read_text(encoding="utf-8")

        self.assertIn("moxa_uport_mode.py", script)
        self.assertIn('exec "$python_bin"', script)


if __name__ == "__main__":
    unittest.main()
