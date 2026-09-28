# Moxa Linux drivers

These wrappers install Moxa's official GPL-licensed source drivers after
verifying the published SHA-512 checksum. The Linux portable build bundles the
complete source archives next to these scripts. If an archive is absent, the
wrapper downloads the same version from Moxa over HTTPS and verifies it before
running any vendor code.

## UPort 1150I in RS-485 mode

The in-kernel `ti_usb_3410_5052` driver recognizes the UPort, but it does not
provide the interface-mode control needed here. Install Moxa's Linux 6.x driver
with RS-485 two-wire as the compiled default:

```sh
sudo ./install_uport_1150i.sh --mode rs485-2w
```

The installer always writes its complete output to
`/var/log/serial-loopback-tester/uport-install.log`. If Moxa creates its own
compiler `build.log`, that file is preserved at
`/var/log/serial-loopback-tester/uport-build.log` before the temporary build
directory is removed. Both files are readable without `sudo` after the run.

To apply the setting immediately to a connected device as well, install your
distribution's `setserial` package and add, for example,
`--device /dev/ttyUSB0`. Other accepted modes are `rs485-4w`, `rs422`, and
`rs232`.

## NPort 5410 Real TTY

The NPort is a network device and does not require a local hardware driver when
the application uses a raw `socket://` endpoint. Real TTY is optional and
creates conventional `/dev/ttyrXX` device nodes. Configure all required NPort
channels as **Real COM Mode** first, then run:

```sh
sudo ./install_nport_real_tty.sh --nport-ip 192.168.1.50
```

The NPort 5410 has four RS-232 ports, so the wrapper defaults to four mappings.
It lets Moxa's `mxaddsvr` choose its standard Real COM ports unless both
`--data-port` and `--command-port` are supplied explicitly.

Both installers require root access, a Linux 6.x kernel, matching kernel
headers, GCC, and Make. Secure Boot systems may require signing/enrolling the
out-of-tree kernel module according to the distribution's procedure.

On Fedora, install all driver build dependencies from the application or
repository root before running either driver installer:

```sh
./install_fedora_driver_dependencies.sh
```

This installs the exact `kernel-devel` package for the running kernel, GCC,
Make, `setserial`, `tar`, and `sha512sum` (provided by `coreutils`). If the
matching 6.x development package has left Fedora's active repositories, the
script tries Fedora's signed Koji archive. Do not update to kernel 7 solely to
obtain build files because the bundled Moxa driver supports kernel 6.x only.

An unsupported kernel 7 build can be attempted explicitly by adding
`--force-unsupported-kernel` to either driver command. This bypasses only the
installer's kernel-version guard; checksum/signature verification remains
enabled, and the vendor source may still require code changes before it will
compile or load.

The UPort installer applies the included
`patches/uport-modern-break-ctl.patch` after verifying and extracting Moxa's
original archive. The patch corrects Moxa's outdated `void` break-control
callback to the `int` callback required by modern USB-serial kernels. The
original downloaded archive remains unchanged and SHA-512 verified.
