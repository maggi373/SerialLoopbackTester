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
