"""Build a Debian package that installs the planner for the system Python.

Run this on a machine that already has customtkinter installed, from the
project folder: python packaging/debian/build_deb.py
"""

import gzip
import hashlib
import io
import os
import re
import tarfile
import time

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
PACKAGE = "debt-payoff-planner"
SHARE = f"usr/share/{PACKAGE}"


def read_version():
    text = open(os.path.join(ROOT, "debt_payoff_app.py"), encoding="utf-8").read()
    match = re.search(r'^APP_VERSION = "([^"]+)"', text, re.M)
    if not match:
        raise SystemExit("APP_VERSION was not found")
    return match.group(1)


def package_dir(module_name):
    module = __import__(module_name)
    return os.path.dirname(module.__file__)


def should_skip(path):
    name = os.path.basename(path)
    return name == "__pycache__" or name == ".DS_Store" or name.endswith(".pyc")


def add_tree(members, source_root, archive_root):
    for dirpath, dirnames, filenames in os.walk(source_root):
        dirnames[:] = [name for name in dirnames if not should_skip(name)]
        for filename in filenames:
            full = os.path.join(dirpath, filename)
            if should_skip(full):
                continue
            rel = os.path.relpath(full, source_root).replace("\\", "/")
            members.append((f"{archive_root}/{rel}", open(full, "rb").read(), 0o644))


def tar_gz(items):
    raw = io.BytesIO()
    with tarfile.open(fileobj=raw, mode="w", format=tarfile.GNU_FORMAT) as archive:
        for name, data, mode in items:
            info = tarfile.TarInfo(name)
            info.size = len(data)
            info.mode = mode
            info.mtime = 0
            info.uid = 0
            info.gid = 0
            info.uname = "root"
            info.gname = "root"
            if name.endswith("/") or data is None:
                info.type = tarfile.DIRTYPE
                info.size = 0
                archive.addfile(info)
            else:
                archive.addfile(info, io.BytesIO(data))
    compressed = io.BytesIO()
    with gzip.GzipFile(fileobj=compressed, mode="wb", mtime=0) as handle:
        handle.write(raw.getvalue())
    return compressed.getvalue()


def ar_member(name, data):
    header = (
        f"{(name + '/')[:16]:<16}"
        f"{'0':<12}{'0':<6}{'0':<6}{'100644':<8}{len(data):<10}`\n"
    ).encode("ascii")
    if len(header) != 60:
        raise SystemExit(f"bad ar header length {len(header)}")
    pad = b"\n" if len(data) % 2 else b""
    return header + data + pad


def main():
    version = read_version()
    debian_version = f"{version}-1"
    launcher = (
        "#!/bin/sh\n"
        f"export PYTHONPATH=\"/usr/share/{PACKAGE}/vendor${{PYTHONPATH:+:$PYTHONPATH}}\"\n"
        f"exec python3 /usr/share/{PACKAGE}/debt_payoff_app.py \"$@\"\n"
    ).encode()
    desktop = open(os.path.join(os.path.dirname(__file__), "debt-payoff-planner.desktop"), "rb").read()
    postinst = (
        "#!/bin/sh\nset -e\n"
        "if [ \"$1\" = \"configure\" ] && command -v gtk-update-icon-cache >/dev/null 2>&1; then\n"
        "    gtk-update-icon-cache -q /usr/share/icons/hicolor || true\n"
        "fi\n"
    ).encode()

    data = [
        ("usr/bin/debt-payoff-planner", launcher, 0o755),
        ("usr/share/applications/debt-payoff-planner.desktop", desktop, 0o644),
        (
            "usr/share/icons/hicolor/256x256/apps/debt-payoff-planner.png",
            open(os.path.join(ROOT, "debt_payoff.png"), "rb").read(),
            0o644,
        ),
    ]
    for filename in ("debt_payoff_app.py", "sample_plan.json", "debt_payoff.ico", "debt_payoff.png"):
        data.append((f"{SHARE}/{filename}", open(os.path.join(ROOT, filename), "rb").read(), 0o644))
    for module in ("customtkinter", "darkdetect", "packaging"):
        add_tree(data, package_dir(module), f"{SHARE}/vendor/{module}")

    md5_lines = []
    installed = 0
    for name, payload, _mode in data:
        md5_lines.append(f"{hashlib.md5(payload).hexdigest()}  {name}")
        installed += len(payload)
    md5sums = ("\n".join(md5_lines) + "\n").encode()

    control = (
        f"Package: {PACKAGE}\n"
        f"Version: {debian_version}\n"
        "Section: utils\n"
        "Priority: optional\n"
        "Architecture: all\n"
        "Depends: python3 (>= 3.9), python3-tk, python3-pil, python3-openpyxl, fonts-dejavu-core\n"
        "Maintainer: Nathan Laero <25691061+nlaero@users.noreply.github.com>\n"
        f"Installed-Size: {max(1, installed // 1024)}\n"
        "Description: Plan how to pay off debts\n"
        " Calculates payoff dates, extra payments, and interest\n"
        " for biweekly, semi-monthly, and monthly schedules.\n"
    ).encode()
    control_tar = tar_gz([
        ("control", control, 0o644),
        ("md5sums", md5sums, 0o644),
        ("postinst", postinst, 0o755),
    ])
    data_tar = tar_gz(data)
    deb = b"!<arch>\n"
    deb += ar_member("debian-binary", b"2.0\n")
    deb += ar_member("control.tar.gz", control_tar)
    deb += ar_member("data.tar.gz", data_tar)
    outfile = os.path.join(os.path.dirname(__file__), f"{PACKAGE}_{debian_version}_all.deb")
    with open(outfile, "wb") as handle:
        handle.write(deb)
    print(outfile)
    print(f"{len(deb)} bytes")


if __name__ == "__main__":
    main()
