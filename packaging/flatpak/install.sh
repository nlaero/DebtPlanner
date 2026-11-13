#!/bin/sh
# Build and install the Flatpak from a git clone. Needs a network connection.
set -eu
ROOT=$(CDPATH= cd -- "$(dirname "$0")/../.." && pwd)
cd "$ROOT"
if ! command -v flatpak >/dev/null 2>&1 || ! command -v flatpak-builder >/dev/null 2>&1; then
    echo "Install flatpak and flatpak-builder, then run this script again." >&2
    echo "Debian/Ubuntu: sudo apt install flatpak flatpak-builder" >&2
    exit 1
fi
flatpak remote-add --user --if-not-exists flathub https://dl.flathub.org/repo/flathub.flatpakrepo
flatpak install --user -y flathub org.freedesktop.Platform//24.08 org.freedesktop.Sdk//24.08
flatpak-builder --user --install --force-clean --disable-rofiles-fuse build-flatpak packaging/flatpak/io.github.nlaero.DebtPayoffPlanner.yml
echo "Installed. Start it with: flatpak run io.github.nlaero.DebtPayoffPlanner"
