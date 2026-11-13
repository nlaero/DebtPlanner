#!/bin/sh
# Install the Debian package from a git clone.
set -eu
DIR=$(CDPATH= cd -- "$(dirname "$0")" && pwd)
sudo apt install "$DIR"/debt-payoff-planner_*_all.deb
