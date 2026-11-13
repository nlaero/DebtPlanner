"""Refresh first_run_test_debt_payoff.py from debt_payoff_app.py.

The test file is a snapshot of the program that always starts like a new
installation. It is rewritten only when APP_VERSION in debt_payoff_app.py
differs from the snapshot, which is when a new version is compiled to the
installer. Pass --force to write the snapshot anyway.
"""

import os
import re
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
SOURCE_NAME = "debt_payoff_app.py"
TEST_NAME = "first_run_test_debt_payoff.py"
SOURCE_PATH = os.path.join(ROOT, SOURCE_NAME)
TEST_PATH = os.path.join(ROOT, TEST_NAME)
MAIN_GUARD = 'if __name__ == "__main__":\n    main()\n'


def read_text(path):
    with open(path, "r", encoding="utf-8", newline="\n") as handle:
        return handle.read()


def app_version(source):
    match = re.search(r'^APP_VERSION = "([^"]+)"', source, re.M)
    if not match:
        raise SystemExit("APP_VERSION was not found in debt_payoff_app.py")
    return match.group(1)


def stamped_version(text):
    match = re.search(r'^FIRST_RUN_TEST_VERSION = "([^"]+)"', text, re.M)
    return match.group(1) if match else None


def build_snapshot(source, version):
    if not source.endswith(MAIN_GUARD):
        raise SystemExit("debt_payoff_app.py no longer ends with the expected main guard")
    body = source[: -len(MAIN_GUARD)]
    header = (
        "#!/usr/bin/env python3\n"
        f'FIRST_RUN_TEST_VERSION = "{version}"\n'
        "# Snapshot of debt_payoff_app.py for first-install testing.\n"
        "# Every launch clears a private data folder and starts like a new installation.\n"
        "# Edit debt_payoff_app.py for program changes. This file is replaced when a new\n"
        "# version is compiled to the installer.\n"
    )
    if body.startswith("#!/usr/bin/env python3\n"):
        body = body[len("#!/usr/bin/env python3\n"):]
    trailer = '''
import shutil

def _first_run_data_dir():
    """Private folder for this test. It is never the real plan folder."""
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP") or os.path.expanduser("~")
    return os.path.join(base, "Debt Payoff Planner First Run Test")


def _reset_first_run_data():
    folder = os.path.abspath(_first_run_data_dir())
    project = os.path.dirname(os.path.abspath(__file__))
    folder_key = os.path.normcase(folder)
    project_key = os.path.normcase(project)
    if folder_key == project_key or folder_key.startswith(project_key + os.sep):
        raise RuntimeError("Refusing to clear the project folder for the first-run test")
    if os.path.isdir(folder):
        shutil.rmtree(folder)
    os.makedirs(folder, exist_ok=True)
    return folder


def _sandboxed_user_data_dir(self):
    folder = _first_run_data_dir()
    os.makedirs(folder, exist_ok=True)
    return folder


DebtPayoffApp._user_data_dir = _sandboxed_user_data_dir
_app_main = main


def _first_run_main():
    _reset_first_run_data()
    _app_main()


if __name__ == "__main__":
    _first_run_main()
'''
    return header + body + trailer


def main():
    force = "--force" in sys.argv
    source = read_text(SOURCE_PATH)
    version = app_version(source)
    current = None
    if os.path.isfile(TEST_PATH):
        current = stamped_version(read_text(TEST_PATH))
    if current == version and not force:
        print(f"First-run test file is already version {version}")
        return
    snapshot = build_snapshot(source, version)
    with open(TEST_PATH, "w", encoding="utf-8", newline="\n") as handle:
        handle.write(snapshot)
    action = "Rebuilt" if current else "Created"
    print(f"{action} {TEST_NAME} from {SOURCE_NAME} version {version}")


if __name__ == "__main__":
    main()
