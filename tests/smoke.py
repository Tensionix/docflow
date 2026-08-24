from __future__ import annotations

import argparse
from pathlib import Path
import subprocess
import sys
import unittest


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def run_command(command: list[str]) -> None:
    completed = subprocess.run(command, cwd=ROOT, check=False)
    if completed.returncode:
        raise SystemExit(completed.returncode)


def run_unittest(pattern: str) -> None:
    suite = unittest.defaultTestLoader.discover(str(ROOT / "tests"), pattern=pattern)
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    if not result.wasSuccessful():
        raise SystemExit(1)


def main() -> int:
    parser = argparse.ArgumentParser(description="Audion DocFlow smoke suite")
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--quick", action="store_true")
    mode.add_argument("--full", action="store_true")
    parser.add_argument("--keep-artifacts", action="store_true", help="Accepted for compatibility; tests use temporary directories.")
    args = parser.parse_args()

    run_command([sys.executable, "-m", "compileall", "-q", "system_core"])
    run_command([sys.executable, "-m", "system_core.ui_nicegui.app", "--smoke"])
    run_unittest("test_workbench.py")
    run_unittest("test_gui_routes.py")
    run_unittest("test_manifest_coverage.py")
    if args.full:
        run_command([sys.executable, str(ROOT / "tests" / "text_hygiene_boundaries.py")])
    print(f"[OK] {'full' if args.full else 'quick'} smoke completed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
