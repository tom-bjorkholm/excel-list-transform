#! /usr/bin/env python3
"""Run focus-sensitive GUI tests manually under controlled conditions.

These tests require an unlocked display and an uninterrupted focused
window. Do not use the computer while they run.
"""

# Copyright (c) 2026 Tom Björkholm
# MIT License

import subprocess
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent / 'common_build_tools' / 'src'))
from build_utils import venv_python    # noqa: E402


def main() -> int:
    """Run focus-sensitive tests and return the exit code."""
    root = Path(__file__).parent
    test_folder = root / 'test'
    cmd = [*venv_python(), '-m', 'pytest', '-m', 'focus_sensitive',
           str(test_folder), '-v']
    return subprocess.run(cmd, check=False, cwd=root).returncode


if __name__ == '__main__':
    sys.exit(main())
