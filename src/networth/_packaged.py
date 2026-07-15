"""Entry point for packaged one-click builds (no-Wine bundle, PyInstaller, …).

Always ensures --pause so a double-clicked console shows the summary before it
closes. Usable as a console-script entry (`networth._packaged:run`) and as a
module (`python -m networth._packaged`).
"""

from __future__ import annotations

import sys

from .update import main


def run() -> int:
    argv = sys.argv[1:]
    if "--pause" not in argv:
        argv = [*argv, "--pause"]
    return main(argv)


if __name__ == "__main__":
    raise SystemExit(run())
