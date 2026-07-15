"""Entry point for the packaged one-click updater.

Double-click launch: always pause before the console closes so the user can
read the summary. CLI arguments still pass through (e.g. a workbook path).
"""

import sys

from networth.update import main

argv = sys.argv[1:]
if "--pause" not in argv:
    argv.append("--pause")
sys.exit(main(argv))
