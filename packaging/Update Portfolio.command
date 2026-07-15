#!/bin/bash
# macOS double-click launcher — runs the updater next to the workbook.
# First run on a fresh Mac: right-click this file -> Open (Gatekeeper).
cd "$(dirname "$0")"
./networth-updater "$@"
