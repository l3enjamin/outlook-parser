#!/usr/bin/env bash
#
# Test Runner Script for WSL2
# Runs pytest via the Windows batch file
#

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Convert WSL path to Windows path for the batch file
WINDOWS_BATCH=$(wslpath -w "$SCRIPT_DIR/run_tests.bat")

# Execute the Windows batch file
# The batch file handles UNC path mapping internally
echo "Running tests via Windows batch..."
cmd.exe /c "$WINDOWS_BATCH" "$@"
