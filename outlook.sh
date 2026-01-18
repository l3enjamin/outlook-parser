#!/usr/bin/env bash
#
# Outlook Bridge Wrapper for WSL2
# Calls the Windows batch script which uses uv for dependency management
#

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Convert WSL path to Windows path for the batch file
WINDOWS_BATCH=$(wslpath -w "$SCRIPT_DIR/outlook.bat")

# Process arguments to convert WSL paths to Windows paths for attachments
# This handles the case where users run: ./outlook.sh send --attach /home/user/file.pdf
args=()
while [[ $# -gt 0 ]]; do
    case "$1" in
        --attach)
            shift
            # Convert all following paths until we hit another flag
            while [[ $# -gt 0 && ! "$1" =~ ^- ]]; do
                # Convert WSL path to Windows path
                winpath=$(wslpath -w "$1")
                args+=("$winpath")
                shift
            done
            ;;
        *)
            # Pass through other arguments as-is
            args+=("$1")
            shift
            ;;
    esac
done

# Execute the Windows batch file with processed arguments
# The batch file handles UNC path mapping internally
cmd.exe /c "$WINDOWS_BATCH" "${args[@]}"
