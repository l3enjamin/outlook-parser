@echo off
REM Change to the script's directory (plugin root)
cd /d "%~dp0"
REM Run the MCP server
uv run --with pywin32 -m mailtool.mcp.server
