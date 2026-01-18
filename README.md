# Mailtool - Outlook Automation

Access your Office 365 email and calendar from WSL2 via Windows Outlook COM automation.

**Uses [uv](https://github.com/astral-sh/uv) for dependency management - no global Python needed!**

## Prerequisites

- Windows with Outlook (classic) installed and running
- WSL2 with `uv` installed (`pip install uv` or `curl -LsSf https://astral.sh/uv/install.sh | sh`)
- `uv.exe` accessible from WSL2 (automatically available if installed on Windows)

## Setup

### 1. Start Outlook

Make sure Outlook is running and logged into your `s.mok@utwente.nl` account.

### 2. That's it!

Dependencies are managed automatically by `uv`. No manual pip installs needed.

## Usage

```bash
# List recent emails
./outlook.sh emails --limit 5

# List calendar events for next 7 days
./outlook.sh calendar --days 7

# Get specific email body (use entry_id from emails command)
./outlook.sh email --id <entry_id>
```

## How It Works

1. WSL2 calls wrapper script (`outlook.sh`)
2. Wrapper calls Windows batch file (`outlook.bat`)
3. Batch file uses `uv run --with pywin32` to execute the Python script
4. Python script uses COM to talk to running Outlook instance
5. Data returned as JSON

## Project Structure

```
mailtool/
├── pyproject.toml          # uv project config
├── outlook.bat             # Windows entry point (uses uv)
├── outlook.sh              # WSL2 wrapper
├── src/
│   └── mailtool_outlook_bridge.py  # COM automation logic
└── .venv/                  # Linux virtualenv (for tooling)
```

## Advantages

- ✅ **uv for dependencies** - No global Python pollution
- ✅ **No API registration** - Uses existing Outlook auth
- ✅ **Works with any Outlook account**
- ✅ **Full access** to email and calendar
- ✅ **Stable** - Doesn't break on UI changes
- ✅ **Cross-shell** - Works from WSL2, PowerShell, etc.

## Limitations

- ⚠️ Outlook must be running on Windows
- ⚠️ Windows-specific (COM automation)
- ⚠️ Read-only for emails (can be extended)

## Claude Code Integration (MCP)

**NEW: This now includes an MCP (Model Context Protocol) server for Claude Code integration!**

Install as a Claude Code plugin to give Claude access to your Outlook:

```bash
# Clone to your Claude Code plugins directory
git clone <repo> ~/.claude-code/plugins/mailtool
```

Then Claude Code can:
- 📧 Read, send, reply to, and manage emails
- 📅 View, create, and manage calendar appointments
- ✅ Create, complete, and manage tasks

See [MCP_INTEGRATION.md](MCP_INTEGRATION.md) for full documentation.

## Future Directions

This could become:
- **CLI Tool**: Full-featured email/calendar CLI
- **Web App**: Backend for a web interface
- **Library**: Importable Python module

## Troubleshooting

### "Could not connect to Outlook"
- Make sure Outlook is running
- Check that you're logged into your account

### "uv.exe not found"
- Install uv on Windows: `powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"`
- Make sure uv is in your Windows PATH

### UNC path warnings (harmless)
- These appear because of WSL2 → Windows path translation
- Safe to ignore, everything still works

## Development

```bash
# Add new dependencies
uv add <package>

# Run on Linux/WSL2 (for tooling)
uv run python <script>

# Run on Windows (for COM automation)
./outlook.sh <command>
```
