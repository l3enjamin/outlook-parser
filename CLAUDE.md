# Mailtool: Outlook COM Automation Bridge

A WSL2-to-Windows bridge for Outlook automation via COM, optimized for AI agent integration.

**Version**: 2.2.0 | **Status**: Production/Stable

## Architecture

**Stack**: Python 3.13+ + pywin32 (COM) → Outlook (Windows)

**Entry Points**:
- `outlook.sh` (WSL2) → `outlook.bat` (Windows) → `src/mailtool/bridge.py`
- `run_tests.sh` (WSL2) → `run_tests.bat` (Windows) → `pytest`
- **MCP Server** → `mcp_server.py` → Claude Code integration (23 tools)

**Dependency Management**: Uses `uv run --with pywin32` for zero-install Windows execution

**MCP Integration**: Model Context Protocol server for Claude Code, Claude Desktop, and other MCP clients

**Development Tools**:
- `ruff` for linting and formatting (replaces Black, isort, Flake8, etc.)
- GitHub Actions CI/CD (windows-latest)
- Pre-commit hooks for code quality

## Key Design Decisions

### O(1) Access Pattern
All item lookups use `GetItemFromID(entry_id)` instead of iteration. This is critical for production use with large mailboxes.

### Recurrence Handling
Calendar events enable `IncludeRecurrences = True` + `Sort("[Start]")`, then apply COM-level `Restrict` filter **before** Python iteration to avoid the "Calendar Bomb" (infinite recurring meetings).

### Path Translation
WSL paths for attachments are auto-converted via `wslpath -w` in `outlook.sh` wrapper before being passed to Windows Python.

### Free/Busy API
Refactored to accept `email_address` directly, defaulting to current user. Legacy `entry_id` parameter supported but deprecated.

## File Structure

```
mailtool/
├── outlook.sh              # WSL2 entry point (translates paths)
├── outlook.bat             # Windows entry point (uv + pywin32)
├── run_tests.sh            # Test runner (WSL2)
├── run_tests.bat           # Test runner (Windows)
├── pytest.ini              # Pytest configuration
├── pyproject.toml          # Project config (uv, ruff, dependencies)
├── mcp_server.py           # MCP server for Claude Code (23 tools)
├── test_mcp_server.py      # MCP server validation script
├── .claude-plugin/
│   └── plugin.json         # Claude Code plugin manifest
├── .github/
│   └── workflows/
│       ├── ci.yml          # Continuous Integration (tests + lint)
│       └── publish.yml     # PyPI publishing
├── src/
│   └── mailtool/
│       ├── __init__.py
│       ├── bridge.py       # Core COM automation (~1100 lines)
│       └── cli.py          # CLI interface
├── tests/
│   ├── __init__.py
│   ├── conftest.py         # Session fixtures, warmup, cleanup
│   ├── test_bridge.py      # Core connectivity (6 tests)
│   ├── test_emails.py      # Email ops (12 tests)
│   ├── test_calendar.py    # Calendar ops (13 tests)
│   └── test_tasks.py       # Task ops (13 tests)
└── docs/
    ├── README.md           # Main project README
    ├── CLAUDE.md           # This file - AI assistant guide
    ├── MCP_INTEGRATION.md  # MCP server documentation
    ├── MCP_SUMMARY.md      # MCP implementation summary
    ├── SUMMARY.md          # Proof of concept summary
    ├── FEATURES.md         # Feature list
    ├── COMMANDS.md         # CLI command reference
    ├── QUICKSTART.md       # Quick start guide
    └── PRODUCTION_UPGRADE.md  # v2.0 upgrade notes
```

## API Patterns

### Return Values
- **Draft emails**: Returns `EntryID` (string) for reference
- **Sent emails**: Returns `True` (boolean)
- **Failed ops**: Returns `False` (boolean)
- **Get ops**: Returns `dict` with full item data or `None`

### Test Isolation
All test-created items use `[TEST]` prefix for identification and auto-cleanup. Tests use real Outlook data - no mocking.

## Recent Changes

### v2.2.0 (Current - MCP Integration)

1. **MCP Server**: Added Model Context Protocol server for Claude Code integration
2. **23 MCP Tools**: Email (9), Calendar (7), Tasks (7) operations exposed via JSON-RPC
3. **Plugin Manifest**: `.claude-plugin/plugin.json` for auto-loading in Claude Code
4. **Zero-Config MCP**: Uses `uv run --with pywin32` for dependency-free execution
5. **Claude Code Skills**: Plugin/skills integration for enhanced AI workflow
6. **Task Analysis**: List tasks, analyze by subject/deadline, recommend cleanup actions
7. **Pre-commit Hooks**: Automated code quality checks via pre-commit
8. **GitHub CI**: Automated testing and linting on Windows runners
9. **Ruff Integration**: Replaced multiple linters with unified ruff configuration

### v2.1.0 (Production Release)

1. **Calendar Bomb Fix**: Added `items.Restrict()` before iterating in `list_calendar_events()`
2. **WSL Path Translation**: Auto-convert attachment paths in `outlook.sh`
3. **Free/Busy Refactor**: Accepts `email_address` directly, defaults to current user
4. **Return Value Docs**: Clarified different return types in `send_email()` docstring
5. **Package Restructure**: Migrated from single-file to proper Python package structure

## Running Tests

```bash
# From WSL2
./run_tests.sh                 # All tests
./run_tests.sh -m email        # Email tests only
./run_tests.sh -m "not slow"   # Skip slow tests

# From Windows
run_tests.bat

# Test MCP server (requires Outlook running)
python test_mcp_server.py
```

## MCP Usage

### Installation

```bash
# Add to Claude Code plugins
cd ~/.claude-code/plugins
git clone <repo> mailtool

# Restart Claude Code - plugin auto-loads
# Start Outlook on Windows
```

### Available MCP Tools

**Email (9 tools)**: `list_emails`, `get_email`, `send_email`, `reply_email`, `forward_email`, `mark_email`, `move_email`, `delete_email`, `search_emails`

**Calendar (7 tools)**: `list_calendar_events`, `create_appointment`, `get_appointment`, `edit_appointment`, `respond_to_meeting`, `delete_appointment`, `get_free_busy`

**Tasks (7 tools)**: `list_tasks`, `list_all_tasks`, `create_task`, `get_task`, `edit_task`, `complete_task`, `delete_task`

### Example Claude Code Interactions

```
You: Show me my last 5 unread emails

You: Create a task "Review Q1 report" due Friday with high priority

You: Schedule a team meeting for tomorrow at 2pm in Room 101

You: Accept the meeting invitation from John

You: What's on my calendar this week?
```

See [MCP_INTEGRATION.md](MCP_INTEGRATION.md) for complete documentation.

## Known Limitations

- **Date Format**: Outlook COM filters use locale-specific formats (currently MM/DD/YYYY HH:MM)
- **Parallel Execution**: COM is apartment-threaded; true parallel test execution not recommended
- **Sent Item ID**: Sent emails move to Sent Items with new EntryID (can't return original ID)

## Development Notes

- **COM Threading**: All COM calls must happen on same thread (session-scoped bridge fixture)
- **Warmup**: Tests include 2-5s warmup to ensure Outlook is responsive
- **Cleanup**: Test artifacts auto-cleaned via prefix-based deletion helpers

## Development Workflow

```bash
# Install dependencies (managed by uv)
uv sync --all-groups

# Run linter and formatter
uv run ruff check .           # Check code
uv run ruff check --fix .     # Auto-fix issues
uv run ruff format .          # Format code

# Run tests
./run_tests.sh                # All tests (WSL2)
run_tests.bat                 # All tests (Windows)
uv run pytest -v              # Direct pytest
uv run pytest -m email        # Run specific marker

# Test MCP server
uv run --with pywin32 python test_mcp_server.py

# Add new dependency
uv add <package>

# Run pre-commit hooks manually
uv run pre-commit run --all-files
```

## Code Quality

- **Python Version**: Requires Python 3.13+
- **Linter/Formatter**: Ruff (unified tool replacing Black, isort, Flake8)
- **Line Length**: 88 characters (Black default)
- **CI/CD**: GitHub Actions runs tests and linting on Windows runners
- **Pre-commit Hooks**: Ensures code quality before commits

## Architecture Patterns

### Bridge Class (`src/mailtool/bridge.py`)
- **O(1) Lookups**: Uses `GetItemFromID(entry_id)` for all item access
- **Safe Attribute Access**: `_safe_get_attr()` wrapper for COM objects
- **Folder Access**: `get_inbox()`, `get_calendar()`, `get_tasks()`, `get_folder_by_name()`
- **Error Handling**: Returns `False` for failures, `None` for not found, EntryID string for drafts

### CLI Interface (`src/mailtool/cli.py`)
- **Entry Point**: `mailtool` command (installed via `uv add`)
- **Subcommands**: `emails`, `email`, `calendar`, `tasks`
- **Output**: JSON-formatted responses for easy parsing

### MCP Server (`mcp_server.py`)
- **Protocol**: JSON-RPC via stdio transport
- **Tool Registration**: All tools defined with JSON Schema validation
- **Initialization**: Single Outlook bridge instance per session
- **Error Responses**: Structured JSON-RPC error objects
