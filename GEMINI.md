# GEMINI.md - Mailtool: Outlook Automation Bridge

This project provides a robust bridge between Windows Outlook and various environments (Windows native, WSL2, and AI Agents via MCP). It uses COM automation to control Outlook directly.

## Project Overview

- **Purpose**: Access Outlook email, calendar, and tasks from Python/CLI/AI Agents.
- **Main Technologies**: 
  - **Python 3.13+** (Core logic)
  - **pywin32** (Windows COM automation)
  - **uv** (Dependency and project management)
  - **MCP Python SDK v2 + FastMCP** (AI Agent integration)
  - **Ruff** (Linting and formatting)
  - **Pytest** (Testing framework)

## Architecture

1.  **OutlookBridge (`src/mailtool/bridge.py`)**: The core class managing COM communication with Outlook. It uses O(1) lookups via `EntryID` for performance.
2.  **CLI (`src/mailtool/cli.py`)**: A command-line interface for manual interaction.
3.  **MCP Server (`src/mailtool/mcp/`)**: A Model Context Protocol server that exposes Outlook functionality to AI agents like Claude Code.
    - `server.py`: FastMCP server implementation with 23 tools.
    - `models.py`: Pydantic models for structured tool outputs.
    - `resources.py`: URI-based data access (e.g., `inbox://emails`).
    - `lifespan.py`: Manages the COM bridge lifecycle (startup/shutdown).
4.  **WSL2 Bridge**: `outlook.sh` and `outlook.bat` wrappers allow running commands from WSL2 that execute on the Windows host.

## Development Workflows

### Environment Setup
The project uses `uv`. To sync dependencies:
```powershell
uv sync --all-groups
```

### Building and Running
- **CLI**: `uv run mailtool <command>` (e.g., `uv run mailtool emails --limit 5`)
- **MCP Server**: `uv run --with pywin32 -m mailtool.mcp.server`
- **Legacy/Native**: `./outlook.sh` (WSL2) or `outlook.bat` (Windows)

### Testing
Tests are categorized into unit, bridge, and MCP tests. They require a running Outlook instance on Windows.
- **All tests**: `uv run pytest`
- **MCP tests**: `uv run --with pytest --with pywin32 python -m pytest tests/mcp/ -v`
- **Markers**: Use `-m email`, `-m calendar`, or `-m tasks` to filter tests.

### Linting and Formatting
Ruff is used for all code quality checks.
```powershell
uv run ruff check .      # Lint
uv run ruff format .     # Format
```

## Development Conventions

- **COM Safety**: Always use `_safe_get_attr()` when accessing COM object properties to handle potential `com_error` exceptions.
- **Async/Sync**: COM operations are inherently synchronous and apartment-threaded. The MCP server uses a thread pool executor in its lifespan to manage the bridge without blocking the async event loop.
- **Test Items**: All items created during tests should be prefixed with `[TEST]` for easy identification and automatic cleanup (see `conftest.py`).
- **Pydantic Models**: All new MCP tools must return a Pydantic model defined in `src/mailtool/mcp/models.py`.
- **Error Handling**: Use custom exceptions from `src/mailtool/mcp/exceptions.py` (`OutlookNotFoundError`, `OutlookComError`, `OutlookValidationError`) for MCP errors.

## Key Files
- `pyproject.toml`: Project metadata, dependencies, and Ruff configuration.
- `CLAUDE.md`: Detailed technical guide for AI assistants (high priority).
- `src/mailtool/bridge.py`: The "brain" of the project; handles all COM logic.
- `src/mailtool/mcp/server.py`: Defines the tools and resources available to AI agents.
