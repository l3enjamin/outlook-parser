# Mailtool - Outlook Automation Bridge

A Python library and CLI tool for accessing Outlook email, calendar, and tasks via Windows COM automation. Optimized for use with AI agents via the Model Context Protocol (MCP).

**Uses [uv](https://github.com/astral-sh/uv) for dependency management - no global Python needed!**

## 🚀 Getting Started

### 1. Prerequisites

- Windows with Outlook (classic) installed and running.
- `uv` installed (`pip install uv` or `winget install astral-sh.uv` or `powershell -c "irm https://astral.sh/uv/install.ps1 | iex"`).

### 2. Installation

Clone the repository and sync the dependencies:

```bash
git clone https://github.com/l3enjamin/outlook-parser.git
cd outlook-parser
uv sync
```

### 3. Start Outlook

Ensure Outlook is running and logged into your account. The bridge communicates directly with the active Outlook process.

## 🤖 MCP Server Setup

This project includes a Model Context Protocol (MCP) server powered by the official Python SDK and FastMCP.

### Workspace Configuration (Gemini CLI)

To use this bridge in your AI workspace, create a `.gemini/mcp.json` file in your project root. 

**Requirements:**
- `--account`: You MUST provide your Outlook email address or account name.
- Feature Flags: You MUST explicitly enable modules using `--mail`, `--calendar`, or `--tasks`.
- Permissions: By default, the server is **read-only**. Use `--rw` to enable write operations (sending, deleting, creating).

```json
{
  "mcpServers": {
    "mailtool": {
      "command": "uv",
      "args": [
        "run",
        "mailtool",
        "mcp",
        "--account", "your-email@example.com",
        "--mail",
        "--calendar",
        "--tasks",
        "--rw"
      ]
    }
  }
}
```

### Manual Execution

Run the MCP server directly from the terminal (replace with your email):

```bash
# Enable everything with write access
uv run mailtool mcp --account your-email@example.com --mail --calendar --tasks --rw

# Enable only mail in read-only mode
uv run mailtool mcp --account your-email@example.com --mail
```

## 🛠️ Usage

### As a CLI Tool

```bash
# List recent emails
uv run mailtool emails --limit 5

# Search emails from a specific sender
uv run mailtool search --sender "John Doe"

# List calendar events for next 7 days
uv run mailtool calendar --days 7

# Get specific email body
uv run mailtool email --id <entry_id>

# List active tasks
uv run mailtool tasks
```

### As a Python Library

```python
from mailtool.bridge import OutlookBridge

# Create bridge instance
bridge = OutlookBridge()

# List emails
emails = bridge.list_emails(limit=5)
for email in emails:
    print(f"{email['subject']}: {email['sender']}")
```

## How It Works

The library uses Windows COM automation to communicate with Outlook:

1. Python creates a COM object to access the running Outlook instance.
2. Uses O(1) direct lookups via `GetItemFromID()` for high performance even with large mailboxes.
3. Returns structured data (emails, calendar events, tasks) as Python dictionaries or Pydantic models.
4. MCP server mode exposes this functionality via JSON-RPC for AI agents.

## Project Structure

```
mailtool/
├── pyproject.toml          # uv project config
├── src/
│   └── mailtool/
│       ├── __init__.py
│       ├── bridge.py       # Core COM automation (~2100 lines)
│       ├── cli.py          # CLI interface
│       └── mcp/            # MCP Server (SDK v2 + FastMCP)
│           ├── __init__.py
│           ├── server.py   # FastMCP server with 23 tools
│           ├── models.py   # Pydantic models
│           ├── lifespan.py # Async COM bridge lifecycle
│           ├── resources.py # 5 resources
│           ├── com_state.py # Thread-safe COM state management
│           └── exceptions.py # Custom exceptions
└── tests/
    ├── conftest.py         # Test fixtures
    ├── test_bridge.py      # Core connectivity tests
    ├── test_emails.py      # Email operation tests
    └── mcp/                # MCP server tests
```

## Advantages

- ✅ **uv for dependencies** - No global Python pollution.
- ✅ **Official MCP SDK v2** - Type-safe, declarative, and maintainable.
- ✅ **Structured output** - Pydantic models for all tool results.
- ✅ **Secure by Default** - Defaults to read-only; requires explicit opt-in for modules.
- ✅ **No API registration** - Uses your local Outlook authentication.
- ✅ **O(1) Access** - Fast performance via EntryID lookups.

## 📖 Command Reference

This tool provides access to Outlook **Email**, **Calendar**, and **Tasks**. It is available both as an **MCP Server** (for AI agents) and a **CLI** (for manual use).

### Email Management

| MCP Tool | CLI Command | Description | Key Parameters |
| :--- | :--- | :--- | :--- |
| `list_emails` | `emails` | List emails in a folder. | `limit`, `folder`, `unread_only` |
| `get_email` | `parsed-email`| Get full parsed details of an email. | `entry_id`, `deduplication_tier` |
| `search_emails` | `search` | Search by subject, sender name, or body. | `subject`, `sender`, `body`, `unread` |
| `search_emails_by_sender`| — | Recommended for specific email addresses. | `sender_email` (Exchange aware) |
| `send_email` | `send` | Send a new email or save a draft. | `to`, `subject`, `body`, `html_body` |
| `reply_email` | `reply` | Reply to an existing email. | `entry_id`, `body`, `reply_all` |
| `forward_email`| `forward` | Forward an email. | `entry_id`, `to`, `body` |

### Calendar Management

| MCP Tool | CLI Command | Description | Key Parameters |
| :--- | :--- | :--- | :--- |
| `list_calendar_events` | `calendar` | List events for the next N days. | `days`, `all_events` |
| `get_appointment` | `appointment` | Get full details/body of an event. | `entry_id` |
| `create_appointment` | `create-appt` | Create a new calendar event. | `subject`, `start`, `end`, `location` |
| `edit_appointment` | `edit-appt` | Modify an existing event. | `entry_id`, plus fields to change |
| `respond_to_meeting` | `respond` | Accept/Decline/Tentative. | `entry_id`, `response` |

### Task Management

| MCP Tool | CLI Command | Description | Key Parameters |
| :--- | :--- | :--- | :--- |
| `list_tasks` | `tasks` | List incomplete tasks. | `include_completed` |
| `get_task` | `task` | Get full task details and body. | `entry_id` |
| `create_task` | `create-task` | Create a new task. | `subject`, `body`, `due_date`, `priority`|
| `edit_task` | `edit-task` | Modify a task. | `entry_id`, `subject`, `body` |

---

## Development

```bash
# Run tests
uv run pytest

# Run linter and formatter
uv run ruff check .
uv run ruff format .
```

### Performance Benchmarks

Performance benchmarks are available in `scripts/benchmarks/` (requires Windows with Outlook running):

```bash
uv run --with pytest --with pywin32 python -m scripts.benchmarks.performance_benchmark
```
