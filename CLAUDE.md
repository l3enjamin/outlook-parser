# Mailtool: Outlook COM Automation Bridge

A WSL2-to-Windows bridge for Outlook automation via COM, optimized for AI agent integration.

**Version**: 2.3.0 (fork of upstream v2.3.0) | **Status**: Active Development

---

## Fork Context & Goals

This fork extends the upstream `mailtool` to make it more useful for **agentic workflows** ΓÇö specifically allowing agents to quickly scan email threads and create or update feature/issue tickets (e.g. GitHub Issues, Jira, Linear).

The core problem upstream doesn't solve well: Outlook reply chains embed the entire quoted history in every email. When an agent fetches 10 emails from a thread, it's reading the same content 10 times. This wastes context tokens and makes it hard to extract the *delta* ΓÇö what's actually new in each reply.

### Approach

- **Deduplication tiers**: configurable stripping of forwarded/quoted content via `deduplication_tier` in `get_email_parsed()`
- **Structured conversation output**: planned `get_email_thread()` tool that returns a full chronological thread with per-message dedup applied, so agents get one clean artifact per conversation
- **Agent-first field set**: the `EmailParsed` model returns `latest_reply`, `parent_found`, `deduplication_tier`, and `fragments` so agents can reason about what was deduplicated

---

## Architecture

**Stack**: Python 3.13+ + pywin32 (COM) ΓåÆ Outlook (Windows)

**Entry Points**:
- `uv run --with pywin32 -m mailtool.cli` (CLI)
- `uv run pytest` (Tests)
- **MCP Server** ΓåÆ `src/mailtool/mcp/server.py` ΓåÆ Claude Code / Cursor / Gemini CLI integration (26 tools, 7 resources)

**Dependency Management**: Uses `uv run --with pywin32` for zero-install Windows execution

**MCP Integration**: Model Context Protocol server using official MCP Python SDK v2 with FastMCP framework

**Development Tools**:
- `ruff` for linting and formatting
- GitHub Actions CI/CD (windows-latest)
- Pre-commit hooks for code quality

---

## Key Design Decisions

### O(1) Access Pattern
All item lookups use `GetItemFromID(entry_id)` instead of iteration. Critical for large mailboxes.

### Recurrence Handling
Calendar events: enable `IncludeRecurrences = True` + `Sort("[Start]")`, then apply COM-level `Restrict` filter **before** Python iteration to avoid the "Calendar Bomb" (infinite recurring meetings).

### Deduplication Architecture
`get_email_parsed()` implements a tiered deduplication strategy:

| Tier | Behaviour | Use Case |
|------|-----------|----------|
| `none` | Return full body as-is (default) | Raw access, archiving |
| `low` | Strip if `In-Reply-To` header present (reply detected by header alone) | Fast dedup, no DB lookup |
| `medium` | Strip only if parent email found in Inbox OR Sent Items by message-id or subject | Conservative dedup |
| `high` | Same as medium for now (content-based search not yet implemented) | Future: fuzzy match |

**Important**: `tier="low"` uses the `In-Reply-To` header as the signal ΓÇö if the header exists, this is a reply and the quoted content is stripped. No local parent lookup required. `tier="medium"` adds local validation before stripping.

### Known Issues (to fix)
1. **`_check_parent_exists()` only searches Inbox** ΓÇö parent emails sent *by you* live in Sent Items. The medium/high tier parent check misses ~50% of real threads. Fix: also search `get_folder_by_name("Sent Items")`.
2. **`tier="low"` currently gates on parent lookup** ΓÇö the current code checks `_check_parent_exists()` for `low` tier too, which contradicts the design intent above. `low` should strip unconditionally when `In-Reply-To` header is present.
3. **No thread/conversation tool yet** ΓÇö `get_email_thread()` is planned (see below) but not yet implemented.

---

## File Structure

```
outlook-parser/
Γö£ΓöÇΓöÇ .github/
Γöé   ΓööΓöÇΓöÇ workflows/
Γöé       Γö£ΓöÇΓöÇ ci.yml              # CI (tests + lint) on windows-latest
Γöé       ΓööΓöÇΓöÇ publish.yml         # PyPI publishing
Γö£ΓöÇΓöÇ src/
Γöé   ΓööΓöÇΓöÇ mailtool/
Γöé       Γö£ΓöÇΓöÇ __init__.py
Γöé       Γö£ΓöÇΓöÇ bridge.py           # Core COM automation (~1400 lines)
Γöé       Γö£ΓöÇΓöÇ cli.py              # CLI interface
Γöé       ΓööΓöÇΓöÇ mcp/                # MCP SDK v2 package
Γöé           Γö£ΓöÇΓöÇ __init__.py
Γöé           Γö£ΓöÇΓöÇ server.py       # FastMCP server (26 tools, 7 resources)
Γöé           Γö£ΓöÇΓöÇ models.py       # Pydantic models (10 models + EmailParsed)
Γöé           Γö£ΓöÇΓöÇ resources.py    # MCP resources (7 resources)
Γöé           Γö£ΓöÇΓöÇ lifespan.py     # Outlook bridge lifecycle management
Γöé           Γö£ΓöÇΓöÇ com_state.py    # COM threading state helpers
Γöé           ΓööΓöÇΓöÇ exceptions.py  # Custom exception classes
Γö£ΓöÇΓöÇ tests/
Γöé   Γö£ΓöÇΓöÇ __init__.py
Γöé   Γö£ΓöÇΓöÇ conftest.py             # Session fixtures, warmup, cleanup
Γöé   Γö£ΓöÇΓöÇ test_bridge.py          # Core connectivity
Γöé   Γö£ΓöÇΓöÇ test_emails.py          # Email ops
Γöé   Γö£ΓöÇΓöÇ test_calendar.py        # Calendar ops
Γöé   Γö£ΓöÇΓöÇ test_tasks.py           # Task ops
Γöé   ΓööΓöÇΓöÇ mcp/                    # MCP SDK v2 tests
Γöé       Γö£ΓöÇΓöÇ test_models.py
Γöé       Γö£ΓöÇΓöÇ test_tools.py
Γöé       Γö£ΓöÇΓöÇ test_resources.py
Γöé       Γö£ΓöÇΓöÇ test_integration.py
Γöé       ΓööΓöÇΓöÇ test_exceptions.py
Γö£ΓöÇΓöÇ CLAUDE.md                   # This file
Γö£ΓöÇΓöÇ GEMINI.md                   # Gemini CLI agent guide
Γö£ΓöÇΓöÇ README.md
Γö£ΓöÇΓöÇ pyproject.toml              # Project config (uv, ruff, dependencies)
Γö£ΓöÇΓöÇ pytest.ini
ΓööΓöÇΓöÇ uv.lock
```

---

## Email Conversation Deduplication

### `get_email_parsed()` ΓÇö Bridge Method

```python
bridge.get_email_parsed(
    entry_id,
    remove_quoted=False,        # Deprecated: use deduplication_tier="low"
    deduplication_tier="none",  # "none" | "low" | "medium" | "high"
    strip_html=True,            # Convert HTML to plain text via BeautifulSoup
)
```

**Returns** (dict matching `EmailParsed` model):

| Field | Type | Description |
|-------|------|-------------|
| `entry_id` | str | Outlook EntryID |
| `subject` | str | Email subject |
| `from` | list[tuple] | Sender (name, email) |
| `to` / `cc` / `bcc` | list[tuple] | Recipients |
| `date` | str | ISO timestamp |
| `message_id` | str | SMTP Message-ID header |
| `headers` | dict | All SMTP headers |
| `body` | str | Final cleaned body (dedup applied) |
| `text_plain` | list[str] | Raw plain text parts |
| `text_html` | list[str] | Raw HTML parts (empty if `strip_html=True`) |
| `attachments` | list[dict] | Attachment metadata (no payload) |
| `latest_reply` | str \| None | Extracted latest reply fragment |
| `deduplication_tier` | str | Which tier was applied |
| `parent_found` | bool \| None | Whether parent email was found locally |

### `_extract_latest_reply()` ΓÇö Uses `mailparser_reply`

Calls `EmailReplyParser.read(text_body).latest_reply` to extract only the newest content from a reply chain. Falls back gracefully if `mailparser_reply` is not installed.

### Planned: `get_email_thread()` ΓÇö Full Conversation

Not yet implemented. Planned signature:

```python
bridge.get_email_thread(
    entry_id,
    deduplication_tier="low",
    strip_html=True,
)
# Returns: list[dict]  ΓÇö chronological list of EmailParsed dicts,
#          one per email in the conversation thread, dedup applied per message.
```

Implementation will use Outlook's `item.GetConversation().GetTable()` to walk the full thread efficiently without iterating every folder. This will be exposed as a new MCP tool `get_email_thread`.

---

## API Patterns

### Return Values
- **Draft emails**: Returns `EntryID` (string)
- **Sent emails**: Returns `True`
- **Failed ops**: Returns `False`
- **Get ops**: Returns `dict` or `None`

### Test Isolation
All test-created items use `[TEST]` prefix for identification and auto-cleanup. Tests run against real Outlook ΓÇö no mocking.

---

## MCP SDK v2 Architecture

### FastMCP Framework

**Key Components**:
- **FastMCP Server**: `src/mailtool/mcp/server.py` ΓÇö 26 tools, 7 resources
- **Pydantic Models**: `src/mailtool/mcp/models.py` ΓÇö 10+ models for structured output
- **MCP Resources**: `src/mailtool/mcp/resources.py` ΓÇö 7 resources
- **Lifespan Management**: `src/mailtool/mcp/lifespan.py` ΓÇö async context manager
- **Custom Exceptions**: `src/mailtool/mcp/exceptions.py` ΓÇö 3 exception types

### Available MCP Tools

**Email (11 tools)**: `list_emails`, `list_unread_emails`, `get_email`, `get_email_parsed`, `send_email`, `reply_email`, `forward_email`, `mark_email`, `move_email`, `delete_email`, `search_emails`, `search_emails_by_sender`

> `get_email_parsed` accepts `deduplication_tier` and `strip_html` params. Prefer this over `get_email` for agent workflows that need to create/update tickets.

**Calendar (7 tools)**: `list_calendar_events`, `create_appointment`, `get_appointment`, `edit_appointment`, `respond_to_meeting`, `delete_appointment`, `get_free_busy`

**Tasks (7 tools)**: `list_tasks`, `list_all_tasks`, `create_task`, `get_task`, `edit_task`, `complete_task`, `delete_task`

### Available MCP Resources

**Email (3)**: `inbox://emails`, `inbox://unread`, `email://{entry_id}`

**Calendar (2)**: `calendar://today`, `calendar://week`

**Tasks (2)**: `tasks://active`, `tasks://all`

### FastMCP Decorator Pattern

```python
@mcp.tool()
def get_email_parsed(
    entry_id: str,
    deduplication_tier: str = "low",
    strip_html: bool = True,
) -> EmailParsed:
    """Get structured email with quoted content removed for agent use.

    Args:
        entry_id: Outlook EntryID of the email
        deduplication_tier: "none" | "low" | "medium" | "high"
        strip_html: Convert HTML body to plain text (default: True)
    """
    bridge = _get_bridge()
    result = bridge.get_email_parsed(entry_id, deduplication_tier=deduplication_tier, strip_html=strip_html)
    if not result:
        raise OutlookNotFoundError(entry_id)
    return EmailParsed(**result)
```

### Pydantic Models

**Email Models**:
- `EmailSummary`: 7 fields (entry_id, subject, sender, sender_name, received_time, unread, has_attachments)
- `EmailDetails`: 8 fields (+ body, html_body)
- `EmailParsed`: Full structured parsed email with dedup metadata (latest_reply, deduplication_tier, parent_found)
- `SendEmailResult`: 3 fields (success, entry_id, message)

**Calendar Models**:
- `AppointmentSummary`: 12 fields
- `AppointmentDetails`: 13 fields (+ body)
- `CreateAppointmentResult`: 3 fields
- `FreeBusyInfo`: 6 fields

**Task Models**:
- `TaskSummary`: 8 fields
- `CreateTaskResult`: 3 fields

**Generic**:
- `OperationResult`: 2 fields (success, message)

### Lifespan Management

```python
@asynccontextmanager
async def outlook_lifespan(app):
    bridge = await loop.run_in_executor(None, _create_bridge)
    await _warmup_bridge(bridge)  # 5 retries, 0.5s delay
    app._bridge = bridge
    yield
    # cleanup: release COM objects + gc.collect()
```

### Custom Exceptions

| Exception | Code | When |
|-----------|------|------|
| `OutlookNotFoundError` | -32602 | Item not found by entry_id |
| `OutlookComError` | -32603 | COM/bridge operation failed |
| `OutlookValidationError` | -32604 | Input validation failed |

---

## MCP Usage

### Installation

```bash
# Install in editable mode (Windows)
cd C:\dev\outlook-parser
uv pip install -e .
```

### Claude Code / Claude Desktop Config

Add to your `claude.json` `mcpServers` section:

```json
{
  "mcpServers": {
    "mailtool": {
      "type": "stdio",
      "command": "uv",
      "args": [
        "run", "--with", "pywin32",
        "-m", "mailtool.mcp.server",
        "--account", "your@email.com"
      ],
      "env": { "PYTHONUNBUFFERED": "1" }
    }
  }
}
```

Restart the client. Ensure Outlook is running on Windows.

### Example Agent Interactions

```
# Ticket creation workflow
"Search for emails about the login bug from this week and create a GitHub issue"
ΓåÆ agent calls search_emails(subject="login bug")
ΓåÆ for each hit: get_email_parsed(entry_id, deduplication_tier="low")
ΓåÆ agent drafts issue from de-duplicated thread content

# Thread review
"Summarise the email conversation about the Q1 roadmap"
ΓåÆ agent calls get_email_thread(entry_id, deduplication_tier="low")  [planned]
ΓåÆ returns chronological thread with only delta content per reply
```

---

## Running Tests

```bash
# All tests
uv run pytest -v

# Specific markers
uv run pytest -m email -v
uv run pytest -m "not slow" -v

# MCP tests only (requires Outlook running)
uv run --with pytest --with pywin32 python -m pytest tests/mcp/ -v

# Test server manually
uv run --with mcp --with pywin32 python -m mailtool.mcp.server
```

---

## Development Workflow

```bash
uv sync --all-groups           # Install deps
uv run ruff check .            # Lint
uv run ruff check --fix .      # Auto-fix
uv run ruff format .           # Format
uv run pre-commit run --all-files  # Pre-commit hooks
uv add <package>               # Add dependency
```

---

## Known Limitations

- **COM Threading**: All COM calls must happen on the same thread (session-scoped bridge fixture)
- **Date Format**: Outlook COM filters use locale-specific formats (MM/DD/YYYY HH:MM)
- **Sent Item ID**: Sent emails move to Sent Items with new EntryID (can't return original ID)
- **Parallel Execution**: COM is apartment-threaded; parallel test execution not recommended
- **`_check_parent_exists` Sent Items gap**: See Known Issues above ΓÇö medium/high tier dedup misses parent emails in Sent Items

---

## Recent Changes

### Fork v2.3.0 (Current)

1. **Deduplication tiers**: `get_email_parsed()` extended with `deduplication_tier` param (`none`/`low`/`medium`/`high`)
2. **`_extract_latest_reply()`**: Uses `mailparser_reply.EmailReplyParser` to isolate latest reply content
3. **`_check_parent_exists()`**: Validates parent email exists before stripping (medium/high tiers)
4. **`EmailParsed` model fields**: Added `latest_reply`, `deduplication_tier`, `parent_found` to return dict
5. **HTML stripping**: `strip_html=True` default converts HTML body to plain text via BeautifulSoup
6. **GEMINI.md**: Added Gemini CLI agent guide alongside CLAUDE.md

### Upstream v2.3.0 (Base)

1. MCP SDK v2 migration (FastMCP framework)
2. 26 tools, 7 resources with Pydantic models
3. Async lifespan management
4. Custom exception classes
5. GitHub Actions CI/CD

### Upstream v2.2.0

1. Initial MCP server implementation
2. 24 tools via JSON-RPC
3. Pre-commit hooks + ruff

### Upstream v2.1.0

1. Calendar Bomb fix (`Restrict` before iteration)
2. WSL path translation
3. Free/Busy refactor
