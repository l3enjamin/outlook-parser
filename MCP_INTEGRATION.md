# Mailtool MCP Integration

**Connect Claude Code to Outlook for email, calendar, and task management**

**Version 2.3.0** - Now powered by the official MCP Python SDK v2 with FastMCP framework!

## What is MCP?

MCP (Model Context Protocol) is a standard for exposing tools and data to AI assistants. This plugin exposes your Outlook data to Claude Code through an MCP server using the official MCP Python SDK v2 and FastMCP framework.

## Features

- **23 Tools** for complete Outlook automation
- **7 Resources** for quick data access
- **Structured Output** - All tools return typed Pydantic models
- **Type Safety** - Full type annotations for better IDE support
- **Error Handling** - Custom exception classes with detailed error messages
- **Logging** - Comprehensive logging for debugging and monitoring
- **Email Management**: Read, send, reply, forward, search, move, and delete emails
- **Calendar Management**: List events, create appointments, respond to meetings, check free/busy
- **Task Management**: List, create, edit, complete, and delete Outlook tasks

## Installation

### 1. Install the Plugin

Add this to your Claude Code plugins directory:

```bash
# On Linux/WSL2
git clone <this-repo> ~/.claude-code/plugins/mailtool

# Or add to your project's .claude-code/plugins directory
```

### 2. Prerequisites

- **Windows** with Outlook (classic) installed
- **Outlook must be running** before using the tools
- **uv** installed on Windows (for dependency management)
- **pywin32** (automatically handled by uv)

### 3. Start Outlook

Make sure Outlook is running and logged into your account before using any MCP tools.

## MCP Configuration

The MCP server is configured in `.mcp.json` at the plugin root:

```json
{
  "default": {
    "command": "uv",
    "args": [
      "run",
      "--with",
      "pywin32",
      "-m",
      "mailtool.mcp.server"
    ],
    "env": {
      "PYTHONUNBUFFERED": "1"
    }
  }
}
```

> **Note:** We use a separate `.mcp.json` file rather than inline `mcpServers` in `plugin.json` due to [Claude Code Bug #16143](https://github.com/anthropics/claude-code/issues/16143) where inline `mcpServers` may be ignored during plugin manifest parsing. This is the recommended pattern for Claude Code plugins.

## Usage Examples

### Email Operations

```bash
# List recent emails
Claude: "Show me my last 5 emails"

# Get email details
Claude: "Get the full email with entry ID 00000000604FC3F48B..."

# Send an email
Claude: "Send an email to john@example.com with subject 'Meeting tomorrow' and body 'Let's meet at 2pm'"

# Reply to an email
Claude: "Reply to email 00000000604FC3F48B... with 'Thanks, I'll be there'"

# Search emails
Claude: "Search for emails with subject 'invoice'"

# Move email to folder
Claude: "Move email 00000000604FC3F48B... to Archive"

# Mark as read/unread
Claude: "Mark email 00000000604FC3F48B... as read"
```

### Calendar Operations

```bash
# List upcoming events
Claude: "Show me my calendar for the next 7 days"

# Create an appointment
Claude: "Create an appointment titled 'Team Meeting' from 2026-01-25 14:00:00 to 2026-01-25 15:00:00 in Room 101"

# Get appointment details
Claude: "Get details for appointment 00000000604FC3F48B..."

# Respond to meeting
Claude: "Accept the meeting invitation 00000000604FC3F48B..."

# Check free/busy
Claude: "Check when john@example.com is free tomorrow"
```

### Task Operations

```bash
# List all tasks
Claude: "Show me all my tasks"

# Create a task
Claude: "Create a task 'Review project proposal' due on 2026-01-30 with high priority"

# Get task details
Claude: "Get details for task 00000000604FC3F48B..."

# Complete a task
Claude: "Mark task 00000000604FC3F48B... as complete"

# Edit a task
Claude: "Update task 00000000604FC3F48B... to be 50% complete"
```

## Available MCP Tools

### Email Tools

- `list_emails(limit: int = 10, folder: str = "Inbox") -> list[EmailSummary]` - List recent emails from inbox or another folder
- `get_email(entry_id: str) -> EmailDetails` - Get full email body and details by entry ID
- `send_email(to: str, subject: str, body: str, ...) -> SendEmailResult` - Send a new email or save as draft
- `reply_email(entry_id: str, body: str, reply_all: bool = False) -> OperationResult` - Reply to an email
- `forward_email(entry_id: str, to: str, body: str = "") -> OperationResult` - Forward an email
- `mark_email(entry_id: str, unread: bool = False) -> OperationResult` - Mark an email as read or unread
- `move_email(entry_id: str, folder: str) -> OperationResult` - Move an email to a different folder
- `delete_email(entry_id: str) -> OperationResult` - Delete an email
- `search_emails(filter_query: str, limit: int = 100) -> list[EmailSummary]` - Search emails using Outlook filter query

### Calendar Tools

- `list_calendar_events(days: int = 7, all_events: bool = False) -> list[AppointmentSummary]` - List calendar events for the next N days or all events
- `create_appointment(subject: str, start: str, end: str, ...) -> CreateAppointmentResult` - Create a new calendar appointment
- `get_appointment(entry_id: str) -> AppointmentDetails` - Get full appointment details by entry ID
- `edit_appointment(entry_id: str, ...) -> OperationResult` - Edit an existing appointment
- `respond_to_meeting(entry_id: str, response: str) -> OperationResult` - Respond to a meeting invitation
- `delete_appointment(entry_id: str) -> OperationResult` - Delete an appointment
- `get_free_busy(email_address: str | None = None, ...) -> FreeBusyInfo` - Get free/busy status for an email address

### Task Tools

- `list_tasks(include_completed: bool = False) -> list[TaskSummary]` - List active tasks
- `list_all_tasks() -> list[TaskSummary]` - List all tasks (including completed)
- `create_task(subject: str, ...) -> CreateTaskResult` - Create a new task
- `get_task(entry_id: str) -> TaskSummary` - Get full task details by entry ID
- `edit_task(entry_id: str, ...) -> OperationResult` - Edit an existing task
- `complete_task(entry_id: str) -> OperationResult` - Mark a task as complete
- `delete_task(entry_id: str) -> OperationResult` - Delete a task

## Available MCP Resources

Resources provide quick read-only access to Outlook data without tool calls:

### Email Resources

- `inbox://emails` - List recent emails (max 50)
- `inbox://unread` - List unread emails (max 50)
- `email://{entry_id}` - Get full email details (template resource)

### Calendar Resources

- `calendar://today` - List today's calendar events
- `calendar://week` - List calendar events for the next 7 days

### Task Resources

- `tasks://active` - List active (incomplete) tasks
- `tasks://all` - List all tasks (including completed)

**Usage:**
```
Claude: "Read the resource inbox://emails"
Claude: "Show me calendar://today"
Claude: "Get tasks://active"
```

## Architecture

**Version 2.3.0** uses the official MCP Python SDK v2 with FastMCP framework:

```
Claude Code (WSL2/Linux)
    ↓ (JSON-RPC via stdio)
FastMCP Server (mailtool.mcp.server)
    ↓ (async context manager)
Outlook COM Bridge (thread pool executor)
    ↓ (COM)
Outlook Application
```

**How it works:**
1. Claude Code calls MCP tools via JSON-RPC over stdio
2. FastMCP server runs on Windows using `uv run --with pywin32`
3. Async lifespan manager creates Outlook COM bridge via thread pool executor
4. Bridge uses COM to communicate with running Outlook instance
5. Results returned as structured Pydantic models, serialized to JSON

**Key SDK v2 Features:**
- **FastMCP Framework**: Declarative tool registration with `@mcp.tool()` decorator
- **Pydantic Models**: All tool outputs are typed models (EmailDetails, AppointmentDetails, TaskSummary, etc.)
- **Structured Output**: Automatic schema generation from Pydantic models
- **Async Lifespan**: COM bridge lifecycle managed by async context manager
- **Custom Exceptions**: OutlookNotFoundError, OutlookComError, OutlookValidationError
- **Resources**: 7 resources for quick data access (inbox://emails, calendar://today, tasks://active, etc.)

## Troubleshooting

### "Could not connect to Outlook"
- Make sure Outlook is running on Windows
- Check that you're logged into your account

### "uv.exe not found"
- Install uv on Windows: `powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"`
- Make sure uv is in your Windows PATH

### MCP server not starting
- Check Claude Code logs for errors
- Ensure `.claude-plugin/plugin.json` is valid JSON
- Verify Python dependencies are installed: `uv run --with pywin32 python -c "import win32com.client"`

### Tools return errors
- Verify Outlook is running
- Check that entry IDs are correct (get them from list commands first)
- Test with the CLI first: `./outlook.sh emails --limit 1`

## Development

### Project Structure

The MCP server is implemented in `src/mailtool/mcp/` using the official MCP Python SDK v2:

- **server.py** - FastMCP server with 23 tools
- **models.py** - Pydantic models for structured output (EmailDetails, AppointmentDetails, TaskSummary, etc.)
- **lifespan.py** - Async context manager for COM bridge lifecycle
- **resources.py** - 7 resources for quick data access
- **exceptions.py** - Custom exception classes (OutlookNotFoundError, OutlookComError, OutlookValidationError)

### Testing the MCP Server Directly

```bash
# Run the server directly (requires Outlook running)
uv run --with pywin32 python -m mailtool.mcp.server

# Run tests
./run_tests.sh -m mcp  # All MCP tests
uv run pytest tests/mcp/ -v  # Specific test file
```

### Adding New Tools

With FastMCP, adding new tools is straightforward:

1. **Add bridge method** in `src/mailtool/bridge.py` if needed
2. **Define Pydantic model** in `src/mailtool/mcp/models.py` for return type
3. **Add tool decorator** in `src/mailtool/mcp/server.py`:

```python
@mcp.tool()
def my_new_tool(param: str) -> MyResultModel:
    """
    Tool description for LLM understanding.

    Args:
        param: Parameter description

    Returns:
        MyResultModel: Result description

    Raises:
        OutlookComError: If bridge is not initialized
    """
    bridge = _get_bridge()
    result = bridge.my_bridge_method(param)
    return MyResultModel(...)
```

4. **Add tests** in `tests/mcp/test_tools.py`
5. **Update documentation** in MCP_INTEGRATION.md and README.md

### Code Patterns

**Tool Definition Pattern:**
```python
@mcp.tool()
def tool_name(param: type, ...) -> ResultModel:
    """Docstring with Args/Returns/Raises sections."""
    bridge = _get_bridge()  # Get bridge from module-level state
    result = bridge.bridge_method(...)
    return ResultModel(...)  # Return Pydantic model
```

**Resource Definition Pattern:**
```python
@mcp.resource(uri="scheme://path")
def resource_name() -> str:
    """Docstring describing the resource."""
    bridge = resources._get_bridge()
    items = bridge.list_items(...)
    return _format_items(items)  # Return formatted text
```

**Exception Pattern:**
```python
from mailtool.mcp.exceptions import OutlookNotFoundError, OutlookComError

if not item:
    raise OutlookNotFoundError(entry_id, "Email not found")

if not bridge:
    raise OutlookComError("Outlook bridge not initialized")
```

## Security Considerations

- This MCP server has full access to your Outlook data
- Only install plugins from trusted sources
- The server runs with your Windows user permissions
- Email addresses and calendar data are transmitted to Claude

## License

MIT License - See LICENSE file for details

## Contributing

Contributions welcome! Please read CONTRIBUTING.md for guidelines.
