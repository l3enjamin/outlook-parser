# Mailtool MCP Integration

**Connect Claude Code to Outlook for email, calendar, and task management**

## What is MCP?

MCP (Model Context Protocol) is a standard for exposing tools and data to AI assistants. This plugin exposes your Outlook data to Claude Code through an MCP server.

## Features

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

- `list_emails` - List recent emails from inbox or another folder
- `get_email` - Get full email body and details by entry ID
- `send_email` - Send a new email or save as draft
- `reply_email` - Reply to an email
- `forward_email` - Forward an email
- `mark_email` - Mark an email as read or unread
- `move_email` - Move an email to a different folder
- `delete_email` - Delete an email
- `search_emails` - Search emails using Outlook filter query

### Calendar Tools

- `list_calendar_events` - List calendar events for the next N days or all events
- `create_appointment` - Create a new calendar appointment
- `get_appointment` - Get full appointment details by entry ID
- `edit_appointment` - Edit an existing appointment
- `respond_to_meeting` - Respond to a meeting invitation
- `delete_appointment` - Delete an appointment
- `get_free_busy` - Get free/busy status for an email address

### Task Tools

- `list_tasks` - List all tasks
- `create_task` - Create a new task
- `get_task` - Get full task details by entry ID
- `edit_task` - Edit an existing task
- `complete_task` - Mark a task as complete
- `delete_task` - Delete a task

## Architecture

```
Claude Code (WSL2/Linux)
    ↓ (JSON-RPC via stdio)
MCP Server (Windows Python)
    ↓ (COM)
Outlook Application
```

**How it works:**
1. Claude Code calls MCP tools via JSON-RPC
2. MCP server runs on Windows using `uv run --with pywin32`
3. Server uses COM to communicate with running Outlook instance
4. Results returned as JSON to Claude Code

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

### Testing the MCP Server Directly

```bash
# Test by sending JSON-RPC requests via stdin
echo '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"test","version":"1.0"}}}' | uv run --with pywin32 mcp_server.py

# List tools
echo '{"jsonrpc":"2.0","id":2,"method":"tools/list"}' | uv run --with pywin32 mcp_server.py
```

### Adding New Tools

1. Add the tool method in `mailtool/bridge.py`
2. Add tool schema in `mcp_server.py`'s `list_tools()` method
3. Add tool handler in `mcp_server.py`'s `call_tool()` method
4. Update documentation

## Security Considerations

- This MCP server has full access to your Outlook data
- Only install plugins from trusted sources
- The server runs with your Windows user permissions
- Email addresses and calendar data are transmitted to Claude

## License

MIT License - See LICENSE file for details

## Contributing

Contributions welcome! Please read CONTRIBUTING.md for guidelines.
