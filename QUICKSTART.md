# Mailtool MCP Quick Start Guide

Get Claude Code talking to your Outlook in 5 minutes!

## Step 1: Install the Plugin

```bash
# Navigate to your Claude Code plugins directory
cd ~/.claude-code/plugins

# Clone the repository
git clone <your-repo-url> mailtool

# Or if developing locally, copy/symlink the project
```

## Step 2: Verify Prerequisites

```bash
# Check that Outlook is running on Windows
# (Open Outlook manually if needed)

# Verify uv is installed
uv --version

# Test the MCP server
cd mailtool
python test_mcp_server.py
```

## Step 3: Restart Claude Code

Close and restart Claude Code to load the new plugin.

## Step 4: Start Using It!

### Basic Examples

In Claude Code, try:

```
Show me my last 5 emails
```

```
Create a task called "Test MCP integration" with high priority
```

```
List my calendar events for the next 7 days
```

```
Send an email to myself with subject "MCP test" and body "Testing Claude Code MCP integration"
```

## What Works Right Now

✅ **Email**: List, get details, send, reply, forward, mark read/unread, move, delete, search
✅ **Calendar**: List events, create appointments, get details, edit, respond to meetings, delete
✅ **Tasks**: List, create, get details, edit, complete, delete

## Common Issues

**"Server not initialized"**
→ Make sure Outlook is running on Windows

**"uv.exe not found"**
→ Install uv: `powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"`

**"Could not connect to Outlook"**
→ Start Outlook and log into your account

**Tools return errors**
→ Check the entry IDs are correct (use list commands first)

## Next Steps

- Read [MCP_INTEGRATION.md](MCP_INTEGRATION.md) for complete documentation
- See [COMMANDS.md](COMMANDS.md) for CLI usage (without Claude)
- Check [CLAUDE.md](CLAUDE.md) for architecture details

## Tips for Best Experience

1. **Keep Outlook running** - The MCP server connects to your running Outlook instance
2. **Use list commands first** - Get entry IDs from list operations before acting on specific items
3. **Be specific** - Claude works best with clear instructions like "Send an email to john@example.com with subject X and body Y"
4. **Test with CLI first** - Use `./outlook.sh` to verify operations work before trying via Claude

## Example Workflow

```
You: Show me unread emails from today

Claude: [Lists emails with entry IDs]

You: Reply to the first one with "Thanks for the update, I'll review it tomorrow"

Claude: [Sends reply using the entry ID]

You: Create a task to follow up on this

Claude: [Creates task in Outlook]

You: Mark that email as read

Claude: [Marks email as read]
```

Enjoy! 🚀
