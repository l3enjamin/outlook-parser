# MCP SDK v2 Migration Plan - Mailtool Outlook Automation

**Status**: DRAFT - Ready for Implementation
**Version Target**: MCP Python SDK v2 (alpha/main branch)
**Timeline**: 3-4 weeks (thorough refactor)
**Priority**: Production-ready with full feature parity

---

## Executive Summary

This document outlines a complete migration strategy from the current hand-rolled MCP implementation to the official **MCP Python SDK v2** using the FastMCP framework. The migration will:

✅ **Replace** manual JSON-RPC implementation with FastMCP decorators
✅ **Add** structured output using Pydantic models for type safety
✅ **Implement** MCP Resources for read-only data access
✅ **Preserve** all 23 existing tools with 100% feature parity
✅ **Improve** error handling with proper exception types
✅ **Maintain** Windows/COM compatibility and threading model
✅ **Enable** future enhancements (auth, monitoring, deployment)

### Key Benefits

- **~70% less code**: Declarative decorators vs imperative JSON-RPC
- **Type safety**: Automatic schema generation from type hints
- **Better patterns**: Resources for data, Tools for actions
- **Testability**: In-memory transport for unit testing
- **Production-ready**: Built-in logging, monitoring, deployment tools

---

## Table of Contents

1. [Current State Analysis](#1-current-state-analysis)
2. [Target Architecture](#2-target-architecture)
3. [Migration Strategy](#3-migration-strategy)
4. [Phase-by-Phase Implementation](#4-phase-by-phase-implementation)
5. [Pydantic Models Specification](#5-pydantic-models-specification)
6. [Tool Migration Catalog](#6-tool-migration-catalog)
7. [Resource Design](#7-resource-design)
8. [Testing Strategy](#8-testing-strategy)
9. [Dependencies & Configuration](#9-dependencies--configuration)
10. [Rollback Plan](#10-rollback-plan)
11. [Success Criteria](#11-success-criteria)

---

## 1. Current State Analysis

### 1.1 Implementation Overview

**File**: `C:\dev\mailtool\mcp_server.py` (856 lines)

**Current Approach**: Hand-written JSON-RPC server
- Manual request/response handling via stdio
- Manual tool schema definition (JSON Schema)
- Manual JSON serialization
- No structured output (all returns as JSON strings in TextContent)
- No lifespan management
- Basic error handling

**Tools**: 23 total across 3 domains
- Email: 9 tools (list, get, send, reply, forward, mark, move, delete, search)
- Calendar: 7 tools (list, create, get, edit, respond, delete, free_busy)
- Tasks: 7 tools (list, list_all, create, get, edit, complete, delete)

### 1.2 Bridge Layer (No Changes Needed)

**File**: `C:\dev\mailtool\src\mailtool\bridge.py` (1,108 lines)

**Status**: ✅ **Solid - No changes required**

The bridge layer provides a clean API that will work seamlessly with the SDK:
- All methods return Python dictionaries, strings, or booleans
- COM initialization is handled in `__init__`
- Error handling returns `False` or `None`
- O(1) access pattern via `GetItemFromID()`

**Key Pattern**: Bridge methods return simple Python types that can be converted to Pydantic models.

### 1.3 Critical Constraints

**COM Threading**: All COM calls must happen on the same thread (apartment-threaded)
- ✅ Current implementation is single-threaded async
- ✅ SDK preserves this pattern
- ⚠️ Must avoid multi-threading in tool implementations

**Dependency Management**: Uses `uv run --with pywin32`
- ✅ Keeps pywin32 out of Linux development
- ✅ Pattern works with SDK
- ⚠️ `plugin.json` must be updated to point at the new SDK server entry point

**Protocol Version**: Currently using `2024-11-05`
- ✅ SDK v2 supports this version
- ✅ Backward compatible with Claude Code

### 1.4 Verified Return Shapes & Behavioral Notes (from current code)

These items must be reflected in models and tool behavior to preserve parity:

- **Email details**: `get_email_body()` returns `entry_id`, `subject`, `sender`, `sender_name`, `body`, `html_body`, `received_time`, `has_attachments` **but not** `unread`.
- **Free/busy**: `get_free_busy()` can return `{email, error, resolved}` without `start_date`, `end_date`, or `free_busy` on failure paths.
- **Tasks**: `status` and `priority` can be `None` for some items; `percent_complete` is present but can be `0`.
- **list_tasks**: default is **incomplete-only** (`include_completed=False`), while `list_all_tasks()` returns all.
- **send_email**: bridge supports `file_paths` attachments; current MCP tool does **not** expose attachments.

---

## 2. Target Architecture

### 2.1 Technology Stack

**Core Framework**: MCP Python SDK v2 (main branch)
```toml
dependencies = [
    "mcp>=0.9.0",  # Runtime dependency for SDK server
    "pywin32>=306; sys_platform == 'win32'",
]
```

**Python Version**: Project currently targets Python 3.13; confirm MCP SDK v2 supports 3.13 before pinning.

**Pattern**: FastMCP decorator-based server
```python
from mcp.server.fastmcp import FastMCP

mcp = FastMCP("mailtool-outlook-bridge")

@mcp.tool()
def my_tool(param: str) -> ResultModel:
    """Tool with structured output"""
    return ResultModel(field=param)
```

### 2.2 New File Structure

```
mailtool/
├── src/mailtool/
│   ├── __init__.py
│   ├── bridge.py              # UNCHANGED
│   ├── cli.py                 # UNCHANGED
│   └── mcp/
│       ├── __init__.py         # NEW - MCP package
│       ├── server.py           # NEW - FastMCP server (replaces mcp_server.py)
│       ├── models.py           # NEW - Pydantic models
│       ├── resources.py        # NEW - MCP resource definitions
│       └── lifespan.py         # NEW - Lifecycle management
├── mcp_server.py               # OLD - Keep for reference, then remove
├── tests/
│   ├── test_mcp_server.py     # UPDATE - Test against new server
│   └── mcp/
│       ├── test_models.py      # NEW - Test Pydantic models
│       ├── test_resources.py   # NEW - Test MCP resources
│       └── test_tools.py       # NEW - Test all tools with SDK
└── pyproject.toml              # UPDATE - Add mcp dependency
```

### 2.3 Server Initialization Flow

```
1. Claude Code starts: uv run --with pywin32 -m mailtool.mcp.server
   │
2. FastMCP server initialized
   │
3. Lifespan startup:
   │  ├─ Create OutlookBridge instance
    │  ├─ Warmup connection with a real COM call (e.g., Inbox.Items.Count)
   │  └─ Store in lifespan context
   │
4. MCP initialize handshake
   │  ├─ Client sends initialize request
   │  └─ Server returns capabilities (tools, resources)
   │
5. Tool/Resource calls
   │  ├─ Access bridge from lifespan context
   │  ├─ Call bridge method
   │  ├─ Convert to Pydantic model
   │  └─ Return structured output
   │
6. Lifespan shutdown
   ├─ Release COM objects
   └─ Force garbage collection
```

---

## 3. Migration Strategy

### 3.1 Guiding Principles

1. **100% Feature Parity**: All 23 tools must work identically
2. **Structured Output**: Use Pydantic models for all returns
3. **Type Safety**: Leverage type hints for automatic schema generation
4. **Resources for Data**: Expose read-only data as MCP resources
5. **Preserve COM Patterns**: Maintain apartment-threading model
6. **Test-Driven**: Write tests alongside implementation
7. **Documentation First**: Update docs as we go

### 3.2 Migration Phases

**Week 1: Foundation** (Phases 1-3)
- Set up SDK infrastructure
- Define Pydantic models
- Implement lifespan management
- Migrate 5-10 simple tools

**Week 2: Core Tools** (Phase 4)
- Migrate all email tools (9)
- Add email resources
- Test with Claude Code

**Week 3: Advanced Tools** (Phase 5)
- Migrate calendar tools (7)
- Migrate task tools (7)
- Add calendar and task resources
- Full integration testing

**Week 4: Polish & Deploy** (Phase 6)
- Performance optimization
- Error handling refinement
- Documentation updates
- Production deployment

### 3.3 Parallel Development Approach

**Keep both implementations during migration:**
- `mcp_server.py` - OLD (hand-rolled)
- `src/mailtool/mcp/server.py` - NEW (SDK-based)

**Plugin configuration**:
```json
{
  "mcpServers": {
    "mailtool-new": {
      "command": "uv",
      "args": ["run", "--with", "pywin32", "-m", "mailtool.mcp.server"]
    }
  }
}
```

**Cutover**: Update plugin.json once fully validated

---

## 4. Phase-by-Phase Implementation

### Phase 1: Foundation (Days 1-2)

**Objective**: Set up SDK infrastructure and basic server

**Tasks**:
1. Add `mcp` dependency to `pyproject.toml`
2. Create `src/mailtool/mcp/` package structure
3. Implement basic FastMCP server with lifespan
4. Write simple test tool to validate setup
5. Update CI/CD to install `mcp`

**Code Changes**:

**pyproject.toml**:
```toml
[project]
dependencies = [
    "mcp>=0.9.0",
    "pywin32>=306; sys_platform == 'win32'",
]

[dependency-groups]
dev = [
    "ruff>=0.9.0",
    "pytest>=7.0",
    "pytest-asyncio>=0.21",
    "pytest-timeout>=2.2",
]
```

> If you prefer the existing pattern (`uv run --with pytest --with pytest-timeout`), these can stay out of `dependency-groups.dev`.

**src/mailtool/mcp/__init__.py**:
```python
"""MCP server package for mailtool."""

__version__ = "2.2.0"
```

**src/mailtool/mcp/lifespan.py**:
```python
"""Lifespan management for MCP server."""

from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from dataclasses import dataclass

from mailtool.bridge import OutlookBridge
from mcp.server.fastmcp import FastMCP


@dataclass
class OutlookContext:
    """Application context with Outlook bridge."""
    bridge: OutlookBridge


@asynccontextmanager
async def outlook_lifespan(server: FastMCP) -> AsyncIterator[OutlookContext]:
    """Manage Outlook bridge lifecycle.

    Startup:
    - Create OutlookBridge instance
    - Warmup connection (COM can be slow on first call)

    Shutdown:
    - Release COM objects
    - Force garbage collection
    """
    import asyncio
    import gc
    import time

    # Startup
    bridge = OutlookBridge()

    # Warmup - ensure Outlook is responsive (real COM call, retry like tests)
    for attempt in range(5):
        try:
            inbox = bridge.get_inbox()
            _ = inbox.Items.Count
            break
        except Exception:
            if attempt < 4:
                await asyncio.sleep(0.5)
            else:
                raise

    try:
        yield OutlookContext(bridge=bridge)
    finally:
        # Shutdown
        del bridge
        gc.collect()
```

**src/mailtool/mcp/server.py** (skeleton):
```python
"""FastMCP server for mailtool."""

from mcp.server.fastmcp import FastMCP
from .lifespan import outlook_lifespan

# Create server with lifespan
mcp = FastMCP(
    "mailtool-outlook-bridge",
    lifespan=outlook_lifespan
)

# TODO: Add tools, resources, prompts

if __name__ == "__main__":
    mcp.run()
```

**Testing**:
```bash
# Test server starts
uv run --with mcp --with pywin32 python -m mailtool.mcp.server

# Should see: "Starting MCP server on stdio"
```

**Success Criteria**:
- ✅ Server starts without errors
- ✅ Outlook bridge initializes
- ✅ Clean shutdown works

---

### Phase 2: Pydantic Models (Days 3-4)

**Objective**: Define all data models for structured output

**File**: `src/mailtool/mcp/models.py`

**Models to Define**:

```python
"""Pydantic models for MCP structured output."""

from datetime import datetime
from pydantic import BaseModel, Field
from typing import Optional


# ============================================================================
# EMAIL MODELS
# ============================================================================

class EmailSummary(BaseModel):
    """Summary of an email (for list operations)."""
    entry_id: str = Field(description="Outlook EntryID (unique identifier)")
    subject: str = Field(description="Email subject line")
    sender: str = Field(description="Sender SMTP email address")
    sender_name: str = Field(description="Sender display name")
    received_time: Optional[str] = Field(
        default=None,
        description="When email was received (YYYY-MM-DD HH:MM:SS)"
    )
    unread: Optional[bool] = Field(
        default=None,
        description="True if email is unread (list_emails always sets this)"
    )
    has_attachments: bool = Field(description="True if email has attachments")


class EmailDetails(EmailSummary):
    """Full email details including body (for get operations)."""
    body: str = Field(default="", description="Plain text email body")
    html_body: str = Field(default="", description="HTML email body")


class SendEmailResult(BaseModel):
    """Result of sending an email."""
    success: bool = Field(description="True if email was sent")
    entry_id: Optional[str] = Field(
        default=None,
        description="EntryID if saved as draft, None if sent"
    )
    message: str = Field(description="Human-readable result message")


# ============================================================================
# CALENDAR MODELS
# ============================================================================

class AppointmentSummary(BaseModel):
    """Summary of a calendar event."""
    entry_id: str = Field(description="Outlook EntryID")
    subject: str = Field(description="Appointment subject")
    start: Optional[str] = Field(
        default=None,
        description="Start time (YYYY-MM-DD HH:MM:SS)"
    )
    end: Optional[str] = Field(
        default=None,
        description="End time (YYYY-MM-DD HH:MM:SS)"
    )
    location: str = Field(default="", description="Appointment location")
    organizer: Optional[str] = Field(
        default=None,
        description="Organizer email address"
    )
    all_day: bool = Field(description="True if all-day event")
    required_attendees: str = Field(
        default="",
        description="Semicolon-separated list of required attendees"
    )
    optional_attendees: str = Field(
        default="",
        description="Semicolon-separated list of optional attendees"
    )
    response_status: str = Field(
        description="Meeting response status for user"
    )
    meeting_status: str = Field(description="Meeting status")
    response_requested: bool = Field(description="True if response requested")


class AppointmentDetails(AppointmentSummary):
    """Full appointment details including body."""
    body: str = Field(default="", description="Appointment description/body")


class CreateAppointmentResult(BaseModel):
    """Result of creating an appointment."""
    success: bool = Field(description="True if appointment created")
    entry_id: Optional[str] = Field(
        default=None,
        description="EntryID of created appointment"
    )
    message: str = Field(description="Result message")


class FreeBusyInfo(BaseModel):
    """Free/busy information for an email address."""
    email: str = Field(description="Email address checked")
    start_date: Optional[str] = Field(
        default=None,
        description="Start date (YYYY-MM-DD)"
    )
    end_date: Optional[str] = Field(
        default=None,
        description="End date (YYYY-MM-DD)"
    )
    free_busy: Optional[str] = Field(
        default=None,
        description="Free/busy string with time slot status codes"
    )
    resolved: bool = Field(description="True if email was resolved")
    error: Optional[str] = Field(default=None, description="Error message if failed")


# ============================================================================
# TASK MODELS
# ============================================================================

class TaskSummary(BaseModel):
    """Summary of a task."""
    entry_id: str = Field(description="Outlook EntryID")
    subject: str = Field(description="Task subject")
    body: str = Field(default="", description="Task description")
    due_date: Optional[str] = Field(
        default=None,
        description="Due date (YYYY-MM-DD)"
    )
    status: Optional[int] = Field(
        default=None,
        description="Task status (0=Not started, 1=In progress, 2=Complete)"
    )
    priority: Optional[int] = Field(
        default=None,
        description="Task priority (0=Low, 1=Normal, 2=High)"
    )
    complete: bool = Field(description="True if task is complete")
    percent_complete: int = Field(
        default=0,
        description="Percent complete (0-100)"
    )


class CreateTaskResult(BaseModel):
    """Result of creating a task."""
    success: bool = Field(description="True if task created")
    entry_id: Optional[str] = Field(
        default=None,
        description="EntryID of created task"
    )
    message: str = Field(description="Result message")


# ============================================================================
# COMMON MODELS
# ============================================================================

class OperationResult(BaseModel):
    """Generic result for operations that return True/False."""
    success: bool = Field(description="True if operation succeeded")
    message: str = Field(description="Human-readable result message")
```

**Validation Tasks**:
1. All models align with bridge return types
2. Field descriptions are clear for LLM understanding
3. Optional fields match bridge behavior (None vs empty string)
4. Enums match bridge constants (status, priority, etc.)

**Testing**:
```python
# tests/mcp/test_models.py
def test_email_summary_from_bridge():
    """Test EmailSummary can be created from bridge output."""
    from mailtool.mcp.models import EmailSummary

    bridge_data = {
        "entry_id": "12345",
        "subject": "Test",
        "sender": "test@example.com",
        "sender_name": "Test User",
        "received_time": "2025-01-18 10:00:00",
        "unread": True,
        "has_attachments": False
    }

    summary = EmailSummary(**bridge_data)
    assert summary.entry_id == "12345"
```

**Success Criteria**:
- ✅ All models defined with proper validation
- ✅ Tests pass for model creation
- ✅ Field descriptions are LLM-friendly

---

### Phase 3: Tool Migration - Simple Tools (Days 5-7)

**Objective**: Migrate 5-10 simple tools to validate patterns

**Tools to Migrate**:
1. `get_email` - Simple O(1) lookup
2. `get_appointment` - Simple O(1) lookup
3. `get_task` - Simple O(1) lookup
4. `mark_email` - Simple boolean operation
5. `complete_task` - Simple boolean operation
6. `delete_email` - Simple boolean operation
7. `delete_appointment` - Simple boolean operation
8. `delete_task` - Simple boolean operation

**Pattern Example**:

```python
# src/mailtool/mcp/server.py

from mcp.server.fastmcp import Context, FastMCP
from mcp.server.session import ServerSession
from mailtool.mcp.models import EmailDetails
from mailtool.mcp.lifespan import OutlookContext

mcp = FastMCP("mailtool-outlook-bridge", lifespan=outlook_lifespan)


@mcp.tool()
def get_email(
    entry_id: str,
    ctx: Context[ServerSession, OutlookContext]
) -> EmailDetails:
    """Get full email details including body.

    Args:
        entry_id: Outlook EntryID of the email

    Returns:
        EmailDetails with full email information

    Raises:
        McpError: If email not found or COM error occurs
    """
    from mcp.shared.exceptions import McpError

    bridge = ctx.request_context.lifespan_context.bridge
    result = bridge.get_email_body(entry_id)

    if not result:
        raise McpError(f"Email not found: {entry_id}")

    # Convert bridge dict to Pydantic model
    return EmailDetails(
        entry_id=result["entry_id"],
        subject=result["subject"],
        sender=result["sender"],
        sender_name=result["sender_name"],
        unread=result.get("unread"),
        body=result.get("body", ""),
        html_body=result.get("html_body", ""),
        received_time=result.get("received_time"),
        has_attachments=result["has_attachments"]
    )


@mcp.tool()
def mark_email(
    entry_id: str,
    unread: bool = False,
    ctx: Context[ServerSession, OutlookContext]
) -> OperationResult:
    """Mark an email as read or unread.

    Args:
        entry_id: Email EntryID
        unread: True to mark as unread, False to mark as read

    Returns:
        OperationResult indicating success or failure
    """
    bridge = ctx.request_context.lifespan_context.bridge
    result = bridge.mark_email_read(entry_id, unread=unread)

    if result:
        return OperationResult(
            success=True,
            message=f"Email marked as {'unread' if unread else 'read'}"
        )
    else:
        from mcp.shared.exceptions import McpError
        raise McpError(f"Failed to mark email: {entry_id}")


@mcp.tool()
def delete_email(
    entry_id: str,
    ctx: Context[ServerSession, OutlookContext]
) -> OperationResult:
    """Delete an email.

    Args:
        entry_id: Email EntryID to delete

    Returns:
        OperationResult indicating success or failure
    """
    bridge = ctx.request_context.lifespan_context.bridge
    result = bridge.delete_email(entry_id)

    if result:
        return OperationResult(
            success=True,
            message="Email deleted successfully"
        )
    else:
        from mcp.shared.exceptions import McpError
        raise McpError(f"Failed to delete email: {entry_id}")
```

**Testing**:
```python
# tests/mcp/test_tools.py
import pytest
from mailtool.mcp.server import mcp

@pytest.mark.asyncio
async def test_get_email_success(test_entry_id):
    """Test get_email returns structured data."""
    # Call tool via SDK
    result = await mcp.call_tool("get_email", {"entry_id": test_entry_id})

    # Validate structured output
    assert hasattr(result, 'structured_content')
    assert result.structured_content['subject'] is not None
```

**Success Criteria**:
- ✅ 8 tools migrated successfully
- ✅ Tests pass for all tools
- ✅ Manual testing with Claude Code works

---

### Phase 4: Email Tools & Resources (Days 8-10)

**Objective**: Migrate all email tools and add email resources

**Email Tools to Migrate**:
1. ✅ `get_email` (done in Phase 3)
2. `list_emails` - List with limit/folder
3. `send_email` - Create with all parameters
4. `reply_email` - Reply/reply all
5. `forward_email` - Forward to recipient
6. ✅ `mark_email` (done in Phase 3)
7. `move_email` - Move to folder
8. ✅ `delete_email` (done in Phase 3)
9. `search_emails` - Search with filter

**Complex Tool Example**:

```python
@mcp.tool()
def list_emails(
    limit: int = 10,
    folder: str = "Inbox",
    ctx: Context[ServerSession, OutlookContext]
) -> list[EmailSummary]:
    """List recent emails from a folder.

    Args:
        limit: Maximum number of emails to return (default: 10)
        folder: Folder name (default: "Inbox")

    Returns:
        List of EmailSummary objects
    """
    bridge = ctx.request_context.lifespan_context.bridge
    emails = bridge.list_emails(limit=limit, folder=folder)

    # Convert list of dicts to list of Pydantic models
    return [
        EmailSummary(
            entry_id=e["entry_id"],
            subject=e["subject"],
            sender=e["sender"],
            sender_name=e["sender_name"],
            received_time=e.get("received_time"),
            unread=e["unread"],
            has_attachments=e["has_attachments"]
        )
        for e in emails
    ]


@mcp.tool()
def send_email(
    to: str,
    subject: str,
    body: str,
    cc: Optional[str] = None,
    bcc: Optional[str] = None,
    html_body: Optional[str] = None,
    file_paths: Optional[list[str]] = None,
    save_draft: bool = False,
    ctx: Context[ServerSession, OutlookContext]
) -> SendEmailResult:
    """Send a new email or save as draft.

    Args:
        to: Recipient email address
        subject: Email subject
        body: Plain text body (required if html_body not provided)
        cc: CC recipients (optional)
        bcc: BCC recipients (optional)
        html_body: HTML body (optional, overrides body)
        save_draft: Save to Drafts instead of sending (default: False)

    Returns:
        SendEmailResult with success status and EntryID if draft
    """
    bridge = ctx.request_context.lifespan_context.bridge
    result = bridge.send_email(
        to=to,
        subject=subject,
        body=body,
        cc=cc,
        bcc=bcc,
        html_body=html_body,
        file_paths=file_paths,
        save_draft=save_draft
    )

    # Handle different return types
    if result is False:
        return SendEmailResult(
            success=False,
            entry_id=None,
            message="Failed to send email"
        )
    elif isinstance(result, str):
        # Draft was saved, result is EntryID
        return SendEmailResult(
            success=True,
            entry_id=result,
            message="Draft saved successfully"
        )
    else:
        # Email was sent, result is True
        return SendEmailResult(
            success=True,
            entry_id=None,
            message="Email sent successfully"
        )


@mcp.tool()
def search_emails(
    filter_query: str,
    limit: int = 100,
    ctx: Context[ServerSession, OutlookContext]
) -> list[EmailSummary]:
    """Search emails using Outlook filter query.

    Args:
        filter_query: SQL-like filter (e.g., "[Subject] = 'meeting'")
        limit: Maximum results (default: 100)

    Returns:
        List of EmailSummary objects matching the query

    Examples:
        search_emails("[Subject] = 'invoice'")
        search_emails("urn:schemas:httpmail:subject LIKE '%test%'")
    """
    bridge = ctx.request_context.lifespan_context.bridge
    emails = bridge.search_emails(filter_query=filter_query, limit=limit)

    return [
        EmailSummary(
            entry_id=e["entry_id"],
            subject=e["subject"],
            sender=e["sender"],
            sender_name=e["sender_name"],
            received_time=e.get("received_time"),
            unread=e["unread"],
            has_attachments=e["has_attachments"]
        )
        for e in emails
    ]
```

**Email Resources**:

```python
# src/mailtool/mcp/resources.py

from mcp.server.fastmcp import Context, FastMCP
from mcp.server.session import ServerSession
from mailtool.mcp.models import EmailDetails
from mailtool.mcp.lifespan import OutlookContext
import json

def register_email_resources(mcp: FastMCP):
    """Register email-related MCP resources."""

    @mcp.resource("inbox://emails")
    def get_inbox_emails(ctx: Context[ServerSession, OutlookContext]) -> str:
        """Get current inbox emails as JSON resource.

        This resource provides read-only access to recent emails.
        Use list_emails tool for more control.
        """
        bridge = ctx.request_context.lifespan_context.bridge
        emails = bridge.list_emails(limit=50)
        return json.dumps(emails, indent=2)

    @mcp.resource("inbox://unread")
    def get_unread_emails(ctx: Context[ServerSession, OutlookContext]) -> str:
        """Get unread emails as JSON resource."""
        bridge = ctx.request_context.lifespan_context.bridge

        # Get all emails and filter unread
        all_emails = bridge.list_emails(limit=1000)
        unread = [e for e in all_emails if e["unread"]]

        return json.dumps(unread, indent=2)

    @mcp.resource("email://{entry_id}")
    def get_email_resource(entry_id: str, ctx: Context[ServerSession, OutlookContext]) -> str:
        """Get a specific email by EntryID as JSON resource.

        Args:
            entry_id: Outlook EntryID
        """
        bridge = ctx.request_context.lifespan_context.bridge
        email = bridge.get_email_body(entry_id)

        if not email:
            return json.dumps({"error": "Email not found"})

        return json.dumps(email, indent=2)
```

**Server Integration**:

```python
# src/mailtool/mcp/server.py

from .resources import register_email_resources

# Register resources
register_email_resources(mcp)
```

**Success Criteria**:
- ✅ All 9 email tools migrated
- ✅ 3 email resources defined
- ✅ Email tools tested with Claude Code
- ✅ Resources accessible via MCP

---

### Phase 5: Calendar & Task Tools (Days 11-14)

**Objective**: Migrate remaining tools and add resources

**Calendar Tools (7)**:
1. `list_calendar_events` - List with days/all filter
2. `create_appointment` - Create with attendees
3. ✅ `get_appointment` (done in Phase 3)
4. `edit_appointment` - Edit existing
5. `respond_to_meeting` - Accept/decline/tentative
6. ✅ `delete_appointment` (done in Phase 3)
7. `get_free_busy` - Check availability

**Task Tools (7)**:
1. `list_tasks` - List incomplete only
2. `list_all_tasks` - List all tasks
3. `create_task` - Create with priority
4. ✅ `get_task` (done in Phase 3)
5. `edit_task` - Edit with percent_complete
6. ✅ `complete_task` (done in Phase 3)
7. ✅ `delete_task` (done in Phase 3)

**Example Calendar Tool**:

```python
@mcp.tool()
def list_calendar_events(
    days: int = 7,
    all: bool = False,
    ctx: Context[ServerSession, OutlookContext]
) -> list[AppointmentSummary]:
    """List calendar events for the next N days or all events.

    Args:
        days: Number of days ahead to look (default: 7)
        all: Return all events without date filtering (default: False)

    Returns:
        List of AppointmentSummary objects
    """
    bridge = ctx.request_context.lifespan_context.bridge
    events = bridge.list_calendar_events(days=days, all_events=all)

    return [
        AppointmentSummary(
            entry_id=e["entry_id"],
            subject=e["subject"],
            start=e.get("start"),
            end=e.get("end"),
            location=e["location"],
            organizer=e.get("organizer"),
            all_day=e["all_day"],
            required_attendees=e["required_attendees"],
            optional_attendees=e["optional_attendees"],
            response_status=e["response_status"],
            meeting_status=e["meeting_status"],
            response_requested=e["response_requested"]
        )
        for e in events
    ]


@mcp.tool()
def create_appointment(
    subject: str,
    start: str,
    end: str,
    location: str = "",
    body: str = "",
    all_day: bool = False,
    required_attendees: Optional[str] = None,
    optional_attendees: Optional[str] = None,
    ctx: Context[ServerSession, OutlookContext]
) -> CreateAppointmentResult:
    """Create a new calendar appointment.

    Args:
        subject: Appointment subject
        start: Start time (YYYY-MM-DD HH:MM:SS)
        end: End time (YYYY-MM-DD HH:MM:SS)
        location: Location (optional)
        body: Description/body (optional)
        all_day: All-day event (default: False)
        required_attendees: Semicolon-separated list (optional)
        optional_attendees: Semicolon-separated list (optional)

    Returns:
        CreateAppointmentResult with EntryID if successful
    """
    bridge = ctx.request_context.lifespan_context.bridge
    entry_id = bridge.create_appointment(
        subject=subject,
        start=start,
        end=end,
        location=location,
        body=body,
        all_day=all_day,
        required_attendees=required_attendees,
        optional_attendees=optional_attendees
    )

    if entry_id:
        return CreateAppointmentResult(
            success=True,
            entry_id=entry_id,
            message="Appointment created successfully"
        )
    else:
        return CreateAppointmentResult(
            success=False,
            entry_id=None,
            message="Failed to create appointment"
        )
```

**Calendar Resources**:

```python
def register_calendar_resources(mcp: FastMCP):
    """Register calendar-related MCP resources."""

    @mcp.resource("calendar://today")
    def get_today_calendar(ctx: Context[ServerSession, OutlookContext]) -> str:
        """Get today's calendar events as JSON resource."""
        bridge = ctx.request_context.lifespan_context.bridge
        events = bridge.list_calendar_events(days=1)
        return json.dumps(events, indent=2)

    @mcp.resource("calendar://week")
    def get_week_calendar(ctx: Context[ServerSession, OutlookContext]) -> str:
        """Get this week's calendar events as JSON resource."""
        bridge = ctx.request_context.lifespan_context.bridge
        events = bridge.list_calendar_events(days=7)
        return json.dumps(events, indent=2)
```

**Task Resources**:

```python
def register_task_resources(mcp: FastMCP):
    """Register task-related MCP resources."""

    @mcp.resource("tasks://active")
    def get_active_tasks(ctx: Context[ServerSession, OutlookContext]) -> str:
        """Get active (incomplete) tasks as JSON resource."""
        bridge = ctx.request_context.lifespan_context.bridge
        tasks = bridge.list_tasks(include_completed=False)
        return json.dumps(tasks, indent=2)

    @mcp.resource("tasks://all")
    def get_all_tasks(ctx: Context[ServerSession, OutlookContext]) -> str:
        """Get all tasks as JSON resource."""
        bridge = ctx.request_context.lifespan_context.bridge
        tasks = bridge.list_tasks(include_completed=True)
        return json.dumps(tasks, indent=2)
```

**Success Criteria**:
- ✅ All 23 tools migrated
- ✅ All resources registered
- ✅ Full test suite passes
- ✅ Claude Code integration validated

---

### Phase 6: Polish & Deploy (Days 15-20)

**Objective**: Production-ready deployment

**Tasks**:

1. **Performance Optimization**
   - Batch operations where possible
   - Cache frequently accessed data
   - Optimize COM object cleanup

2. **Error Handling Refinement**
   - Custom exception types for common errors
   - User-friendly error messages
   - Proper logging for debugging

3. **Documentation Updates**
   - Update README.md with SDK patterns
   - Update CLAUDE.md with new architecture
   - Update MCP_INTEGRATION.md
   - Add migration notes

4. **Testing**
   - Full integration test suite
   - Load testing with large folders
   - Memory leak testing
   - COM object cleanup validation

5. **Deployment**
   - Update plugin.json
   - Test in production Claude Code
   - Monitor for issues
   - Keep old version as rollback

**Error Handling Improvements**:

```python
# src/mailtool/mcp/exceptions.py

"""Custom exceptions for mailtool MCP server."""

from mcp.shared.exceptions import McpError


class OutlookNotFoundError(McpError):
    """Raised when an Outlook item is not found."""
    pass


class OutlookComError(McpError):
    """Raised when a COM operation fails."""
    pass


class OutlookValidationError(McpError):
    """Raised when input validation fails."""
    pass


# Usage in tools
@mcp.tool()
def get_email(entry_id: str, ctx: Context) -> EmailDetails:
    """Get email with custom error handling."""
    try:
        bridge = ctx.request_context.lifespan_context.bridge
        result = bridge.get_email_body(entry_id)

        if not result:
            raise OutlookNotFoundError(f"Email not found: {entry_id}")

        return EmailDetails(**result)

    except Exception as e:
        raise OutlookComError(f"Failed to retrieve email: {str(e)}")
```

**Documentation Updates**:

**README.md** - Add new section:
```markdown
## MCP Server (v2 with FastMCP)

The MCP server now uses the official MCP Python SDK v2 with FastMCP framework.

### Features
- **Structured Output**: All tools return validated Pydantic models
- **Resources**: Read-only access to emails, calendar, tasks
- **Type Safety**: Automatic schema generation from type hints
- **Better Errors**: User-friendly error messages

### Migration
See [MCP_SDK_V2_MIGRATION_PLAN.md](MCP_SDK_V2_MIGRATION_PLAN.md) for details.
```

**Plugin Configuration Update**:

```json
{
  "name": "mailtool-outlook-bridge",
  "version": "2.3.0",
  "mcpServers": {
    "mailtool": {
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
}
```

**Success Criteria**:
- ✅ All documentation updated
- ✅ Performance benchmarks acceptable
- ✅ Error handling comprehensive
- ✅ Production deployment successful
- ✅ Rollback plan tested

---

## 5. Pydantic Models Specification

See Phase 2 for complete model definitions.

**Key Design Decisions**:
1. **Separate Summary vs Details models** - For list vs get operations
2. **Optional fields** - Match bridge behavior (None vs empty string)
3. **Field descriptions** - Critical for LLM understanding
4. **Result models** - For operations that return success/failure

---

## 6. Tool Migration Catalog

### 6.1 Tool Mapping Table

| Old Tool | New Tool | Return Type | Complexity |
|----------|----------|-------------|------------|
| `list_emails` | `list_emails` | `list[EmailSummary]` | Medium |
| `get_email` | `get_email` | `EmailDetails` | Simple ✅ |
| `send_email` | `send_email` | `SendEmailResult` | Complex |
| `reply_email` | `reply_email` | `OperationResult` | Medium |
| `forward_email` | `forward_email` | `OperationResult` | Medium |
| `mark_email` | `mark_email` | `OperationResult` | Simple ✅ |
| `move_email` | `move_email` | `OperationResult` | Medium |
| `delete_email` | `delete_email` | `OperationResult` | Simple ✅ |
| `search_emails` | `search_emails` | `list[EmailSummary]` | Medium |
| `list_calendar_events` | `list_calendar_events` | `list[AppointmentSummary]` | Medium |
| `create_appointment` | `create_appointment` | `CreateAppointmentResult` | Complex |
| `get_appointment` | `get_appointment` | `AppointmentDetails` | Simple ✅ |
| `edit_appointment` | `edit_appointment` | `OperationResult` | Complex |
| `respond_to_meeting` | `respond_to_meeting` | `OperationResult` | Medium |
| `delete_appointment` | `delete_appointment` | `OperationResult` | Simple ✅ |
| `get_free_busy` | `get_free_busy` | `FreeBusyInfo` | Medium |
| `list_tasks` | `list_tasks` | `list[TaskSummary]` | Simple |
| `list_all_tasks` | `list_all_tasks` | `list[TaskSummary]` | Simple |
| `create_task` | `create_task` | `CreateTaskResult` | Medium |
| `get_task` | `get_task` | `TaskSummary` | Simple ✅ |
| `edit_task` | `edit_task` | `OperationResult` | Complex |
| `complete_task` | `complete_task` | `OperationResult` | Simple ✅ |
| `delete_task` | `delete_task` | `OperationResult` | Simple ✅ |

### 6.2 Tool Complexity Breakdown

**Simple (8 tools)** - Direct bridge calls, minimal logic:
- get_email, get_appointment, get_task
- mark_email, delete_email, delete_appointment
- complete_task, delete_task

**Medium (9 tools)** - Some conversion or error handling:
- list_emails, list_calendar_events, list_tasks
- list_all_tasks
- reply_email, forward_email, move_email
- search_emails, respond_to_meeting, get_free_busy

**Complex (6 tools)** - Multiple parameters, conditional logic:
- send_email (draft vs sent, multiple optional params)
- create_appointment (attendees, all_day flag)
- edit_appointment (many optional params)
- edit_task (automatic status updates)
- create_task (priority mapping)

---

## 7. Resource Design

### 7.1 Resource URI Patterns

```
email://{entry_id}          # Single email
inbox://emails              # Recent inbox emails
inbox://unread              # Unread emails
calendar://today            # Today's events
calendar://week             # This week's events
calendar://{date}           # Specific date (YYYY-MM-DD)
tasks://active              # Active tasks
tasks://all                 # All tasks
tasks://overdue             # Overdue tasks
```

### 7.2 Resource Implementations

See Phase 4 and Phase 5 for email, calendar, and task resource examples.

### 7.3 Resource vs Tool Guidance

**When to Use Resources**:
- Read-only data access
- Frequently accessed data (caching potential)
- Large datasets (can be streamed)
- Background data (context for LLM)

**When to Use Tools**:
- Actions with side effects
- Parameterized queries
- Operations that modify data
- Long-running operations

---

## 8. Testing Strategy

### 8.1 Test Structure

```
tests/
├── conftest.py              # Existing fixtures (keep)
├── test_bridge.py           # Existing (keep)
├── test_emails.py           # Existing (keep)
├── test_calendar.py         # Existing (keep)
├── test_tasks.py            # Existing (keep)
├── test_mcp_server.py       # UPDATE - Test SDK server
└── mcp/
    ├── test_models.py       # NEW - Test Pydantic models
    ├── test_resources.py    # NEW - Test MCP resources
    ├── test_tools.py        # NEW - Test all tools
    └── conftest.py          # NEW - MCP-specific fixtures
```

### 8.2 Test Fixtures

```python
# tests/mcp/conftest.py

"""MCP server test fixtures."""

import pytest
from mailtool.mcp.server import mcp
from mailtool.bridge import OutlookBridge


@pytest.fixture
def mcp_server():
    """Provide MCP server instance."""
    return mcp


@pytest.fixture
def outlook_bridge():
    """Provide OutlookBridge instance."""
    return OutlookBridge()


@pytest.fixture
def sample_email(outlook_bridge):
    """Create a sample email for testing."""
    entry_id = outlook_bridge.send_email(
        to="test@example.com",
        subject="[TEST] Sample Email",
        body="Test body",
        save_draft=True
    )
    yield entry_id

    # Cleanup
    try:
        outlook_bridge.delete_email(entry_id)
    except:
        pass
```

### 8.3 Tool Testing Pattern

```python
# tests/mcp/test_tools.py

"""Test MCP tools with SDK client."""

import pytest
from mcp import ClientSession
from mcp.client.stdio import stdio_client, StdioServerParameters


@pytest.mark.asyncio
async def test_list_emails():
    """Test list_emails tool."""
    server_params = StdioServerParameters(
        command="uv",
        args=["run", "--with", "pywin32", "-m", "mailtool.mcp.server"]
    )

    async with stdio_client(server_params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()

            # Call tool
            result = await session.call_tool(
                "list_emails",
                {"limit": 5, "folder": "Inbox"}
            )

            # Validate structured output
            assert hasattr(result, 'structured_content')
            assert isinstance(result.structured_content, list)
            assert len(result.structured_content) <= 5


@pytest.mark.asyncio
async def test_get_email_not_found():
    """Test get_email with invalid EntryID."""
    server_params = StdioServerParameters(
        command="uv",
        args=["run", "--with", "pywin32", "-m", "mailtool.mcp.server"]
    )

    async with stdio_client(server_params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()

            # Should raise error for invalid EntryID
            with pytest.raises(Exception):  # McpError
                await session.call_tool(
                    "get_email",
                    {"entry_id": "invalid_entry_id"}
                )
```

### 8.4 Model Testing Pattern

```python
# tests/mcp/test_models.py

"""Test Pydantic models."""

from mailtool.mcp.models import EmailSummary, EmailDetails
import pytest


def test_email_summary_valid():
    """Test EmailSummary with valid data."""
    data = {
        "entry_id": "12345",
        "subject": "Test",
        "sender": "test@example.com",
        "sender_name": "Test",
        "received_time": "2025-01-18 10:00:00",
        "unread": True,
        "has_attachments": False
    }

    summary = EmailSummary(**data)
    assert summary.entry_id == "12345"
    assert summary.unread is True


def test_email_summary_missing_optional():
    """Test EmailSummary with missing optional fields."""
    data = {
        "entry_id": "12345",
        "subject": "Test",
        "sender": "test@example.com",
        "sender_name": "Test",
        "received_time": None,
        "unread": False,
        "has_attachments": False
    }

    summary = EmailSummary(**data)
    assert summary.received_time is None
```

### 8.5 Resource Testing Pattern

```python
# tests/mcp/test_resources.py

"""Test MCP resources."""

import pytest
from mcp import ClientSession
from mcp.client.stdio import stdio_client, StdioServerParameters


@pytest.mark.asyncio
async def test_inbox_emails_resource():
    """Test inbox://emails resource."""
    server_params = StdioServerParameters(
        command="uv",
        args=["run", "--with", "pywin32", "-m", "mailtool.mcp.server"]
    )

    async with stdio_client(server_params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()

            # Read resource
            result = await session.read_resource("inbox://emails")

            # Validate
            assert len(result.contents) > 0
            assert result.contents[0].type == "text"
```

### 8.6 Integration Testing

```python
# tests/mcp/test_integration.py

"""End-to-end integration tests."""

import pytest
from mcp import ClientSession
from mcp.client.stdio import stdio_client, StdioServerParameters


@pytest.mark.asyncio
@pytest.mark.integration
async def test_full_email_workflow():
    """Test complete email workflow."""
    server_params = StdioServerParameters(
        command="uv",
        args=["run", "--with", "pywin32", "-m", "mailtool.mcp.server"]
    )

    async with stdio_client(server_params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()

            # 1. List emails
            emails_result = await session.call_tool(
                "list_emails",
                {"limit": 1}
            )

            # 2. Create email
            send_result = await session.call_tool(
                "send_email",
                {
                    "to": "test@example.com",
                    "subject": "[TEST] Integration Test",
                    "body": "Test",
                    "save_draft": True
                }
            )

            # 3. Get the draft
            entry_id = send_result.structured_content['entry_id']
            email_result = await session.call_tool(
                "get_email",
                {"entry_id": entry_id}
            )

            # 4. Delete the draft
            delete_result = await session.call_tool(
                "delete_email",
                {"entry_id": entry_id}
            )

            assert delete_result.structured_content['success'] is True
```

---

## 9. Dependencies & Configuration

### 9.1 Dependency Changes

**pyproject.toml**:
```toml
[project]
name = "mailtool"
version = "2.3.0"  # Bump version
requires-python = ">=3.13"
dependencies = [
    "pywin32>=306; sys_platform == 'win32'",
]

[dependency-groups]
dev = [
    "ruff>=0.9.0",
    "pytest>=7.0",
    "pytest-asyncio>=0.21",  # For async tests
    "mcp>=0.9.0",  # Add MCP SDK for development
]
```

**Install commands**:
```bash
# Add MCP SDK
uv add --group dev mcp

# Sync dependencies
uv sync --all-groups
```

### 9.2 Plugin Configuration

**.claude-plugin/plugin.json**:
```json
{
  "name": "mailtool-outlook-bridge",
  "version": "2.3.0",
  "description": "Outlook automation via COM with MCP SDK v2",
  "author": "Sam <s.mok@utwente.nl>",
  "license": "MIT",
  "requires": {
    "claude-code": ">=0.6.0"
  },
  "mcpServers": {
    "mailtool": {
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
}
```

### 9.3 CI/CD Updates

**.github/workflows/ci.yml**:
```yaml
name: CI

on:
  push:
    branches: [main]
  pull_request:

jobs:
  test:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v4

      - name: Install uv
        run: powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"

      - name: Set up Python
        run: uv venv

      - name: Install dependencies
        run: uv sync --all-groups

      - name: Run tests
        run: uv run pytest -v

      - name: Run linting
        run: uv run ruff check .

      - name: Test MCP server
        run: uv run --with mcp --with pywin32 python test_mcp_server.py
```

---

## 10. Rollback Plan

### 10.1 Rollback Triggers

- Critical bug in production
- Performance degradation
- Claude Code incompatibility
- Data loss or corruption

### 10.2 Rollback Procedure

1. **Immediate**: Update plugin.json to point to old server
   ```json
   {
     "mcpServers": {
       "mailtool": {
         "command": "uv",
         "args": ["run", "--with", "pywin32", "mcp_server.py"]
       }
     }
   }
   ```

2. **Short-term**: Keep `mcp_server.py` (old version) for 2 weeks

3. **Investigation**: Review logs, reproduce issue, fix

4. **Re-deploy**: Test fix, update plugin.json again

### 10.3 Rollback Testing

```bash
# Test old server still works
uv run --with pywin32 mcp_server.py

# Test with inspector
npx -y @modelcontextprotocol/inspector

# Connect to old server
# Verify all 23 tools work
```

---

## 11. Success Criteria

### 11.1 Functional Requirements

- ✅ All 23 tools migrated with identical behavior
- ✅ All tools return structured output (Pydantic models)
- ✅ All tools tested with Claude Code
- ✅ Resources accessible for read-only data
- ✅ Error handling comprehensive
- ✅ COM threading model preserved

### 11.2 Non-Functional Requirements

- ✅ Code reduction: ~70% less than manual implementation
- ✅ Type safety: 100% of tools use type hints
- ✅ Test coverage: All tools have tests
- ✅ Documentation: Updated for SDK patterns
- ✅ Performance: No degradation vs old implementation
- ✅ Memory: No leaks (COM objects cleaned up)

### 11.3 Validation Checklist

**Week 1**:
- [ ] SDK infrastructure set up
- [ ] Pydantic models defined
- [ ] Lifespan management working
- [ ] 5-10 simple tools migrated

**Week 2**:
- [ ] All email tools migrated
- [ ] Email resources working
- [ ] Email tests passing
- [ ] Claude Code integration validated

**Week 3**:
- [ ] All calendar tools migrated
- [ ] All task tools migrated
- [ ] All resources working
- [ ] Full test suite passing

**Week 4**:
- [ ] Performance optimized
- [ ] Error handling refined
- [ ] Documentation complete
- [ ] Production deployment successful

---

## 12. Appendix

### 12.1 Common Patterns

**Pattern 1: Simple Get Operation**
```python
@mcp.tool()
def get_thing(entry_id: str, ctx: Context) -> ThingDetails:
    bridge = ctx.request_context.lifespan_context.bridge
    result = bridge.get_thing(entry_id)
    if not result:
        raise McpError(f"Thing not found: {entry_id}")
    return ThingDetails(**result)
```

**Pattern 2: List Operation**
```python
@mcp.tool()
def list_things(limit: int = 10, ctx: Context) -> list[ThingSummary]:
    bridge = ctx.request_context.lifespan_context.bridge
    things = bridge.list_things(limit=limit)
    return [ThingSummary(**t) for t in things]
```

**Pattern 3: Boolean Operation**
```python
@mcp.tool()
def do_thing(entry_id: str, ctx: Context) -> OperationResult:
    bridge = ctx.request_context.lifespan_context.bridge
    result = bridge.do_thing(entry_id)
    if result:
        return OperationResult(success=True, message="Done")
    else:
        raise McpError(f"Failed to do thing: {entry_id}")
```

**Pattern 4: Create Operation**
```python
@mcp.tool()
def create_thing(params: CreateParams, ctx: Context) -> CreateResult:
    bridge = ctx.request_context.lifespan_context.bridge
    entry_id = bridge.create_thing(**params.dict())
    if entry_id:
        return CreateResult(success=True, entry_id=entry_id)
    else:
        return CreateResult(success=False, entry_id=None, message="Failed")
```

### 12.2 Error Handling Best Practices

```python
# DO: Use custom exceptions
raise OutlookNotFoundError(f"Email not found: {entry_id}")

# DON'T: Return generic errors
return {"error": "Not found"}

# DO: Validate inputs early
if not entry_id:
    raise OutlookValidationError("entry_id is required")

# DO: Log errors for debugging
import logging
logger = logging.getLogger(__name__)
try:
    result = bridge.some_operation()
except Exception as e:
    logger.exception("Failed to some_operation")
    raise OutlookComError(f"Operation failed: {e}")
```

### 12.3 COM Threading Best Practices

```python
# DO: Keep all COM calls on same thread
@mcp.tool()
async def my_tool(entry_id: str, ctx: Context) -> Result:
    # This is safe - single thread
    bridge = ctx.request_context.lifespan_context.bridge
    return bridge.get_thing(entry_id)

# DON'T: Use thread pools
@mcp.tool()
def my_tool_bad(entry_id: str) -> Result:
    from concurrent.futures import ThreadPoolExecutor
    # This breaks COM apartment threading!
    with ThreadPoolExecutor() as pool:
        return pool.submit(bridge.get_thing, entry_id).result()
```

### 12.4 Migration Checklist

**Pre-Migration**:
- [ ] Read this entire plan
- [ ] Review MCP SDK v2 documentation
- [ ] Set up development environment
- [ ] Create feature branch
- [ ] Back up current implementation

**Migration**:
- [ ] Complete Phase 1: Foundation
- [ ] Complete Phase 2: Pydantic models
- [ ] Complete Phase 3: Simple tools
- [ ] Complete Phase 4: Email tools + resources
- [ ] Complete Phase 5: Calendar + tasks
- [ ] Complete Phase 6: Polish + deploy

**Post-Migration**:
- [ ] All tests passing
- [ ] Claude Code integration validated
- [ ] Documentation updated
- [ ] Old implementation archived
- [ ] Feature merged to main

---

## 13. Next Steps

1. **Review this plan** with stakeholders
2. **Create feature branch**: `git checkout -b feature/mcp-sdk-v2-migration`
3. **Start Phase 1**: Set up SDK infrastructure
4. **Track progress**: Update this document with completion dates
5. **Ask questions**: Use Claude Code for help during implementation

---

**Document Version**: 1.0
**Last Updated**: 2025-01-18
**Author**: Claude (with human review)
**Status**: Ready for Implementation
