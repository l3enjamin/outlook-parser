# GEMINI.md - Mailtool: Outlook Automation Bridge

This file is a comprehensive guide for AI agents (like Gemini) to understand the logic, architecture, and design patterns of the `outlook-parser` codebase. Read this carefully to safely navigate and modify the project.

## 🌟 The Core Problem We Solve

Outlook exposes its data via Windows COM. However, Outlook reply chains embed the entire quoted history in every single email. If an AI agent fetches 10 emails from a thread, it reads the exact same content 10 times, wasting context tokens and making it hard to extract the *delta* (what's actually new in each reply).

This project wraps Outlook COM in an **MCP Server (Model Context Protocol)** with built-in **Thread Deduplication**, so AI agents receive only the new text in every message, or can safely request an entire clean thread at once.

## 🧠 Core Architecture & Logic

### 1. O(1) Direct Access (Performance)

Every item in Outlook has a unique string called an `EntryID`. The core logic of the bridge (`src/mailtool/bridge.py`) relies almost entirely on `namespace.GetItemFromID(entry_id)`. We **never** iterate over folders to find specific items if we already have the ID, as COM iteration in Python is extremely slow.

### 2. COM Threading Constraints

Windows COM uses the Single-Threaded Apartment (STA) model. This dictates our architecture:

* COM objects generally cannot be passed between threads.
* COM calls are inherently synchronous holding the thread execution until resolved.
* **How we handle it**: The MCP Server (`src/mailtool/mcp/server.py`) is fully asynchronous using `FastMCP`. To not block its event loop, it initializes a synchronous Outlook bridge inside a ThreadPoolExecutor (see `lifespan.py`). The async MCP tool handlers dispatch tasks down to the synchronous `bridge.py`.

### 3. Agent-First Email Thread Deduplication

This is the most critical logic for you to understand in `bridge.py`.

* **`get_email_parsed()`**: Fetches a single email and returns an `EmailParsed` Pydantic model. It accepts a `deduplication_tier` parameter.
  * **`tier="low"` (Default & Preferred)**: Extremely fast. It checks for the presence of an `In-Reply-To` SMTP header. If present, it signals a reply. We use `mailparser-reply` to strip the quoted text, returning just the new `latest_reply` fragment alongside an array of all `fragments`. **No folder lookup is needed**.
  * **`tier="medium"`**: Slower. Actually searches the `Inbox` and `Sent Items` folders to verify the parent email truly exists before stripping quoted text.
* **`get_email_thread()`**: The recommended entry point for agents analyzing conversational tickets. It takes a single `entry_id` and uses `item.GetConversation().GetTable()` to efficiently walk the entire chain. It returns an `EmailThread` model containing a chronological list of `EmailParsed` objects, each already deduplicated.

### 4. Avoiding the Calendar Bomb

When fetching calendar events (`list_calendar_events`), Outlook can return infinite items for a recurring meeting without an end date. To avoid crashing:

1. We set `Items.IncludeRecurrences = True`.
2. We MUST sort the items array immediately: `Items.Sort("[Start]")`.
3. **Crucially**: We apply a COM-level string filter: `Items.Restrict("[Start] >= '...' AND [End] <= '...'")` **before** iterating over them in Python.

## 📂 Codebase Mental Map

* **`src/mailtool/bridge.py`**: The brain. Contains the `OutlookBridge` class using `pywin32`. All actual Outlook interaction and data parsing happens here. (~2400 lines)
* **`src/mailtool/mcp/server.py`**: The FastMCP server. Exposes 27 tools (Email, Calendar, Tasks) that AI agents can call. It acts as the routing layer to map tool arguments to `bridge.py` methods.
* **`src/mailtool/mcp/models.py`**: Pydantic models (e.g., `EmailParsed`, `EmailThread`). All MCP tools return strictly typed outputs based on these models.
* **`src/mailtool/mcp/resources.py`**: Read-only MCP resources (like `inbox://emails`).
* **`src/mailtool/cli.py`**: A fallback CLI interface for human usage using `argparse`.

## 🛠️ MCP Tool Overview

The server provides 27 tools. As an agent, you should lean heavily on:

* **`get_email_thread`**: Returns a clean, chronological thread.
* **`search_emails` / `search_emails_by_sender`**: For finding entry points.
* **`send_email` / `reply_email` / `forward_email`**: For taking action.

## 🚨 Development Conventions & Rules

1. **Dependencies**: Use `uv run --with pywin32` or `uv sync` to manage python dependencies. Avoid global python pollution. The tool is run via `uv` scripts.
2. **Safety First**: COM operations can easily throw exceptions for missing properties. We use `_safe_get_attr()` when accessing object properties to handle `com_error` exceptions quietly.
3. **Exceptions**: Use custom exceptions from `exceptions.py` when raising errors to MCP (`OutlookNotFoundError`, `OutlookComError`, `OutlookValidationError`).
4. **Testing**: Prefix test items with `[TEST]` for automatic cleanup. Run with `uv run pytest`. Tests interact with the live Outlook client, there is no mocking.
