# Mailtool: Outlook COM Automation Bridge

A WSL2-to-Windows bridge for Outlook automation via COM, optimized for AI agent integration.

## Architecture

**Stack**: Python + pywin32 (COM) → Outlook (Windows)

**Entry Points**:
- `outlook.sh` (WSL2) → `outlook.bat` (Windows) → `src/mailtool_outlook_bridge.py`
- `run_tests.sh` (WSL2) → `run_tests.bat` (Windows) → `pytest`

**Dependency Management**: Uses `uv run --with pywin32` for zero-install Windows execution

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
├── src/
│   └── mailtool_outlook_bridge.py  # Core COM automation (~1240 lines)
└── tests/
    ├── conftest.py         # Session fixtures, warmup, cleanup
    ├── test_bridge.py      # Core connectivity (6 tests)
    ├── test_emails.py      # Email ops (12 tests)
    ├── test_calendar.py    # Calendar ops (13 tests)
    └── test_tasks.py       # Task ops (13 tests)
```

## API Patterns

### Return Values
- **Draft emails**: Returns `EntryID` (string) for reference
- **Sent emails**: Returns `True` (boolean)
- **Failed ops**: Returns `False` (boolean)
- **Get ops**: Returns `dict` with full item data or `None`

### Test Isolation
All test-created items use `[TEST]` prefix for identification and auto-cleanup. Tests use real Outlook data - no mocking.

## Recent Changes (v2.0 → v2.1)

1. **Calendar Bomb Fix**: Added `items.Restrict()` before iterating in `list_calendar_events()`
2. **WSL Path Translation**: Auto-convert attachment paths in `outlook.sh`
3. **Free/Busy Refactor**: Accepts `email_address` directly, defaults to current user
4. **Return Value Docs**: Clarified different return types in `send_email()` docstring

## Running Tests

```bash
# From WSL2
./run_tests.sh                 # All tests
./run_tests.sh -m email        # Email tests only
./run_tests.sh -m "not slow"   # Skip slow tests

# From Windows
run_tests.bat
```

## Known Limitations

- **Date Format**: Outlook COM filters use locale-specific formats (currently MM/DD/YYYY HH:MM)
- **Parallel Execution**: COM is apartment-threaded; true parallel test execution not recommended
- **Sent Item ID**: Sent emails move to Sent Items with new EntryID (can't return original ID)

## Development Notes

- **COM Threading**: All COM calls must happen on same thread (session-scoped bridge fixture)
- **Warmup**: Tests include 2-5s warmup to ensure Outlook is responsive
- **Cleanup**: Test artifacts auto-cleaned via prefix-based deletion helpers
