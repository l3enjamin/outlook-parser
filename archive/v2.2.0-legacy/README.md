# Mailtool v2.2.0 Legacy MCP Implementation

This directory contains the legacy MCP server implementation from mailtool v2.2.0, archived after successful migration to MCP SDK v2 in v2.3.0.

## Migration Timeline

- **2025-01-19**: v2.3.0 released with MCP SDK v2 migration
- **2025-02-02**: 2-week production rollback window ended successfully
- **2025-02-02**: Legacy files archived to this directory

## Archived Files

### `mcp_server.py`
- Hand-rolled MCP server implementation using stdio transport
- 23 tools for Outlook automation (email, calendar, tasks)
- Custom JSON-RPC request/response handling
- Replaced by: `src/mailtool/mcp/server.py` (FastMCP framework)

### `test_mcp_server.py`
- Manual test script for legacy MCP server
- JSON-RPC request validation via subprocess
- Replaced by: `tests/mcp/` test suite (166 automated tests)

## Why These Were Archived

These files were kept for 2 weeks after the v2.3.0 production deployment to allow for immediate rollback if issues were discovered. After 2 weeks of successful production operation with no rollbacks required, the files were archived for historical reference.

## Do Not Use

These files are **NOT** in active use and should not be referenced or executed. The current MCP server implementation is in:

- **Server**: `src/mailtool/mcp/server.py`
- **Tests**: `tests/mcp/`
- **Models**: `src/mailtool/mcp/models.py`
- **Resources**: `src/mailtool/mcp/resources.py`
- **Lifespan**: `src/mailtool/mcp/lifespan.py`
- **Exceptions**: `src/mailtool/mcp/exceptions.py`

## Migration Details

For information about the MCP SDK v2 migration, see:

- `docs/MCP_INTEGRATION.md` - Current MCP server documentation
- `docs/ROLLBACK_PLAN.md` - Rollback procedure (archived for reference)
- `docs/DEPLOYMENT_GUIDE.md` - Production deployment guide
- `scripts/ralph/progress.txt` - Complete migration history

## Key Improvements in v2.3.0

1. **Official MCP SDK**: Using `mcp>=0.9.0` with FastMCP framework
2. **Structured Output**: All tools return Pydantic models (10 models)
3. **MCP Resources**: 7 resources for read-only data access
4. **Type Safety**: Full type hints with Pydantic validation
5. **Error Handling**: 3 custom exception classes
6. **Async Lifespan**: Proper async context manager for COM bridge lifecycle
7. **Comprehensive Testing**: 166 automated tests (vs 1 manual test script)
8. **Logging**: Comprehensive logging for debugging and monitoring
9. **Documentation**: Extensive documentation for users and developers

## Contact

For questions about the legacy implementation or migration details, refer to the project documentation or git history.
