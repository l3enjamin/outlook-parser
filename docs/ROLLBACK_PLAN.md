# Rollback Plan: MCP SDK v2 Migration (v2.3.0)

**Version**: 2.3.0
**Date**: 2025-01-19
**Purpose**: Emergency rollback procedure for MCP SDK v2 migration issues

## Overview

This document describes the rollback procedure from MCP SDK v2 (FastMCP framework) back to the hand-rolled MCP implementation (v2.2.0) in case of production issues.

**Current State (v2.3.0)**:
- MCP Server: `src/mailtool/mcp/server.py` (FastMCP framework)
- Plugin Entry Point: `mailtool.mcp.server` module
- Architecture: Official MCP Python SDK v2 with FastMCP
- Features: 23 tools, 7 resources, structured output, custom exceptions

**Rollback State (v2.2.0)**:
- MCP Server: `mcp_server.py` (legacy implementation, 855 lines)
- Plugin Entry Point: Script path execution
- Architecture: Hand-rolled MCP implementation
- Features: 23 tools, no resources, basic error handling

## Rollback Triggers

Consider rollback if ANY of the following occur:

1. **Critical Failures**:
   - Claude Code cannot connect to MCP server
   - All tools fail with McpError exceptions
   - Structured output causes data corruption or loss

2. **Performance Issues**:
   - Tool execution time >2x slower than v2.2.0 baseline
   - Memory leaks causing crashes after repeated operations
   - COM objects not released properly causing Outlook instability

3. **Compatibility Issues**:
   - FastMCP framework incompatibilities with Claude Code
   - Pydantic model validation blocking legitimate operations
   - Resource URIs not resolving correctly

4. **Data Loss Risks**:
   - Email/calendar/task operations fail silently
   - EntryID mapping issues causing wrong items to be modified
   - Draft emails lost due to return type changes

## Rollback Procedure

### Phase 1: Immediate Rollback (Emergency)

**Time Estimate**: 2 minutes
**Impact**: All MCP tools unavailable during rollback

1. **Stop Claude Code**:
   ```bash
   # Close Claude Code completely to release MCP connections
   ```

2. **Revert plugin.json**:
   ```bash
   cd C:\dev\mailtool
   git checkout main -- .claude-plugin/plugin.json
   # OR manually edit to use v2.2.0 settings:
   # - Change version from "2.3.0" to "2.2.0"
   # - Change command args from "-m mailtool.mcp.server" to "${CLAUDE_PLUGIN_ROOT}/mcp_server.py"
   # - Change description to remove "MCP SDK v2"
   ```

3. **Verify old server exists**:
   ```bash
   # Check that legacy mcp_server.py exists
   ls mcp_server.py
   # Should show file exists (277 lines, hand-rolled implementation)
   ```

4. **Restart Claude Code**:
   ```bash
   # Launch Claude Code - it will auto-load the reverted plugin.json
   ```

5. **Test basic operations**:
   - Try `list_emails` tool
   - Try `get_email` tool
   - Verify no McpError exceptions

**Success Criteria**: Basic email operations work, no structured output errors

### Phase 2: Code Rollback (If Needed)

**Time Estimate**: 5 minutes
**Impact**: All development work on v2.3.0 paused

1. **Create rollback branch**:
   ```bash
   git checkout -b rollback/mcp-sdk-v2-migration
   git reset --hard v2.2.0
   ```

2. **Revert dependencies**:
   ```bash
   # Edit pyproject.toml to remove mcp dependency
   uv remove mcp
   uv sync --all-groups
   ```

3. **Delete new MCP package**:
   ```bash
   # Remove new SDK v2 implementation
   rm -rf src/mailtool/mcp/
   rm -rf tests/mcp/
   ```

4. **Restore legacy files**:
   ```bash
   # Verify legacy files are in place
   ls mcp_server.py test_mcp_server.py
   ```

5. **Run tests**:
   ```bash
   ./run_tests.sh
   # All bridge tests should pass
   # Legacy MCP tests (test_mcp_server.py) should pass
   ```

**Success Criteria**: All tests pass, plugin loads successfully

### Phase 3: Full Reversion (Last Resort)

**Time Estimate**: 15 minutes
**Impact**: Lose all v2.3.0 development work

1. **Tag current state**:
   ```bash
   git tag -a rollback-mcp-sdk-v2-migration -m "Rollback point for MCP SDK v2 migration (v2.3.0)"
   # Push tag if remote exists
   git push origin rollback-mcp-sdk-v2-migration || echo "No remote configured"
   ```

2. **Find v2.2.0 commit**:
   ```bash
   git log --oneline --grep="2.2.0"
   # Or find commit from release notes
   git checkout <commit-hash>
   git checkout -b production/rollback-to-v2.2.0
   ```

3. **Update plugin.json**:
   ```bash
   # Ensure plugin.json points to v2.2.0 implementation
   # Version: 2.2.0
   # Command: ${CLAUDE_PLUGIN_ROOT}/mcp_server.py
   ```

4. **Reinstall dependencies**:
   ```bash
   uv sync --all-groups
   ```

5. **Validate installation**:
   ```bash
   ./run_tests.sh
   python test_mcp_server.py
   ```

6. **Deploy to production**:
   ```bash
   git push origin production/rollback-to-v2.2.0
   # Users pull latest code and restart Claude Code
   ```

**Success Criteria**: Production system fully reverted to v2.2.0

## Verification Steps

After rollback, verify the following:

### Basic Functionality
- [ ] Claude Code connects to MCP server successfully
- [ ] `list_emails` tool returns email list (not structured output)
- [ ] `get_email` tool returns email details (dict, not EmailDetails)
- [ ] `send_email` tool sends emails correctly
- [ ] No McpError exceptions in logs

### Performance Baseline
- [ ] `list_emails(limit=10)` completes in <2 seconds
- [ ] `get_email(entry_id)` completes in <1 second
- [ ] Memory usage stable after 20 operations
- [ ] No COM reference leaks

### Compatibility
- [ ] All 23 tools accessible in Claude Code
- [ ] Tool signatures match v2.2.0 documentation
- [ ] No Pydantic validation errors
- [ ] Resources not accessible (expected in v2.2.0)

## Rollback Decision Tree

```
Issue Detected
    |
    v
Can it be fixed with a hotfix?
    | Yes -> Implement hotfix, test, deploy
    | No -> Continue
    v
Is it a critical blocker?
    | No -> Document issue, continue using v2.3.0
    | Yes -> Continue
    v
Phase 1 Rollback (plugin.json only)
    |
    v
Fixed?
    | Yes -> Document issue, investigate root cause
    | No -> Continue
    v
Phase 2 Rollback (code rollback)
    |
    v
Fixed?
    | Yes -> Plan re-migration with fixes
    | No -> Continue
    v
Phase 3 Rollback (full reversion)
    |
    v
Fixed? -> Yes -> Analyze failure, plan next steps
```

## Known Issues & Mitigations

### Issue 1: Structured Output Changes
**Symptom**: Tools return Pydantic models instead of dicts
**Impact**: Scripts expecting dict output may fail
**Mitigation**: Update scripts to handle Pydantic models OR rollback to v2.2.0

### Issue 2: Resource URIs Not Found
**Symptom**: `calendar://today` returns 404 error
**Impact**: Resource-based queries fail
**Mitigation**: Use tool equivalents (list_calendar_events) OR rollback

### Issue 3: Custom Exception Types
**Symptom**: OutlookNotFoundError not caught by generic exception handlers
**Impact**: Error handling may fail
**Mitigation**: Update exception handlers to catch McpError OR rollback

### Issue 4: Performance Degradation
**Symptom**: Tools slower than v2.2.0 baseline
**Impact**: User experience degraded
**Mitigation**: Profile and optimize OR rollback to v2.2.0

## Communication Plan

### Pre-Rollback
1. Document issue in GitHub Issues
2. Tag issue as "critical" and "rollback"
3. Notify users via README banner

### During Rollback
1. Update README with rollback notice
2. Post rollback status in issue thread
3. Monitor for additional issues

### Post-Rollback
1. Document root cause analysis
2. Create fix plan for re-migration
3. Schedule re-migration with proper testing

## Re-Migration Plan

After rollback and root cause fix:

1. **Fix root cause** in feature branch
2. **Add regression tests** for the fix
3. **Run full test suite** including manual Claude Code testing
4. **Create beta release** for limited testing
5. **Monitor beta** for 1 week
6. **Production deployment** with monitoring

## Retention Policy

**Keep v2.2.0 Available**: 2 weeks after successful v2.3.0 deployment

- **Week 1**: Monitor for critical issues
- **Week 2**: Monitor for edge cases and performance issues
- **After Week 2**: Archive old implementation (see US-049)

**Legacy Files Retention**:
- `mcp_server.py` - Keep until US-049 (archive old implementation)
- `test_mcp_server.py` - Keep until US-049
- Git commit history - Keep indefinitely for reference
- Rollback tag (if created) - Keep indefinitely for reference

## Emergency Contacts

- **Developer**: Sam <s.mok@utwente.nl>
- **GitHub Issues**: https://github.com/sammok/mailtool/issues
- **Documentation**: See CLAUDE.md, README.md, MCP_INTEGRATION.md

## Test Plan

Before rollback, test the following:

1. **Legacy server starts**:
   ```bash
   uv run --with pywin32 python mcp_server.py
   # Should start without errors (no version banner in legacy code)
   # Press Ctrl+C to stop
   ```

2. **Legacy tools work**:
   ```bash
   python test_mcp_server.py
   # All tests should pass
   ```

3. **Plugin loads**:
   - Update plugin.json to v2.2.0 settings
   - Restart Claude Code
   - Verify tools appear in tool list

4. **Basic operations**:
   - `list_emails(limit=5)` returns 5 emails
   - `get_email(entry_id)` returns email details
   - `send_email(to="test@example.com", subject="Test")` sends email

## Success Metrics

Rollback successful if:
- [ ] All 23 tools accessible and functional
- [ ] No McpError or structured output errors
- [ ] Performance within 10% of v2.2.0 baseline
- [ ] No data loss or corruption
- [ ] Claude Code stable with no crashes

## Appendix: Quick Reference

### plugin.json v2.3.0 (Current)
```json
{
  "version": "2.3.0",
  "mcpServers": {
    "mailtool": {
      "command": "uv",
      "args": ["run", "--with", "pywin32", "-m", "mailtool.mcp.server"]
    }
  }
}
```

### plugin.json v2.2.0 (Rollback)
```json
{
  "version": "2.2.0",
  "mcpServers": {
    "mailtool": {
      "command": "uv",
      "args": ["run", "--with", "pywin32", "${CLAUDE_PLUGIN_ROOT}/mcp_server.py"]
    }
  }
}
```

### Key Differences
- **v2.3.0**: Module invocation (`-m mailtool.mcp.server`)
- **v2.2.0**: Script path (`${CLAUDE_PLUGIN_ROOT}/mcp_server.py`)
- **v2.3.0**: Structured output (Pydantic models)
- **v2.2.0**: Dict output (plain Python dicts)
- **v2.3.0**: Resources available (7 resources)
- **v2.2.0**: No resources (tools only)

---

**Last Updated**: 2025-01-19
**Status**: Ready for emergency rollback
**Next Review**: After US-049 (archive old implementation)
