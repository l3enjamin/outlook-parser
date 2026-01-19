# Production Deployment Guide: MCP SDK v2 (v2.3.0)

**Version**: 2.3.0
**Date**: 2025-01-19
**Status**: Ready for Production Deployment

## Overview

This document provides the production deployment procedure for MCP SDK v2 migration (v2.3.0) to Claude Code.

**Pre-Deployment Checklist**:
- ✅ All 166 MCP tests passing
- ✅ All 50 bridge tests passing
- ✅ Automated test suite passing (6 test suites)
- ✅ Performance benchmarks created
- ✅ Memory leak tests implemented
- ✅ Rollback plan documented (ROLLBACK_PLAN.md)
- ✅ Manual testing guide created (MANUAL_TESTING_GUIDE.md)
- ✅ CI/CD workflow updated
- ✅ Plugin configuration updated (v2.3.0)

## Deployment Procedure

### Phase 1: Pre-Deployment Validation (15 min)

#### 1.1 Run Automated Test Suite
```bash
# Run all MCP tests
uv run --with pytest --with pywin32 python -m pytest tests/mcp/ -v

# Expected: 166 passed
```

#### 1.2 Run Automated Server Validation
```bash
# Run automated server structure tests
uv run --with mcp python scripts/manual_mcp_test.py

# Expected: 6 test suites passing
```

#### 1.3 Verify Plugin Configuration
```bash
# Check plugin.json points to SDK v2 server
cat .claude-plugin/plugin.json

# Expected:
# - version: "2.3.0"
# - command: "uv"
# - args: ["run", "--with", "pywin32", "-m", "mailtool.mcp.server"]
```

#### 1.4 Verify Rollback Files Present
```bash
# Check legacy files exist for rollback
ls -lh mcp_server.py test_mcp_server.py

# Expected: Both files exist (not deleted)
```

### Phase 2: Deployment (5 min)

#### 2.1 Commit Any Outstanding Changes
```bash
# Check git status
git status

# Commit any changes (if needed)
git add .
git commit -m "feat(mcp): US-048 - Production deployment preparation"
```

#### 2.2 Merge to Main Branch
```bash
# Switch to main branch
git checkout main

# Pull latest changes
git pull origin main

# Merge feature branch
git merge feature/mcp-sdk-v2-migration

# Push to remote
git push origin main
```

#### 2.3 Tag Release (Optional)
```bash
# Create git tag for v2.3.0
git tag -a v2.3.0 -m "Release v2.3.0: MCP SDK v2 Migration"

# Push tag to remote
git push origin v2.3.0
```

### Phase 3: Production Activation (2 min)

#### 3.1 Restart Claude Code
1. Close all Claude Code instances
2. Reopen Claude Code
3. Verify MCP plugin loads successfully
4. Check Claude Code logs for errors

#### 3.2 Verify Server Startup
1. Open Claude Code on Windows
2. Verify Outlook is running
3. Check MCP server connection in Claude Code UI
4. Expected: Server connects without errors

### Phase 4: Smoke Testing (10 min)

#### 4.1 Test Email Tools
```
You: List the last 5 emails from my Inbox
```
Expected: Returns list of EmailSummary with 7 fields

```
You: Get the details of the first email
```
Expected: Returns EmailDetails with body content

#### 4.2 Test Calendar Tools
```
You: Show me my calendar for today
```
Expected: Returns list of AppointmentSummary for today

#### 4.3 Test Task Tools
```
You: List my active tasks
```
Expected: Returns list of TaskSummary (incomplete tasks)

#### 4.4 Test Resources
```
You: Read my inbox://emails resource
```
Expected: Returns formatted text of recent emails

#### 4.5 Test Error Handling
```
You: Get email with ID fake-entry-id
```
Expected: Returns OutlookNotFoundError with entry_id attribute

### Phase 5: Monitoring (First 24 hours)

#### 5.1 Monitor Performance Metrics
- Tool execution time (should be <2x v2.2.0 baseline)
- Memory usage (should not grow continuously)
- COM object cleanup (Outlook should remain stable)

#### 5.2 Monitor Error Rates
- McpError exceptions (should be <5% of operations)
- OutlookNotFoundError (expected for invalid entry_ids)
- OutlookComError (should be rare, <1% of operations)

#### 5.3 Monitor User Feedback
- Claude Code interaction logs
- Error reports from users
- Performance complaints

### Phase 6: Rollback Decision (First 24 hours)

#### Rollback Triggers
If ANY of the following occur, initiate rollback (see ROLLBACK_PLAN.md):

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

#### Rollback Procedure
See ROLLBACK_PLAN.md for detailed rollback instructions:
- Phase 1: Immediate rollback (2 min) - revert plugin.json only
- Phase 2: Code rollback (5 min) - remove MCP SDK v2 code
- Phase 3: Full reversion (15 min) - complete rollback to v2.2.0

## Post-Deployment Actions

### After 24 Hours (If No Rollback)

1. **Update Documentation**:
   - Mark v2.3.0 as current stable version in README.md
   - Update CHANGELOG.md with release notes

2. **Archive Legacy Files** (Do NOT delete yet):
   - Keep mcp_server.py for 2 weeks (until 2025-02-02)
   - Keep test_mcp_server.py for 2 weeks
   - Update docs/ROLLBACK_PLAN.md with deployment date

3. **Monitor Continuously**:
   - Check Claude Code logs daily for first week
   - Review error rates weekly
   - Gather user feedback

### After 2 Weeks (If Successful)

1. **Archive Old Implementation** (US-049):
   - Move mcp_server.py to archive/
   - Move test_mcp_server.py to archive/
   - Update ROLLBACK_PLAN.md to note end of rollback window

2. **Final Validation** (US-050):
   - All tests passing
   - All tools working
   - Documentation complete
   - No known issues
   - Ready for release

## Success Criteria

Deployment is successful if:
- ✅ All smoke tests pass (email, calendar, task, resources)
- ✅ No critical failures in first 24 hours
- ✅ Performance degradation <20% compared to v2.2.0
- ✅ Error rate <5% of operations
- ✅ No data loss or corruption reported
- ✅ User feedback is positive

## Emergency Contacts

- **Developer**: Sam <s.mortazavi@utwente.nl>
- **Rollback Plan**: docs/ROLLBACK_PLAN.md
- **Manual Testing**: docs/MANUAL_TESTING_GUIDE.md

## Deployment Timeline

- **Pre-Deployment Validation**: 15 min
- **Deployment (merge to main)**: 5 min
- **Production Activation**: 2 min
- **Smoke Testing**: 10 min
- **Total Time to Deploy**: 32 min
- **Monitoring Period**: 24 hours
- **Rollback Window**: 2 weeks (until 2025-02-02)

## Notes

- Legacy files (mcp_server.py, test_mcp_server.py) must be kept for 2 weeks after deployment
- Rollback plan is tested and documented in docs/ROLLBACK_PLAN.md
- Manual testing guide available in docs/MANUAL_TESTING_GUIDE.md
- All tests pass (216 total: 50 bridge + 166 MCP)
- Plugin configuration already updated to v2.3.0
- No breaking changes to tool signatures (all 23 tools compatible)
- New features in v2.3.0: 7 resources, structured output, custom exceptions

## Next Steps

1. Complete smoke testing in Phase 4
2. Monitor for 24 hours (Phase 5)
3. If successful, proceed with post-deployment actions
4. If issues occur, execute rollback plan
5. After 2 weeks, archive old implementation (US-049)
6. Final validation and sign-off (US-050)
