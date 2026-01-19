# Manual MCP Server Testing Guide

This document provides comprehensive instructions for manually testing the MCP SDK v2 server implementation with Claude Code.

## Overview

The manual testing process validates that:
1. All 23 MCP tools are registered and callable
2. All 7 MCP resources are accessible
3. Structured output (Pydantic models) is returned correctly
4. Error handling works as expected
5. Lifespan management is properly configured

## Prerequisites

- Windows with Outlook running
- MCP server package installed: `uv sync --all-groups`
- Claude Code with MCP plugin loaded

## Automated Test Suite

Run the automated test suite:

```bash
uv run --with mcp python scripts/manual_mcp_test.py
```

This validates:
- Tool registration (23 tools)
- Resource registration (7 resources)
- Pydantic model validation (6 models)
- Tool signatures (5 key tools)
- Custom exception classes (3 exceptions)
- Lifespan management configuration

## Manual Claude Code Testing

### Step 1: Start Claude Code

1. Open Claude Code on Windows
2. Verify Outlook is running
3. Verify MCP plugin is loaded: `.claude-plugins/mailtool/.claude-plugin/plugin.json`

### Step 2: Test Email Tools (9 tools)

#### 2.1 List Emails
```
You: List the last 5 emails from my Inbox
```
Expected: Returns list of EmailSummary with 7 fields each (entry_id, subject, sender, sender_name, received_time, unread, has_attachments)

#### 2.2 Get Email Details
```
You: Get the full details of the email with entry_id [ENTRY_ID_FROM_PREVIOUS]
```
Expected: Returns EmailDetails with body and html_body fields

#### 2.3 Send Email
```
You: Send an email to test@example.com with subject "Test Email" and body "This is a test"
```
Expected: Returns SendEmailResult with success=True, entry_id=None (sent email)

#### 2.4 Save Draft
```
You: Create a draft email to test@example.com with subject "Draft Test"
```
Expected: Returns SendEmailResult with success=True, entry_id=<draft_entry_id>

#### 2.5 Reply to Email
```
You: Reply to the email with entry_id [ENTRY_ID] saying "Thanks for the message"
```
Expected: Returns OperationResult with success=True, message="Email replied successfully"

#### 2.6 Forward Email
```
You: Forward the email with entry_id [ENTRY_ID] to forward@example.com
```
Expected: Returns OperationResult with success=True

#### 2.7 Mark Email
```
You: Mark the email with entry_id [ENTRY_ID] as unread
```
Expected: Returns OperationResult with success=True, message="Email marked as unread"

#### 2.8 Move Email
```
You: Move the email with entry_id [ENTRY_ID] to Archive folder
```
Expected: Returns OperationResult with success=True, message="Email moved to Archive"

#### 2.9 Search Emails
```
You: Search for emails with subject containing "project"
```
Expected: Returns list of EmailSummary matching the filter

#### 2.10 Delete Email
```
You: Delete the email with entry_id [ENTRY_ID]
```
Expected: Returns OperationResult with success=True, message="Email deleted successfully"

### Step 3: Test Calendar Tools (7 tools)

#### 3.1 List Calendar Events
```
You: List my calendar events for the next 7 days
```
Expected: Returns list of AppointmentSummary with 12 fields each

#### 3.2 Get Appointment Details
```
You: Get the full details of the appointment with entry_id [ENTRY_ID]
```
Expected: Returns AppointmentDetails with body field

#### 3.3 Create Appointment
```
You: Create an appointment titled "Test Meeting" tomorrow at 2pm to 3pm in Room 101
```
Expected: Returns CreateAppointmentResult with success=True, entry_id=<appointment_entry_id>

#### 3.4 Edit Appointment
```
You: Edit the appointment with entry_id [ENTRY_ID] to change the location to "Room 202"
```
Expected: Returns OperationResult with success=True

#### 3.5 Respond to Meeting
```
You: Accept the meeting invitation with entry_id [ENTRY_ID]
```
Expected: Returns OperationResult with success=True, message="Meeting accepted successfully"

#### 3.6 Get Free/Busy
```
You: Check my free/busy status for tomorrow
```
Expected: Returns FreeBusyInfo with email, dates, free_busy string, resolved=True

#### 3.7 Delete Appointment
```
You: Delete the appointment with entry_id [ENTRY_ID]
```
Expected: Returns OperationResult with success=True

### Step 4: Test Task Tools (7 tools)

#### 4.1 List Tasks
```
You: List my active tasks
```
Expected: Returns list of TaskSummary with 8 fields each (incomplete tasks only)

#### 4.2 List All Tasks
```
You: List all my tasks including completed ones
```
Expected: Returns list of TaskSummary (all tasks)

#### 4.3 Get Task Details
```
You: Get the full details of the task with entry_id [ENTRY_ID]
```
Expected: Returns TaskSummary with body field

#### 4.4 Create Task
```
You: Create a task titled "Test Task" due next Friday with high priority
```
Expected: Returns CreateTaskResult with success=True, entry_id=<task_entry_id>

#### 4.5 Edit Task
```
You: Edit the task with entry_id [ENTRY_ID] to change the due date to tomorrow
```
Expected: Returns OperationResult with success=True

#### 4.6 Complete Task
```
You: Mark the task with entry_id [ENTRY_ID] as complete
```
Expected: Returns OperationResult with success=True, message="Task completed successfully"

#### 4.7 Delete Task
```
You: Delete the task with entry_id [ENTRY_ID]
```
Expected: Returns OperationResult with success=True

### Step 5: Test Resources (7 resources)

#### 5.1 Email Resources
```
You: Show me my recent emails using the inbox://emails resource
You: Show me my unread emails using the inbox://unread resource
You: Get the full email with entry_id [ENTRY_ID] using the email://[ENTRY_ID] resource
```
Expected: Returns formatted text summaries of emails

#### 5.2 Calendar Resources
```
You: Show me today's calendar using the calendar://today resource
You: Show me this week's calendar using the calendar://week resource
```
Expected: Returns formatted text summaries of calendar events

#### 5.3 Task Resources
```
You: Show me my active tasks using the tasks://active resource
You: Show me all my tasks using the tasks://all resource
```
Expected: Returns formatted text summaries of tasks

### Step 6: Test Error Handling

#### 6.1 Not Found Errors
```
You: Get the email with entry_id "invalid_entry_id_12345"
```
Expected: Raises OutlookNotFoundError with error code -32602

#### 6.2 Validation Errors
```
You: Create an appointment with invalid date format
```
Expected: Raises OutlookValidationError with error code -32604

#### 6.3 COM Errors
```
You: (Close Outlook) List my emails
```
Expected: Raises OutlookComError with error code -32603

## Success Criteria

All tests should pass with the following results:

### Tool Registration
- 23 tools registered
- All tools have descriptions
- All tools have proper signatures

### Resource Registration
- 7 resources registered (6 regular + 1 template)
- All resources have valid URIs
- Template resources have parameter matching

### Model Validation
- All Pydantic models validate correctly
- Models serialize to JSON properly
- Field types match expectations

### Tool Signatures
- Key tools (list_emails, get_email, send_email, create_appointment, create_task) have correct signatures
- Parameters are properly typed
- Return types are correctly annotated

### Exception Classes
- OutlookNotFoundError works correctly
- OutlookComError works correctly
- OutlookValidationError works correctly
- All exceptions have proper error codes

### Lifespan Management
- Lifespan is configured on FastMCP server
- Bridge state function exists
- Async context manager is properly set up

## Troubleshooting

### Issue: Tools not found
**Solution**: Verify plugin.json points to correct entry point: `mailtool.mcp.server`

### Issue: Bridge not initialized
**Solution**: Ensure Outlook is running on Windows before starting Claude Code

### Issue: Structured output not working
**Solution**: Verify Pydantic models are correctly defined and imported

### Issue: Resources not accessible
**Solution**: Check that resources are registered in server.py

### Issue: Exception handling not working
**Solution**: Verify custom exception classes are imported and used correctly

## Test Results Template

Use this template to record your test results:

```
## Test Results - [Date]

### Email Tools (9/9 passed)
- [ ] list_emails
- [ ] get_email
- [ ] send_email
- [ ] reply_email
- [ ] forward_email
- [ ] mark_email
- [ ] move_email
- [ ] delete_email
- [ ] search_emails

### Calendar Tools (7/7 passed)
- [ ] list_calendar_events
- [ ] create_appointment
- [ ] get_appointment
- [ ] edit_appointment
- [ ] respond_to_meeting
- [ ] delete_appointment
- [ ] get_free_busy

### Task Tools (7/7 passed)
- [ ] list_tasks
- [ ] list_all_tasks
- [ ] create_task
- [ ] get_task
- [ ] edit_task
- [ ] complete_task
- [ ] delete_task

### Resources (7/7 passed)
- [ ] inbox://emails
- [ ] inbox://unread
- [ ] email://{entry_id}
- [ ] calendar://today
- [ ] calendar://week
- [ ] tasks://active
- [ ] tasks://all

### Error Handling (3/3 passed)
- [ ] OutlookNotFoundError
- [ ] OutlookComError
- [ ] OutlookValidationError

### Automated Test Suite
- [ ] All 6 test suites passed

**Overall Result: [PASS/FAIL]**

**Notes:**
[Record any issues or observations]
```

## Next Steps

After successful manual testing:
1. Update PRD to mark US-046 as passing
2. Proceed to US-048 (Production Deployment)
3. Monitor for issues in production
4. Keep rollback plan ready (docs/ROLLBACK_PLAN.md)
