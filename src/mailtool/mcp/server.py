"""Mailtool MCP Server

This module provides the main FastMCP server instance for Outlook automation.
It implements the Model Context Protocol (MCP) using the official MCP Python SDK v2
with the FastMCP framework.

The server provides tools and resources for Outlook email, calendar, and task management.
All tools return structured Pydantic models for type safety and LLM understanding.
"""

import argparse
import logging
from typing import TYPE_CHECKING

from mcp.server import FastMCP

from mailtool.mcp.com_state import ensure_com_initialized
from mailtool.mcp.exceptions import OutlookComError, OutlookNotFoundError
from mailtool.mcp.lifespan import outlook_lifespan
from mailtool.mcp.models import (
    AppointmentDetails,
    AppointmentSummary,
    CreateAppointmentResult,
    CreateTaskResult,
    EmailParsed,
    EmailSummary,
    FreeBusyInfo,
    OperationResult,
    SendEmailResult,
    TaskSummary,
)
from mailtool.mcp.resources import (
    register_calendar_resources,
    register_email_resources,
    register_task_resources,
)

if TYPE_CHECKING:
    from mailtool.bridge import OutlookBridge

# Configure logging for the MCP server
# Logs are written to stderr for debugging and monitoring
logger = logging.getLogger(__name__)

# Global variable to store default account from CLI args
# This is set in main() before the server runs, and read by the lifespan
_default_account: str | None = None


# Create FastMCP server instance
# The lifespan parameter manages Outlook COM bridge lifecycle (creation, warmup, cleanup)
# Note: The lifespan reads from the global _default_account variable
mcp = FastMCP(
    name="mailtool-outlook-bridge",
    lifespan=outlook_lifespan,
)

# Register resources
register_email_resources(mcp)
register_calendar_resources(mcp)
register_task_resources(mcp)

# Module-level bridge instance (set by lifespan, accessed by tools)
_bridge: "OutlookBridge | None" = None


def _get_bridge():
    """Get the current bridge instance

    Returns:
        OutlookBridge: The bridge instance

    Raises:
        OutlookComError: If bridge is not initialized (server not running)
    """
    global _bridge

    # Ensure COM is initialized for the current thread before accessing bridge
    ensure_com_initialized()

    if _bridge is None:
        logger.error("Outlook bridge not initialized. Is the server running?")
        raise OutlookComError("Outlook bridge not initialized. Is the server running?")
    logger.debug("Retrieved Outlook bridge instance")
    return _bridge


# ============================================================================
# Email Tools
# ============================================================================


@mcp.tool()
def list_emails(
    limit: int = 10, folder: str = "Inbox", unread_only: bool = False
) -> list[EmailSummary]:
    """
    List emails from the specified folder.

    Retrieves a list of email summaries from the specified folder, sorted by
    received time (most recent first). Uses O(1) direct access for each email.

    Args:
        limit: Maximum number of emails to return (default: 10)
        folder: Folder name to list emails from (default: "Inbox")
        unread_only: If True, only return unread emails (default: False)

    Returns:
        list[EmailSummary]: List of email summaries with basic information

    Raises:
        OutlookComError: If bridge is not initialized or folder cannot be accessed
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    if unread_only and folder == "Inbox":
        # Use optimized search for unread inbox emails
        result = bridge.search_emails(filter_query="[Unread] = TRUE", limit=limit)
    else:
        # For other folders or all emails, use standard list
        # Note: If unread_only is True for non-Inbox, we rely on post-filtering
        # which respects the limit on FETCHED items, not returned items.
        # Ideally search_emails should support folders, but for now this covers the main use case.
        items = bridge.list_emails(limit=limit * 2 if unread_only else limit, folder=folder)
        if unread_only:
            result = [e for e in items if e["unread"]][:limit]
        else:
            result = items

    # Convert bridge result to list of EmailSummary models
    return [
        EmailSummary(
            entry_id=email["entry_id"],
            subject=email["subject"],
            sender=email["sender"],
            sender_name=email["sender_name"],
            received_time=email["received_time"],
            unread=email["unread"],
            has_attachments=email["has_attachments"],
        )
        for email in result
    ]


@mcp.tool()
def get_email(
    entry_id: str,
    remove_quoted: bool = False,
    deduplication_tier: str = "none",
    strip_html: bool = True,
) -> EmailParsed:
    """
    Get full structured email details by entry ID.

    Retrieves complete email information including headers, body parts, recipients,
    and metadata using `mail-parser` logic for rich structure. Uses O(1) direct access.

    Deduplication Tiers:
    - 'none': (Default) No deduplication. Returns full body.
    - 'low': Checks In-Reply-To/References metadata. If the parent email exists in Outlook,
             strips the quoted content from the body. (Matches default behavior of remove_quoted=True)
    - 'medium': If 'low' fails, also checks for matching Subject (ignoring RE:/FW:).
    - 'high': Same as medium currently.

    Args:
        entry_id: Outlook EntryID of the email (O(1) direct access)
        remove_quoted: DEPRECATED - use deduplication_tier='low'. If True, sets tier to 'low'.
        deduplication_tier: Strategy for removing duplicate quoted text ('none', 'low', 'medium', 'high').
        strip_html: If True, remove HTML code from body and clear text_html field (default: True).

    Returns:
        EmailParsed: Richly structured email object

    Raises:
        OutlookNotFoundError: If email not found
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Get parsed email from bridge
    result = bridge.get_email_parsed(
        entry_id,
        remove_quoted=remove_quoted,
        deduplication_tier=deduplication_tier,
        strip_html=strip_html,
    )

    # Check if email was found
    if result is None:
        logger.error(f"Email not found: {entry_id}")
        raise OutlookNotFoundError("Email not found", entry_id=entry_id)

    logger.debug(f"Retrieved parsed email: {entry_id}")

    # Convert bridge result to EmailParsed model
    return EmailParsed(**result)


@mcp.tool()
def mark_email(entry_id: str, unread: bool = False) -> OperationResult:
    """
    Mark an email as read or unread.

    Changes the read/unread status of an email using O(1) direct access via EntryID.

    Args:
        entry_id: Outlook EntryID of the email (O(1) direct access)
        unread: True to mark as unread, False to mark as read (default: False)

    Returns:
        OperationResult: Result of the operation with success status and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Mark email as read/unread via bridge
    result = bridge.mark_email_read(entry_id, unread=unread)

    # Convert boolean result to OperationResult
    if result:
        return OperationResult(
            success=True,
            message=f"Email marked as {'unread' if unread else 'read'}",
        )
    else:
        return OperationResult(
            success=False,
            message=f"Failed to mark email as {'unread' if unread else 'read'}",
        )


@mcp.tool()
def delete_email(entry_id: str) -> OperationResult:
    """
    Delete an email.

    Permanently deletes an email using O(1) direct access via EntryID.

    Args:
        entry_id: Outlook EntryID of the email (O(1) direct access)

    Returns:
        OperationResult: Result of the operation with success status and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Delete email via bridge
    result = bridge.delete_email(entry_id)

    # Convert boolean result to OperationResult
    if result:
        return OperationResult(
            success=True,
            message="Email deleted successfully",
        )
    else:
        return OperationResult(
            success=False,
            message="Failed to delete email",
        )


@mcp.tool()
def send_email(
    to: str,
    subject: str,
    body: str,
    cc: str | None = None,
    bcc: str | None = None,
    html_body: str | None = None,
    file_paths: list[str] | None = None,
    save_draft: bool = False,
) -> SendEmailResult:
    """
    Send an email or save it as a draft.

    Creates and sends an email, or saves it to the Drafts folder.
    Supports attachments, CC/BCC recipients, and HTML body.

    Args:
        to: Primary recipient email address
        subject: Email subject line
        body: Plain text email body
        cc: CC recipients (optional)
        bcc: BCC recipients (optional)
        html_body: HTML email body (optional, overrides plain text body if provided)
        file_paths: List of file paths to attach (optional)
        save_draft: If True, save to Drafts instead of sending (default: False)

    Returns:
        SendEmailResult: Result with success status, draft entry ID (if saved), and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Send email via bridge
    result = bridge.send_email(
        to=to,
        subject=subject,
        body=body,
        cc=cc,
        bcc=bcc,
        html_body=html_body,
        file_paths=file_paths,
        save_draft=save_draft,
    )

    # Convert bridge result to SendEmailResult
    # Bridge returns: False (failed), True (sent), str (draft EntryID)
    if result is False:
        return SendEmailResult(
            success=False,
            entry_id=None,
            message="Failed to send email",
        )
    elif result is True:
        return SendEmailResult(
            success=True,
            entry_id=None,
            message="Email sent successfully",
        )
    else:  # str - draft EntryID
        return SendEmailResult(
            success=True,
            entry_id=result,
            message=f"Email saved as draft: {result}",
        )


@mcp.tool()
def reply_email(entry_id: str, body: str, reply_all: bool = False) -> OperationResult:
    """
    Reply to an email.

    Replies to an email using O(1) direct access via EntryID.
    Can reply to sender only or reply to all recipients.

    Args:
        entry_id: Outlook EntryID of the email (O(1) direct access)
        body: Reply body text
        reply_all: True to reply to all recipients, False to reply to sender only (default: False)

    Returns:
        OperationResult: Result of the operation with success status and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Reply to email via bridge
    result = bridge.reply_email(entry_id, body=body, reply_all=reply_all)

    # Convert boolean result to OperationResult
    if result:
        return OperationResult(
            success=True,
            message=f"Email {'replied to all' if reply_all else 'replied'} successfully",
        )
    else:
        return OperationResult(
            success=False,
            message=f"Failed to {'reply to all' if reply_all else 'reply'}",
        )


@mcp.tool()
def forward_email(entry_id: str, to: str, body: str = "") -> OperationResult:
    """
    Forward an email.

    Forwards an email to a recipient using O(1) direct access via EntryID.
    Optionally adds additional body text.

    Args:
        entry_id: Outlook EntryID of the email (O(1) direct access)
        to: Recipient email address to forward to
        body: Optional additional body text (default: "")

    Returns:
        OperationResult: Result of the operation with success status and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Forward email via bridge
    result = bridge.forward_email(entry_id, to=to, body=body)

    # Convert boolean result to OperationResult
    if result:
        return OperationResult(
            success=True,
            message="Email forwarded successfully",
        )
    else:
        return OperationResult(
            success=False,
            message="Failed to forward email",
        )


@mcp.tool()
def move_email(entry_id: str, folder: str) -> OperationResult:
    """
    Move an email to a different folder.

    Moves an email to a specified folder using O(1) direct access via EntryID.

    Args:
        entry_id: Outlook EntryID of the email (O(1) direct access)
        folder: Target folder name (e.g., "Archive", "Drafts", "Sent Items")

    Returns:
        OperationResult: Result of the operation with success status and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Move email via bridge
    result = bridge.move_email(entry_id, folder_name=folder)

    # Convert boolean result to OperationResult
    if result:
        return OperationResult(
            success=True,
            message=f"Email moved to {folder}",
        )
    else:
        return OperationResult(
            success=False,
            message=f"Failed to move email to {folder}",
        )


@mcp.tool()
def search_emails(filter_query: str, limit: int = 100) -> list[EmailSummary]:
    """
    Search emails using Outlook filter query.

    Searches emails in the Inbox using Outlook Restriction filter (O(1) search).
    Supports SQL-like filter syntax for advanced queries.

    NOTE: For searching by sender email address, especially for internal/Exchange
    users, use search_emails_by_sender() instead. The SenderEmailAddress filter
    does not work for Exchange addresses (internal emails).

    Args:
        filter_query: SQL-like filter query string (e.g., "[Subject] LIKE '%meeting%'")
        limit: Maximum number of results to return (default: 100)

    Returns:
        list[EmailSummary]: List of matching email summaries

    Raises:
        OutlookComError: If bridge is not initialized

    Examples:
        search_emails("[Subject] LIKE '%project%'")  # Search by subject
        search_emails("[Unread] = TRUE")  # Find unread emails
        search_emails("[SenderName] LIKE '%John%'")  # Search by sender name (works better than email)
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Search emails via bridge
    result = bridge.search_emails(filter_query=filter_query, limit=limit)

    # Convert bridge result to list of EmailSummary models
    return [
        EmailSummary(
            entry_id=email["entry_id"],
            subject=email["subject"],
            sender=email["sender"],
            sender_name=email["sender_name"],
            received_time=email["received_time"],
            unread=email["unread"],
            has_attachments=email["has_attachments"],
        )
        for email in result
    ]


@mcp.tool()
def search_emails_by_sender(
    sender_email: str, limit: int = 100, folder: str = "Inbox"
) -> list[EmailSummary]:
    """
    Search emails by sender email address (handles Exchange addresses).

    This method properly handles both SMTP and Exchange email addresses.
    For Exchange users (internal emails), it resolves the Exchange address
    to SMTP address before matching.

    This is the recommended way to search for emails from a specific sender,
    especially for internal/Exchange email addresses where SenderEmailAddress
    filter would not work.

    Args:
        sender_email: Email address to search for (e.g., "f.muijzer@utwente.nl")
        limit: Maximum number of results to return (default: 100)
        folder: Folder name to search in (default: "Inbox")

    Returns:
        list[EmailSummary]: List of matching email summaries

    Raises:
        OutlookComError: If bridge is not initialized

    Examples:
        search_emails_by_sender("john@example.com")  # Search by sender
        search_emails_by_sender("f.muijzer@utwente.nl", limit=50)  # Internal Exchange user
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Search emails by sender via bridge
    result = bridge.search_by_sender(sender_email=sender_email, limit=limit, folder=folder)

    # Convert bridge result to list of EmailSummary models
    return [
        EmailSummary(
            entry_id=email["entry_id"],
            subject=email["subject"],
            sender=email["sender"],
            sender_name=email["sender_name"],
            received_time=email["received_time"],
            unread=email["unread"],
            has_attachments=email["has_attachments"],
        )
        for email in result
    ]


# ============================================================================
# Calendar Tools
# ============================================================================


@mcp.tool()
def list_calendar_events(
    days: int = 7, all_events: bool = False
) -> list[AppointmentSummary]:
    """
    List calendar events for the next N days.

    Retrieves a list of calendar events/appointments from the Outlook calendar.
    Supports date filtering and includes recurring meetings.
    Uses O(1) direct access and COM-level filtering for performance.

    Args:
        days: Number of days ahead to look (default: 7)
        all_events: If True, return all events without date filtering (default: False)

    Returns:
        list[AppointmentSummary]: List of appointment summaries with basic information

    Raises:
        OutlookComError: If bridge is not initialized

    Note:
        This function handles the "Calendar Bomb" issue by applying COM-level
        Restrict filters before Python iteration to prevent infinite recurring meetings.
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # List calendar events via bridge
    result = bridge.list_calendar_events(days=days, all_events=all_events)

    # Convert bridge result to list of AppointmentSummary models
    return [
        AppointmentSummary(
            entry_id=event["entry_id"],
            subject=event["subject"],
            start=event["start"],
            end=event["end"],
            location=event["location"],
            organizer=event["organizer"],
            all_day=event["all_day"],
            required_attendees=event["required_attendees"],
            optional_attendees=event["optional_attendees"],
            response_status=event["response_status"],
            meeting_status=event["meeting_status"],
            response_requested=event["response_requested"],
        )
        for event in result
    ]


@mcp.tool()
def get_appointment(entry_id: str) -> AppointmentDetails:
    """
    Get full appointment details and body by entry ID.

    Retrieves complete appointment information including body content and all
    meeting metadata using O(1) direct access via EntryID.

    Args:
        entry_id: Outlook EntryID of the appointment (O(1) direct access)

    Returns:
        AppointmentDetails: Complete appointment details including body content

    Raises:
        OutlookNotFoundError: If appointment not found
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Get appointment details from bridge
    result = bridge.get_appointment(entry_id)

    # Check if appointment was found
    if result is None:
        logger.error(f"Appointment not found: {entry_id}")
        raise OutlookNotFoundError("Appointment not found", entry_id=entry_id)

    logger.debug(f"Retrieved appointment: {entry_id}")

    # Convert bridge result to AppointmentDetails model
    # Note: AppointmentDetails extends AppointmentSummary, adding 'body' field
    return AppointmentDetails(
        entry_id=result["entry_id"],
        subject=result["subject"],
        start=result["start"],
        end=result["end"],
        location=result["location"],
        organizer=result["organizer"],
        body=result["body"],
        all_day=result["all_day"],
        required_attendees=result["required_attendees"],
        optional_attendees=result["optional_attendees"],
        response_status=result["response_status"],
        meeting_status=result["meeting_status"],
        response_requested=result["response_requested"],
    )


@mcp.tool()
def delete_appointment(entry_id: str) -> OperationResult:
    """
    Delete an appointment.

    Permanently deletes an appointment using O(1) direct access via EntryID.

    Args:
        entry_id: Outlook EntryID of the appointment (O(1) direct access)

    Returns:
        OperationResult: Result of the operation with success status and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Delete appointment via bridge
    result = bridge.delete_appointment(entry_id)

    # Convert boolean result to OperationResult
    if result:
        return OperationResult(
            success=True,
            message="Appointment deleted successfully",
        )
    else:
        return OperationResult(
            success=False,
            message="Failed to delete appointment",
        )


@mcp.tool()
def create_appointment(
    subject: str,
    start: str,
    end: str,
    location: str = "",
    body: str = "",
    all_day: bool = False,
    required_attendees: str | None = None,
    optional_attendees: str | None = None,
) -> CreateAppointmentResult:
    """
    Create a calendar appointment.

    Creates a new appointment or meeting in the Outlook calendar.
    Supports all-day events, location, body/description, and attendees.

    Args:
        subject: Appointment subject line
        start: Start timestamp in 'YYYY-MM-DD HH:MM:SS' format
        end: End timestamp in 'YYYY-MM-DD HH:MM:SS' format
        location: Appointment location (default: "")
        body: Appointment body/description text (default: "")
        all_day: True for all-day event, False for timed event (default: False)
        required_attendees: Semicolon-separated list of required attendees (optional)
        optional_attendees: Semicolon-separated list of optional attendees (optional)

    Returns:
        CreateAppointmentResult: Result with success status, appointment entry ID (if created), and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Create appointment via bridge
    result = bridge.create_appointment(
        subject=subject,
        start=start,
        end=end,
        location=location,
        body=body,
        all_day=all_day,
        required_attendees=required_attendees,
        optional_attendees=optional_attendees,
    )

    # Convert bridge result to CreateAppointmentResult
    # Bridge returns: str (EntryID) if successful, None if failed
    if result:
        return CreateAppointmentResult(
            success=True,
            entry_id=result,
            message=f"Appointment created successfully: {result}",
        )
    else:
        return CreateAppointmentResult(
            success=False,
            entry_id=None,
            message="Failed to create appointment",
        )


@mcp.tool()
def edit_appointment(
    entry_id: str,
    required_attendees: str | None = None,
    optional_attendees: str | None = None,
    subject: str | None = None,
    start: str | None = None,
    end: str | None = None,
    location: str | None = None,
    body: str | None = None,
) -> OperationResult:
    """
    Edit an existing appointment.

    Updates an existing appointment's fields using O(1) direct access via EntryID.
    Only updates fields that are provided (non-None parameters).

    Args:
        entry_id: Outlook EntryID of the appointment (O(1) direct access)
        required_attendees: Semicolon-separated list of required attendees (optional)
        optional_attendees: Semicolon-separated list of optional attendees (optional)
        subject: New subject (optional)
        start: New start timestamp in 'YYYY-MM-DD HH:MM:SS' format (optional)
        end: New end timestamp in 'YYYY-MM-DD HH:MM:SS' format (optional)
        location: New location (optional)
        body: New body/description text (optional)

    Returns:
        OperationResult: Result of the operation with success status and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Edit appointment via bridge
    result = bridge.edit_appointment(
        entry_id=entry_id,
        required_attendees=required_attendees,
        optional_attendees=optional_attendees,
        subject=subject,
        start=start,
        end=end,
        location=location,
        body=body,
    )

    # Convert boolean result to OperationResult
    if result:
        return OperationResult(
            success=True,
            message="Appointment updated successfully",
        )
    else:
        return OperationResult(
            success=False,
            message="Failed to update appointment",
        )


@mcp.tool()
def respond_to_meeting(entry_id: str, response: str) -> OperationResult:
    """
    Respond to a meeting invitation.

    Responds to a meeting invitation request using O(1) direct access via EntryID.
    Supports accept, decline, and tentative responses.

    Args:
        entry_id: Outlook EntryID of the meeting invitation (O(1) direct access)
        response: Response type - "accept", "decline", or "tentative"

    Returns:
        OperationResult: Result of the operation with success status and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Respond to meeting via bridge
    result = bridge.respond_to_meeting(entry_id, response)

    # Convert boolean result to OperationResult
    if result:
        return OperationResult(
            success=True,
            message=f"Meeting {response.lower()}ed successfully",
        )
    else:
        return OperationResult(
            success=False,
            message=f"Failed to {response.lower()} meeting",
        )


@mcp.tool()
def get_free_busy(
    email_address: str | None = None,
    start_date: str | None = None,
    end_date: str | None = None,
) -> FreeBusyInfo:
    """
    Get free/busy status for an email address.

    Retrieves free/busy information for a specified email address using the
    Outlook FreeBusy method. Defaults to current user and today's date if not specified.

    Args:
        email_address: Email address to check (optional, defaults to current user)
        start_date: Start date in 'YYYY-MM-DD' format (optional, defaults to today)
        end_date: End date in 'YYYY-MM-DD' format (optional, defaults to start + 1 day)

    Returns:
        FreeBusyInfo: Free/busy information with email, dates, status string, and resolution status

    Raises:
        OutlookComError: If bridge is not initialized

    Note:
        Free/busy status codes: 0=Free, 1=Tentative, 2=Busy, 3=Out of Office, 4=Working Elsewhere
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Get free/busy info via bridge
    result = bridge.get_free_busy(
        email_address=email_address,
        start_date=start_date,
        end_date=end_date,
    )

    # Convert bridge result to FreeBusyInfo model
    # Bridge returns dict with different fields depending on success/error
    return FreeBusyInfo(
        email=result.get("email", email_address or "unknown"),
        start_date=result.get("start_date"),
        end_date=result.get("end_date"),
        free_busy=result.get("free_busy"),
        resolved=result.get("resolved", False),
        error=result.get("error"),
    )


# ============================================================================
# Task Tools
# ============================================================================


@mcp.tool()
def list_tasks(include_completed: bool = False) -> list[TaskSummary]:
    """
    List tasks from the Outlook Tasks folder.

    Retrieves a list of tasks from the Outlook Tasks folder.
    By default, returns only incomplete tasks. Optionally includes completed tasks.

    Args:
        include_completed: If True, return all tasks. If False (default), return only incomplete tasks.

    Returns:
        list[TaskSummary]: List of task summaries with basic information

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # List tasks via bridge
    result = bridge.list_tasks(include_completed=include_completed)

    # Convert bridge result to list of TaskSummary models
    return [
        TaskSummary(
            entry_id=task["entry_id"],
            subject=task["subject"],
            body=task["body"],
            due_date=task["due_date"],
            status=task["status"],
            priority=task["priority"],
            complete=task["complete"],
            percent_complete=task["percent_complete"],
        )
        for task in result
    ]


@mcp.tool()
def get_task(entry_id: str) -> TaskSummary:
    """
    Get full task details and body by entry ID.

    Retrieves complete task information including body content and all
    task metadata using O(1) direct access via EntryID.

    Args:
        entry_id: Outlook EntryID of the task (O(1) direct access)

    Returns:
        TaskSummary: Complete task details including body content

    Raises:
        OutlookNotFoundError: If task not found
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Get task details from bridge
    result = bridge.get_task(entry_id)

    # Check if task was found
    if result is None:
        logger.error(f"Task not found: {entry_id}")
        raise OutlookNotFoundError("Task not found", entry_id=entry_id)

    logger.debug(f"Retrieved task: {entry_id}")

    # Convert bridge result to TaskSummary model
    # Note: TaskSummary includes all fields from bridge.get_task()
    return TaskSummary(
        entry_id=result["entry_id"],
        subject=result["subject"],
        body=result["body"],
        due_date=result["due_date"],
        status=result["status"],
        priority=result["priority"],
        complete=result["complete"],
        percent_complete=result["percent_complete"],
    )


@mcp.tool()
def complete_task(entry_id: str) -> OperationResult:
    """
    Mark a task as complete.

    Marks a task as complete with 100% percent complete status using O(1)
    direct access via EntryID.

    Args:
        entry_id: Outlook EntryID of the task (O(1) direct access)

    Returns:
        OperationResult: Result of the operation with success status and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Mark task as complete via bridge
    result = bridge.complete_task(entry_id)

    # Convert boolean result to OperationResult
    if result:
        return OperationResult(
            success=True,
            message="Task marked as complete",
        )
    else:
        return OperationResult(
            success=False,
            message="Failed to mark task as complete",
        )


@mcp.tool()
def delete_task(entry_id: str) -> OperationResult:
    """
    Delete a task.

    Permanently deletes a task using O(1) direct access via EntryID.

    Args:
        entry_id: Outlook EntryID of the task (O(1) direct access)

    Returns:
        OperationResult: Result of the operation with success status and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Delete task via bridge
    result = bridge.delete_task(entry_id)

    # Convert boolean result to OperationResult
    if result:
        return OperationResult(
            success=True,
            message="Task deleted successfully",
        )
    else:
        return OperationResult(
            success=False,
            message="Failed to delete task",
        )


@mcp.tool()
def create_task(
    subject: str,
    body: str = "",
    due_date: str | None = None,
    priority: int = 1,
) -> CreateTaskResult:
    """
    Create a new task.

    Creates a new task in the Outlook Tasks folder.
    Supports task description, due date, and priority level.

    Args:
        subject: Task subject/title
        body: Task description or body text (default: "")
        due_date: Due date in 'YYYY-MM-DD' format (optional)
        priority: Task priority - 0=Low, 1=Normal (default), 2=High

    Returns:
        CreateTaskResult: Result with success status, task entry ID (if created), and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Create task via bridge
    # Note: bridge parameter is 'importance', not 'priority'
    result = bridge.create_task(
        subject=subject,
        body=body,
        due_date=due_date,
        importance=priority,
    )

    # Convert bridge result to CreateTaskResult
    # Bridge returns: str (EntryID) if successful, None if failed
    if result:
        return CreateTaskResult(
            success=True,
            entry_id=result,
            message=f"Task created successfully: {result}",
        )
    else:
        return CreateTaskResult(
            success=False,
            entry_id=None,
            message="Failed to create task",
        )


@mcp.tool()
def edit_task(
    entry_id: str,
    subject: str | None = None,
    body: str | None = None,
    due_date: str | None = None,
    priority: int | None = None,
    percent_complete: float | None = None,
    complete: bool | None = None,
) -> OperationResult:
    """
    Edit an existing task.

    Updates an existing task in the Outlook Tasks folder.
    Only updates fields that are provided (non-None parameters).
    Supports updating subject, description, due date, priority, completion status, and percent complete.

    Args:
        entry_id: Task entry ID to edit
        subject: New task subject/title (optional)
        body: New task description or body text (optional)
        due_date: New due date in 'YYYY-MM-DD' format (optional)
        priority: New task priority - 0=Low, 1=Normal, 2=High (optional)
        percent_complete: New percent complete value 0-100 (optional)
        complete: Mark task as complete or incomplete (optional)

    Returns:
        OperationResult: Result with success status and message

    Raises:
        OutlookComError: If bridge is not initialized
    """
    # Get bridge from module-level state
    bridge = _get_bridge()

    # Edit task via bridge
    # Note: bridge parameter is 'importance', not 'priority'
    result = bridge.edit_task(
        entry_id=entry_id,
        subject=subject,
        body=body,
        due_date=due_date,
        importance=priority,
        percent_complete=percent_complete,
        complete=complete,
    )

    # Convert bridge result to OperationResult
    # Bridge returns: True if successful, False if failed
    if result:
        return OperationResult(success=True, message="Task edited successfully")
    else:
        return OperationResult(success=False, message="Failed to edit task")


def main(default_account: str | None = None):
    """Entry point for the MCP server.

    Args:
        default_account: Optional default account name/email (bypasses argparse if provided)
    """
    global _default_account

    # If default_account is provided directly, use it
    # Otherwise, parse CLI arguments
    if default_account is None:
        # Parse CLI arguments for default account
        parser = argparse.ArgumentParser(
            description="Mailtool MCP Server - Outlook automation via MCP",
            epilog=(
                "Examples:\n"
                "  uv run --with pywin32 -m mailtool.mcp.server\n"
                "  uv run --with pywin32 -m mailtool.mcp.server --account 'john@example.com'\n"
                "  uv run --with pywin32 -m mailtool.mcp.server --acc 'john@example.com'\n"
            ),
            formatter_class=argparse.RawDescriptionHelpFormatter,
        )
        parser.add_argument(
            "--account",
            "--acc",
            dest="account",
            help="Default account name or email address for Outlook operations",
        )

        args = parser.parse_args()
        default_account = args.account

    # Set the global default account that the lifespan will read
    _default_account = default_account

    # Run the MCP server with stdio transport
    # Uses the global 'mcp' instance which has all tools registered
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
