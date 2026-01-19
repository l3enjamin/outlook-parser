"""
Pydantic Models for MCP Tools

This module contains all Pydantic models for structured output from MCP tools.
Models provide type safety, automatic schema generation, and descriptive field metadata.
"""

from pydantic import BaseModel, Field

# ============================================================================
# Email Models (US-004)
# ============================================================================


class EmailSummary(BaseModel):
    """Summary representation of an email for list views"""

    entry_id: str = Field(description="Outlook EntryID for O(1) direct access")
    subject: str = Field(description="Email subject line")
    sender: str = Field(description="SMTP email address of sender")
    sender_name: str = Field(description="Display name of sender")
    received_time: str | None = Field(
        default=None,
        description="Received timestamp in 'YYYY-MM-DD HH:MM:SS' format or None",
    )
    unread: bool = Field(description="Whether the email is unread")
    has_attachments: bool = Field(description="Whether the email has attachments")


class EmailDetails(BaseModel):
    """Full email details including body content"""

    entry_id: str = Field(description="Outlook EntryID for O(1) direct access")
    subject: str = Field(description="Email subject line")
    sender: str = Field(description="SMTP email address of sender")
    sender_name: str = Field(description="Display name of sender")
    body: str = Field(description="Plain text email body")
    html_body: str = Field(description="HTML email body")
    received_time: str | None = Field(
        default=None,
        description="Received timestamp in 'YYYY-MM-DD HH:MM:SS' format or None",
    )
    has_attachments: bool = Field(description="Whether the email has attachments")


class SendEmailResult(BaseModel):
    """Result of sending an email or saving a draft"""

    success: bool = Field(description="Whether the operation succeeded")
    entry_id: str | None = Field(
        default=None, description="EntryID of saved draft (None if sent or failed)"
    )
    message: str = Field(description="Human-readable result message")


# ============================================================================
# Calendar Models (US-005)
# ============================================================================


class AppointmentSummary(BaseModel):
    """Summary representation of a calendar event for list views"""

    entry_id: str = Field(description="Outlook EntryID for O(1) direct access")
    subject: str = Field(description="Appointment subject line")
    start: str | None = Field(
        default=None,
        description="Start timestamp in 'YYYY-MM-DD HH:MM:SS' format or None",
    )
    end: str | None = Field(
        default=None,
        description="End timestamp in 'YYYY-MM-DD HH:MM:SS' format or None",
    )
    location: str = Field(description="Appointment location", default="")
    organizer: str | None = Field(
        default=None, description="Email address of the meeting organizer"
    )
    all_day: bool = Field(description="Whether this is an all-day event")
    required_attendees: str = Field(
        default="", description="Semicolon-separated list of required attendees"
    )
    optional_attendees: str = Field(
        default="", description="Semicolon-separated list of optional attendees"
    )
    response_status: str = Field(
        description="Meeting response status: None, Organizer, Tentative, Accepted, Declined, NotResponded, Unknown"
    )
    meeting_status: str = Field(
        description="Meeting status: NonMeeting, Meeting, Received, Canceled, Unknown"
    )
    response_requested: bool = Field(description="Whether response was requested")


class AppointmentDetails(AppointmentSummary):
    """Full appointment details including body content"""

    body: str = Field(default="", description="Appointment body/description text")


class CreateAppointmentResult(BaseModel):
    """Result of creating an appointment"""

    success: bool = Field(description="Whether the operation succeeded")
    entry_id: str | None = Field(
        default=None, description="EntryID of created appointment (None if failed)"
    )
    message: str = Field(description="Human-readable result message")


class FreeBusyInfo(BaseModel):
    """Free/busy information for an email address"""

    email: str = Field(description="Email address that was checked")
    start_date: str | None = Field(
        default=None, description="Start date in 'YYYY-MM-DD' format or None if error"
    )
    end_date: str | None = Field(
        default=None, description="End date in 'YYYY-MM-DD' format or None if error"
    )
    free_busy: str | None = Field(
        default=None,
        description="Free/busy string with time slot status codes (0=Free, 1=Tentative, 2=Busy, 3=OOF, 4=Working Elsewhere) or None if error",
    )
    resolved: bool = Field(
        description="Whether the email address was successfully resolved"
    )
    error: str | None = Field(
        default=None, description="Error message if the operation failed"
    )


# ============================================================================
# Task Models (US-006)
# ============================================================================
# TODO: Implement TaskSummary, CreateTaskResult


# ============================================================================
# Common Result Models (US-007)
# ============================================================================
# TODO: Implement OperationResult
