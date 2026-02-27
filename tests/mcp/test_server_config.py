"""Unit tests for MCP server configuration logic."""

import sys
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

# Add src directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent.parent / "src"))

# Mock missing dependencies for environment without them
if "mcp" not in sys.modules:
    mcp_mock = MagicMock()
    sys.modules["mcp"] = mcp_mock
    sys.modules["mcp.server"] = MagicMock()
    sys.modules["mcp.shared"] = MagicMock()
    sys.modules["mcp.shared.exceptions"] = MagicMock()
    sys.modules["mcp.types"] = MagicMock()

if "pydantic" not in sys.modules:
    pydantic_mock = MagicMock()
    sys.modules["pydantic"] = pydantic_mock
    # Ensure BaseModel is a class that can be inherited from
    class BaseModel:
        def __init__(self, **kwargs):
            for k, v in kwargs.items():
                setattr(self, k, v)
    pydantic_mock.BaseModel = BaseModel

if "pydantic_settings" not in sys.modules:
    sys.modules["pydantic_settings"] = MagicMock()

from mailtool.mcp.server import configure_server_features

# Mock tool sets for testing
MOCK_MAIL_TOOLS = {"mail_ro", "mail_rw"}
MOCK_CALENDAR_TOOLS = {"cal_ro", "cal_rw"}
MOCK_TASK_TOOLS = {"task_ro", "task_rw"}
MOCK_READ_ONLY_TOOLS = {"mail_ro", "cal_ro", "task_ro"}
MOCK_ALL_TOOLS = MOCK_MAIL_TOOLS | MOCK_CALENDAR_TOOLS | MOCK_TASK_TOOLS

@pytest.fixture
def mock_mcp():
    """Create a mock FastMCP instance."""
    return MagicMock()

@pytest.fixture
def patched_tools():
    """Patch the tool constants in the server module."""
    with patch("mailtool.mcp.server.MAIL_TOOLS", MOCK_MAIL_TOOLS), \
         patch("mailtool.mcp.server.CALENDAR_TOOLS", MOCK_CALENDAR_TOOLS), \
         patch("mailtool.mcp.server.TASK_TOOLS", MOCK_TASK_TOOLS), \
         patch("mailtool.mcp.server.READ_ONLY_TOOLS", MOCK_READ_ONLY_TOOLS), \
         patch("mailtool.mcp.server.ALL_TOOLS", MOCK_ALL_TOOLS):
        yield

@pytest.fixture
def mock_registrations():
    """Patch the resource registration functions."""
    with patch("mailtool.mcp.server.register_email_resources") as m_mail, \
         patch("mailtool.mcp.server.register_calendar_resources") as m_cal, \
         patch("mailtool.mcp.server.register_task_resources") as m_task:
        yield m_mail, m_cal, m_task

def test_configure_all_enabled_read_write(mock_mcp, patched_tools, mock_registrations):
    """Verify configuration with all features enabled in RW mode."""
    m_mail, m_cal, m_task = mock_registrations

    configure_server_features(
        mock_mcp,
        enable_mail=True,
        enable_calendar=True,
        enable_tasks=True,
        is_rw=True
    )

    # In RW mode with all enabled, no tools should be removed
    assert mock_mcp.remove_tool.call_count == 0

    # All resources should be registered
    m_mail.assert_called_once_with(mock_mcp)
    m_cal.assert_called_once_with(mock_mcp)
    m_task.assert_called_once_with(mock_mcp)

def test_configure_mail_only_read_only(mock_mcp, patched_tools, mock_registrations):
    """Verify configuration with only mail enabled in read-only mode."""
    m_mail, m_cal, m_task = mock_registrations

    configure_server_features(
        mock_mcp,
        enable_mail=True,
        enable_calendar=False,
        enable_tasks=False,
        is_rw=False
    )

    # Should keep only mail_ro
    # Should remove: mail_rw, cal_ro, cal_rw, task_ro, task_rw
    removed_tools = {call.args[0] for call in mock_mcp.remove_tool.call_args_list}
    expected_removed = {"mail_rw", "cal_ro", "cal_rw", "task_ro", "task_rw"}
    assert removed_tools == expected_removed

    # Only mail resources should be registered
    m_mail.assert_called_once_with(mock_mcp)
    m_cal.assert_not_called()
    m_task.assert_not_called()

def test_configure_calendar_only_read_write(mock_mcp, patched_tools, mock_registrations):
    """Verify configuration with only calendar enabled in RW mode."""
    m_mail, m_cal, m_task = mock_registrations

    configure_server_features(
        mock_mcp,
        enable_mail=False,
        enable_calendar=True,
        enable_tasks=False,
        is_rw=True
    )

    # Should keep: cal_ro, cal_rw
    # Should remove: mail_ro, mail_rw, task_ro, task_rw
    removed_tools = {call.args[0] for call in mock_mcp.remove_tool.call_args_list}
    expected_removed = {"mail_ro", "mail_rw", "task_ro", "task_rw"}
    assert removed_tools == expected_removed

    # Only calendar resources should be registered
    m_mail.assert_not_called()
    m_cal.assert_called_once_with(mock_mcp)
    m_task.assert_not_called()

def test_configure_tasks_only_read_only(mock_mcp, patched_tools, mock_registrations):
    """Verify configuration with only tasks enabled in read-only mode."""
    m_mail, m_cal, m_task = mock_registrations

    configure_server_features(
        mock_mcp,
        enable_mail=False,
        enable_calendar=False,
        enable_tasks=True,
        is_rw=False
    )

    # Should keep: task_ro
    # Should remove: mail_ro, mail_rw, cal_ro, cal_rw, task_rw
    removed_tools = {call.args[0] for call in mock_mcp.remove_tool.call_args_list}
    expected_removed = {"mail_ro", "mail_rw", "cal_ro", "cal_rw", "task_rw"}
    assert removed_tools == expected_removed

    # Only task resources should be registered
    m_mail.assert_not_called()
    m_cal.assert_not_called()
    m_task.assert_called_once_with(mock_mcp)

def test_configure_all_disabled(mock_mcp, patched_tools, mock_registrations):
    """Verify configuration with all features disabled."""
    m_mail, m_cal, m_task = mock_registrations

    configure_server_features(
        mock_mcp,
        enable_mail=False,
        enable_calendar=False,
        enable_tasks=False,
        is_rw=False
    )

    # Should remove all tools
    removed_tools = {call.args[0] for call in mock_mcp.remove_tool.call_args_list}
    assert removed_tools == MOCK_ALL_TOOLS

    # No resources should be registered
    m_mail.assert_not_called()
    m_cal.assert_not_called()
    m_task.assert_not_called()
