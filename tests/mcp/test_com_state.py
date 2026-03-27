"""Tests for COM state management.

This module tests the thread-local COM initialization logic used by the MCP server.
"""

import threading
from unittest.mock import MagicMock, patch

import pytest

# Mock pythoncom before importing com_state to avoid any side effects
# although com_state only imports it, it doesn't call it at module level.
with patch("pythoncom.CoInitialize"), patch("pythoncom.CoUninitialize"):
    from mailtool.mcp.com_state import (
        _ComCleanup,
        _thread_local,
        ensure_com_initialized,
        is_com_initialized_for_thread,
    )


@pytest.fixture(autouse=True)
def reset_thread_local():
    """Reset the thread-local state before each test."""
    if hasattr(_thread_local, "com_initialized"):
        del _thread_local.com_initialized
    if hasattr(_thread_local, "_cleanup"):
        del _thread_local._cleanup
    yield


class TestComState:
    """Tests for COM state management functions."""

    @patch("pythoncom.CoInitialize")
    def test_ensure_com_initialized_first_time(self, mock_co_init):
        """Test that COM is initialized on the first call."""
        assert not is_com_initialized_for_thread()

        ensure_com_initialized()

        mock_co_init.assert_called_once()
        assert is_com_initialized_for_thread()
        assert hasattr(_thread_local, "_cleanup")
        assert isinstance(_thread_local._cleanup, _ComCleanup)

    @patch("pythoncom.CoInitialize")
    def test_ensure_com_initialized_idempotent(self, mock_co_init):
        """Test that COM is only initialized once per thread."""
        ensure_com_initialized()
        assert mock_co_init.call_count == 1

        # Second call should not trigger CoInitialize again
        ensure_com_initialized()
        assert mock_co_init.call_count == 1

    def test_is_com_initialized_for_thread(self):
        """Test is_com_initialized_for_thread returns correct status."""
        assert not is_com_initialized_for_thread()

        _thread_local.com_initialized = True
        assert is_com_initialized_for_thread()

        _thread_local.com_initialized = False
        assert not is_com_initialized_for_thread()

    @patch("pythoncom.CoInitialize")
    def test_thread_isolation(self, mock_co_init):
        """Test that COM state is isolated between threads."""
        # Initialize in main thread
        ensure_com_initialized()
        assert is_com_initialized_for_thread()
        assert mock_co_init.call_count == 1

        results = {"initialized": None}

        def check_thread():
            # In a new thread, it should not be initialized
            results["initialized"] = is_com_initialized_for_thread()
            # And we can initialize it independently
            ensure_com_initialized()

        thread = threading.Thread(target=check_thread)
        thread.start()
        thread.join()

        assert results["initialized"] is False
        # CoInitialize should have been called again (once for the new thread)
        assert mock_co_init.call_count == 2


class TestComCleanup:
    """Tests for _ComCleanup sentinel class."""

    @patch("pythoncom.CoUninitialize")
    def test_cleanup_on_deletion(self, mock_co_uninit):
        """Test that CoUninitialize is called when _ComCleanup is deleted."""
        cleanup = _ComCleanup()

        # Manually trigger deletion
        del cleanup

        mock_co_uninit.assert_called_once()

    @patch("pythoncom.CoUninitialize")
    def test_cleanup_handles_exception(self, mock_co_uninit):
        """Test that cleanup handles exceptions during CoUninitialize."""
        mock_co_uninit.side_effect = Exception("COM Error")

        cleanup = _ComCleanup()

        # This should not raise an exception even if CoUninitialize fails
        # since it's called during __del__ and wrapped in try-except
        try:
            del cleanup
        except Exception as e:
            pytest.fail(f"_ComCleanup.__del__ raised an exception: {e}")

        mock_co_uninit.assert_called_once()
