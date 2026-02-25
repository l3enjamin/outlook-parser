"""COM State Management for MCP Server

This module provides thread-safe COM initialization tracking for the MCP server.
COM must be initialized once per thread that accesses COM objects.

This module is shared between server.py and resources.py to ensure consistent
COM state across all MCP tools and resources.
"""

import logging
import threading

import pythoncom

# Configure logging
logger = logging.getLogger(__name__)

# Thread-local storage for per-thread COM state.
# threading.local() provides implicit per-thread isolation with no locking required.
_thread_local = threading.local()


class _ComCleanup:
    """Triggers CoUninitialize when a thread's local storage is cleaned up.

    When a thread exits, Python cleans up thread-local variables. This class
    uses __del__ to ensure CoUninitialize is called for every thread that
    initialized COM, preventing COM resource leaks from worker threads that
    die quietly after processing MCP tool calls.
    """

    def __del__(self) -> None:
        try:
            pythoncom.CoUninitialize()
            logger.debug("COM uninitialized via thread-local cleanup")
        except Exception as e:
            logger.warning(f"Error during COM cleanup in thread-local finalizer: {e}")


def ensure_com_initialized() -> None:
    """Ensure COM is initialized for the current thread.

    This function is called by every MCP tool and resource to ensure COM is available
    in the calling thread. COM is initialized at most once per thread; a _ComCleanup
    sentinel stored in thread-local storage ensures CoUninitialize is called
    automatically when the thread exits, preventing resource leaks.

    Thread Safety:
        This function is thread-safe — threading.local() provides implicit
        per-thread isolation with no locking required.
    """
    if not getattr(_thread_local, "com_initialized", False):
        thread_id = threading.get_ident()
        logger.debug(f"Initializing COM for thread {thread_id}")
        pythoncom.CoInitialize()
        _thread_local.com_initialized = True
        # Attach cleanup sentinel — __del__ fires when thread-local storage is freed on exit
        _thread_local._cleanup = _ComCleanup()
        logger.debug(f"COM initialized for thread {thread_id}")


def is_com_initialized_for_thread(thread_id: int | None = None) -> bool:
    """Check if COM is initialized for the current thread.

    Args:
        thread_id: Unused, kept for API compatibility. Always checks the current thread
                   since per-thread state is stored in threading.local().

    Returns:
        bool: True if COM is initialized for the current thread.
    """
    return getattr(_thread_local, "com_initialized", False)
