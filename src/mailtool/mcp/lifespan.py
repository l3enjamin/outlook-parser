"""Lifespan Management for MCP Server

This module provides Outlook bridge lifecycle management for the MCP server.
It handles:
- Creating and warming up OutlookBridge instance on startup
- Setting module-level bridge state for tool access
- Releasing COM objects and forcing garbage collection on shutdown
"""

import asyncio
import gc
import logging
from contextlib import asynccontextmanager

from mailtool.bridge import OutlookBridge

# Configure logging for the lifespan manager
# Logs are written to stderr for debugging and monitoring
logger = logging.getLogger(__name__)


@asynccontextmanager
async def outlook_lifespan(app):
    """Async context manager for Outlook bridge lifecycle

    This function manages the complete lifecycle of the Outlook COM bridge:
    1. Creates OutlookBridge instance on startup
    2. Warms up the connection with retry attempts
    3. Sets module-level bridge state for tool access
    4. Cleans up COM objects on shutdown

    Args:
        app: The FastMCP server instance (used to access module state)

    Yields:
        None: The bridge is set in server._bridge module state

    Raises:
        Exception: If Outlook cannot be connected to after retry attempts
    """
    bridge = None
    try:
        logger.info("Initializing Outlook bridge...")

        # Create Outlook bridge instance (synchronous COM call)
        # Note: We run this in a thread pool since COM calls are synchronous
        loop = asyncio.get_event_loop()
        bridge = await loop.run_in_executor(None, _create_bridge)

        logger.info("Outlook bridge created successfully")

        # Warmup: Test that COM is responsive with retries
        max_retries = 5
        retry_delay = 0.5  # seconds

        for attempt in range(1, max_retries + 1):
            try:
                logger.debug(f"Warmup attempt {attempt}/{max_retries}")
                # Run a real COM call to ensure Outlook is responsive
                await loop.run_in_executor(None, _warmup_bridge, bridge)
                logger.info("Outlook bridge warmed up successfully")
                break  # Success - exit retry loop
            except Exception as e:
                logger.warning(f"Warmup attempt {attempt}/{max_retries} failed: {e}")
                if attempt == max_retries:
                    logger.error(
                        f"Outlook warmup failed after {max_retries} attempts: {e}"
                    )
                    raise Exception(
                        f"Outlook warmup failed after {max_retries} attempts: {e}"
                    ) from e
                # Wait before retry
                await asyncio.sleep(retry_delay)

        # Set module-level bridge state for tools to access
        # Import here to avoid circular imports
        from mailtool.mcp import server

        server._bridge = bridge

        # Set bridge in resources module for resource access
        from mailtool.mcp import resources

        resources._set_bridge(bridge)

        logger.info("Outlook bridge initialized and ready")

        # Yield for server to start
        yield

    finally:
        # Cleanup: Release COM objects and force garbage collection
        logger.info("Shutting down Outlook bridge...")
        if bridge is not None:
            try:
                # Release COM references
                bridge.outlook = None
                bridge.namespace = None
                logger.debug("Released COM references")
            except Exception as e:
                logger.error(f"Error releasing COM references: {e}")

        # Force Python garbage collection to release COM objects
        gc.collect()
        logger.info("Outlook bridge shutdown complete")


def _create_bridge() -> OutlookBridge:
    """Synchronous function to create OutlookBridge instance

    Returns:
        OutlookBridge: The initialized bridge instance

    Raises:
        Exception: If Outlook cannot be connected to or launched
    """
    logger.debug("Creating OutlookBridge instance")
    return OutlookBridge()


def _warmup_bridge(bridge: OutlookBridge) -> None:
    """Synchronous warmup function to test COM connectivity

    Args:
        bridge: The OutlookBridge instance to test

    Raises:
        Exception: If COM call fails (Outlook not responsive)
    """
    logger.debug("Testing COM connectivity via warmup")
    inbox = bridge.get_inbox()
    # Make a real COM call to test connectivity
    count = inbox.Items.Count
    logger.debug(f"Warmup successful: Inbox has {count} items")
