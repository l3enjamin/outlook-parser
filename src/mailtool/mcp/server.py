"""Mailtool MCP Server

This module provides the main FastMCP server instance for Outlook automation.
It implements the Model Context Protocol (MCP) using the official MCP Python SDK v2
with the FastMCP framework.

The server provides 23 tools and 7 resources for Outlook email, calendar, and task management.
All tools return structured Pydantic models for type safety and LLM understanding.
"""

from mcp.server import FastMCP

from mailtool.mcp.lifespan import outlook_lifespan

# Create FastMCP server instance
# The lifespan parameter manages Outlook COM bridge lifecycle (creation, warmup, cleanup)
mcp = FastMCP(
    name="mailtool-outlook-bridge",
    lifespan=outlook_lifespan,
)

# Tools and resources will be registered here in subsequent user stories
# TODO: Register tools (US-008 to US-037)
# TODO: Register resources (US-022 to US-037)


if __name__ == "__main__":
    # Run the MCP server with stdio transport
    # This is the standard transport for MCP clients like Claude Code
    mcp.run(transport="stdio")
