"""
Mailtool: Outlook COM Automation Bridge for WSL2

A WSL2-to-Windows bridge for Outlook automation via COM.
Optimized for AI agent integration with O(1) access patterns.
"""

# Lazy import to allow package import on non-Windows platforms
try:
    from mailtool.bridge import OutlookBridge
except ImportError:
    OutlookBridge = None

__version__ = "2.1.0"
__all__ = ["OutlookBridge"]
