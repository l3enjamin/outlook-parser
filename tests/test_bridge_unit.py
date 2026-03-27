"""
Unit tests for OutlookBridge helper methods.
"""
import pytest
from mailtool.bridge import OutlookBridge

class TestOutlookBridgeUnit:
    """Unit tests for OutlookBridge static methods and helpers"""

    def test_safe_get_attr(self):
        """Test _safe_get_attr handles normal access, missing attrs, and exceptions"""

        # 1. Normal attribute access
        class NormalObject:
            def __init__(self):
                self.value = "test_value"

        obj = NormalObject()
        assert OutlookBridge._safe_get_attr(obj, "value") == "test_value"
        assert OutlookBridge._safe_get_attr(obj, "missing", default="default") == "default"

        # 2. Attribute access raising Exception (simulating COM error)
        class ErrorObject:
            @property
            def error_attr(self):
                raise Exception("COM Error")

        error_obj = ErrorObject()
        assert OutlookBridge._safe_get_attr(error_obj, "error_attr", default="error_default") == "error_default"

    def test_escape_dasl_query(self):
        """Test _escape_dasl_query escapes single quotes and handles None/empty strings"""

        # Normal case
        assert OutlookBridge._escape_dasl_query("normal") == "normal"

        # Single quote
        assert OutlookBridge._escape_dasl_query("o'reilly") == "o''reilly"

        # Multiple single quotes
        assert OutlookBridge._escape_dasl_query("'quoted'") == "''quoted''"

        # Empty string
        assert OutlookBridge._escape_dasl_query("") == ""

        # None
        assert OutlookBridge._escape_dasl_query(None) == ""

        # Non-string (should be converted to string and escaped if needed)
        assert OutlookBridge._escape_dasl_query(123) == "123"

        # Single quote at start and end
        assert OutlookBridge._escape_dasl_query("'") == "''"

        # String with multiple quotes and other characters
        assert OutlookBridge._escape_dasl_query("It's a 'test' string") == "It''s a ''test'' string"

        # 3. Attribute access raising specific COM error (if available)
        # Since we might be mocking win32com, let's just use a general Exception subclass
        # that mimics a COM error structure if needed, but Exception is broad enough.
        # The code catches Exception, so any exception works.
