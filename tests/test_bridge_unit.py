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

        # 3. Attribute access raising specific COM error (if available)
        # Since we might be mocking win32com, let's just use a general Exception subclass
        # that mimics a COM error structure if needed, but Exception is broad enough.
        # The code catches Exception, so any exception works.
