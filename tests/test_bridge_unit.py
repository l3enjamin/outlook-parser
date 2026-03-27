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

    def test_download_attachments_optimized(self, tmp_path):
        """Test download_attachments with the new optimized path handling"""
        import os
        from unittest.mock import MagicMock, patch

        bridge = OutlookBridge()

        # Mocking the item and attachments
        mock_item = MagicMock()
        mock_attachments = MagicMock()
        mock_item.Attachments = mock_attachments
        mock_attachments.Count = 2

        att1 = MagicMock()
        att1.FileName = "test1.txt"
        att2 = MagicMock()
        att2.FileName = "test2.txt"

        mock_attachments.Item.side_effect = [att1, att2]

        bridge.get_item_by_id = MagicMock(return_value=mock_item)

        download_dir = str(tmp_path / "downloads")

        # We need to patch os.path.abspath to return predictable results in the mock environment
        with patch("os.path.abspath", side_effect=lambda x: x):
            downloaded = bridge.download_attachments("fake_id", download_dir)

        assert len(downloaded) == 2
        assert os.path.join(download_dir, "test1.txt") in downloaded
        assert os.path.join(download_dir, "test2.txt") in downloaded

        from unittest.mock import call

        # Verify optimization: Attachments and Count properties were each accessed only once.
        assert mock_item.mock_calls == [call.Attachments]
        assert mock_attachments.mock_calls == [call.Count, call.Item(1), call.Item(2)]
