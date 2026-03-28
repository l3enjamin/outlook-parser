import unittest
from unittest.mock import MagicMock, patch
import sys
import os

# Set up PYTHONPATH
sys.path.insert(0, os.path.join(os.getcwd(), "src"))

# Mock necessary modules
sys.modules["win32com"] = MagicMock()
sys.modules["win32com.client"] = MagicMock()
sys.modules["pythoncom"] = MagicMock()
sys.modules["mailparser"] = MagicMock()
sys.modules["mailparser_reply"] = MagicMock()
sys.modules["mcp"] = MagicMock()
sys.modules["mcp.server"] = MagicMock()
sys.modules["mailtool.mcp.server"] = MagicMock()

from mailtool.bridge import OutlookBridge
from mailtool.cli import main

class TestSearchSecurity(unittest.TestCase):
    def setUp(self):
        self.bridge = OutlookBridge()
        self.bridge._search_emails_raw = MagicMock()
        self.bridge.list_emails = MagicMock()

    def test_search_emails_no_filter_query(self):
        """Verify that filter_query parameter is removed and logic using it is gone."""
        # Check that search_emails no longer accepts filter_query
        import inspect
        sig = inspect.signature(self.bridge.search_emails)
        self.assertNotIn('filter_query', sig.parameters)

    def test_search_emails_structured(self):
        """Verify that structured search still works."""
        self.bridge.search_emails(subject="TestSubject", limit=50)

        # Verify it calls _search_emails_raw with a DASL query
        self.bridge._search_emails_raw.assert_called_once()
        args, kwargs = self.bridge._search_emails_raw.call_args
        query = args[0]
        self.assertIn("@SQL=", query)
        self.assertIn("urn:schemas:httpmail:subject", query)
        self.assertIn("TestSubject", query)
        self.assertEqual(args[1], 50)

    def test_search_emails_escaping(self):
        """Verify that special characters (single quotes) are escaped in DASL query."""
        malicious_subject = "Project' OR '1'='1"
        self.bridge.search_emails(subject=malicious_subject)

        self.bridge._search_emails_raw.assert_called_once()
        args, _ = self.bridge._search_emails_raw.call_args
        query = args[0]

        # Single quotes should be doubled in DASL/SQL
        expected_escaped = "Project'' OR ''1''=''1"
        self.assertIn(expected_escaped, query)
        # Verify it doesn't contain the unescaped single quote that could break the query
        # (Except where it's part of the doubled quote or the surrounding quotes of the LIKE pattern)
        # Actually, DASL uses single quotes for string literals.
        # The filter generated is: "urn:schemas:httpmail:subject" LIKE '%Project'' OR ''1''=''1%'
        self.assertIn(f"LIKE '%{expected_escaped}%'", query)

    @patch("mailtool.bridge.OutlookBridge")
    @patch("mailtool.cli._check_platform")
    @patch("mailtool.cli._check_pywin32")
    def test_cli_search_no_query_arg(self, mock_check_py, mock_check_plat, mock_bridge_cls):
        """Verify that CLI search command no longer has --query argument."""
        # We need to ensure that the mocked OutlookBridge is the one called in mailtool.cli.main
        mock_bridge = mock_bridge_cls.return_value

        # Try to run with structured args, it should succeed
        with patch("sys.argv", ["mailtool", "search", "--subject", "TestSubject"]):
            mock_bridge.search_emails.return_value = []
            with patch("sys.stdout"):
                main()
            mock_bridge.search_emails.assert_called_with(
                limit=100,
                subject="TestSubject",
                sender=None,
                body=None,
                unread=None,
                has_attachments=None
            )

        # Try to run with --query, it should fail in argparse (SystemExit 2)
        with patch("sys.argv", ["mailtool", "search", "--query", "some raw query"]):
            # argparse uses stderr to print errors
            with patch("sys.stderr") as mock_stderr:
                with self.assertRaises(SystemExit):
                    main()

if __name__ == "__main__":
    unittest.main()
