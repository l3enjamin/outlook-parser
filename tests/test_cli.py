import sys
import unittest
from unittest.mock import MagicMock, patch

# Mock win32 modules before importing anything that might use them
sys.modules["win32com"] = MagicMock()
sys.modules["win32com.client"] = MagicMock()
sys.modules["pythoncom"] = MagicMock()

# Mock mailparser and mailparser_reply
sys.modules["mailparser"] = MagicMock()
sys.modules["mailparser_reply"] = MagicMock()

# Mock mcp.server to avoid importing real MCP dependencies
sys.modules["mcp"] = MagicMock()
sys.modules["mcp.server"] = MagicMock()
# Mock mailtool.mcp.server entirely to avoid complex dependency chains
sys.modules["mailtool.mcp.server"] = MagicMock()

# Now we can safely import mailtool.cli
from mailtool.cli import main, _check_pywin32, _check_platform  # noqa: E402


class TestPyWin32Check(unittest.TestCase):
    def setUp(self):
        self.stderr_patch = patch("sys.stderr")
        self.mock_stderr = self.stderr_patch.start()

    def tearDown(self):
        self.stderr_patch.stop()

    def test_check_pywin32_success(self):
        """Test pywin32 check when package is available"""
        with patch("importlib.util.find_spec") as mock_find_spec:
            # Return a mock spec (not None) to simulate package presence
            mock_find_spec.return_value = MagicMock()
            # Should not raise any exception
            _check_pywin32()
            mock_find_spec.assert_called_with("win32com.client")

    def test_check_pywin32_missing(self):
        """Test pywin32 check when package is missing"""
        with patch("importlib.util.find_spec") as mock_find_spec:
            # Return None to simulate missing package
            mock_find_spec.return_value = None

            with self.assertRaises(SystemExit) as cm:
                _check_pywin32()

            self.assertEqual(cm.exception.code, 1)
            # Verify error message was printed
            self.assertTrue(self.mock_stderr.write.called)
            # Check for key phrase in output
            output = "".join(call.args[0] for call in self.mock_stderr.write.call_args_list)
            self.assertIn("pywin32 is required but not installed", output)

    def test_check_pywin32_import_error(self):
        """Test pywin32 check when ImportError is raised during check"""
        with patch("importlib.util.find_spec") as mock_find_spec:
            mock_find_spec.side_effect = ImportError("Test Import Error")

            with self.assertRaises(SystemExit) as cm:
                _check_pywin32()

            self.assertEqual(cm.exception.code, 1)
            output = "".join(call.args[0] for call in self.mock_stderr.write.call_args_list)
            self.assertIn("pywin32 is required but not installed", output)

    def test_check_pywin32_value_error(self):
        """Test pywin32 check when ValueError is raised during check"""
        with patch("importlib.util.find_spec") as mock_find_spec:
            mock_find_spec.side_effect = ValueError("Test Value Error")

            with self.assertRaises(SystemExit) as cm:
                _check_pywin32()

            self.assertEqual(cm.exception.code, 1)
            output = "".join(call.args[0] for call in self.mock_stderr.write.call_args_list)
            self.assertIn("pywin32 is required but not installed", output)


class TestCLI(unittest.TestCase):
    def setUp(self):
        # Patch sys.stdout and sys.stderr to capture output
        self.stdout_patch = patch("sys.stdout")
        self.stderr_patch = patch("sys.stderr")
        self.mock_stdout = self.stdout_patch.start()
        self.mock_stderr = self.stderr_patch.start()

        # Patch sys.argv - default to just the script name
        self.argv_patch = patch("sys.argv", ["mailtool"])
        self.mock_argv = self.argv_patch.start()

        # Patch sys.platform to be "win32"
        self.platform_patch = patch("sys.platform", "win32")
        self.mock_platform = self.platform_patch.start()

        # Patch _check_pywin32 (we assume pywin32 check passes)
        # We need to patch where it is defined in mailtool.cli
        self.check_pywin32_patch = patch("mailtool.cli._check_pywin32")
        self.mock_check_pywin32 = self.check_pywin32_patch.start()

        # Patch OutlookBridge
        # Since mailtool.cli imports it inside main, we need to patch
        # mailtool.bridge.OutlookBridge so that when it is imported, it gets the mock.
        # But we also need to ensure mailtool.bridge is importable (handled by sys.modules mocks above)
        self.bridge_patch = patch("mailtool.bridge.OutlookBridge")
        self.mock_bridge_cls = self.bridge_patch.start()
        self.mock_bridge = self.mock_bridge_cls.return_value

    def tearDown(self):
        self.stdout_patch.stop()
        self.stderr_patch.stop()
        self.argv_patch.stop()
        self.platform_patch.stop()
        self.check_pywin32_patch.stop()
        self.bridge_patch.stop()

    def test_main_no_args(self):
        """Test main with no arguments (should print help and exit)"""
        # When no args are provided (except script name), argparse prints help and exits
        with patch("sys.argv", ["mailtool"]):
            with self.assertRaises(SystemExit) as cm:  # noqa: PT027
                main()
            self.assertEqual(cm.exception.code, 1)  # noqa: PT009

    def test_check_platform_non_windows(self):
        """Test _check_platform on non-Windows OS"""
        with patch("sys.platform", "linux"):
            # Capture stderr to verify error message
            with patch("sys.stderr") as mock_stderr:
                with self.assertRaises(SystemExit) as cm:
                    _check_platform()
                self.assertEqual(cm.exception.code, 1)

                # Verify error message was printed
                # mock_stderr.write is called multiple times
                calls = mock_stderr.write.call_args_list
                # Combine all calls to check the full message
                full_message = "".join(call[0][0] for call in calls)
                self.assertIn("Error: mailtool requires Windows", full_message)

    def test_check_platform_windows(self):
        """Test _check_platform on Windows OS (should pass)"""
        # setUp already patches sys.platform to "win32", but we can be explicit
        with patch("sys.platform", "win32"):
            # Should not raise any exception
            _check_platform()

    def test_emails_command(self):
        """Test emails command"""
        with patch("sys.argv", ["mailtool", "emails", "--limit", "5"]):
            self.mock_bridge.list_emails.return_value = [{"subject": "Test Email"}]
            main()
            self.mock_bridge.list_emails.assert_called_with(limit=5, folder="Inbox")

    def test_calendar_command(self):
        """Test calendar command"""
        with patch("sys.argv", ["mailtool", "calendar", "--days", "3"]):
            self.mock_bridge.list_calendar_events.return_value = [{"subject": "Meeting"}]
            main()
            self.mock_bridge.list_calendar_events.assert_called_with(days=3, all_events=False)

    def test_tasks_command(self):
        """Test tasks command"""
        with patch("sys.argv", ["mailtool", "tasks"]):
            self.mock_bridge.list_tasks.return_value = [{"subject": "Task 1"}]
            main()
            self.mock_bridge.list_tasks.assert_called()

    def test_send_email(self):
        """Test send email command"""
        with patch("sys.argv", ["mailtool", "send", "--to", "test@example.com", "--subject", "Hello", "--body", "World"]):
            self.mock_bridge.send_email.return_value = True
            main()
            self.mock_bridge.send_email.assert_called_with(
                "test@example.com", "Hello", "World", None, None, html_body=None, file_paths=None, save_draft=False
            )

    def test_create_appt(self):
        """Test create appointment command"""
        with patch("sys.argv", ["mailtool", "create-appt", "--subject", "Meeting", "--start", "2023-01-01 10:00:00", "--end", "2023-01-01 11:00:00"]):
            self.mock_bridge.create_appointment.return_value = "EntryID123"
            main()
            self.mock_bridge.create_appointment.assert_called()

    def test_create_task(self):
        """Test create task command"""
        with patch("sys.argv", ["mailtool", "create-task", "--subject", "My Task"]):
            self.mock_bridge.create_task.return_value = "EntryID123"
            main()
            self.mock_bridge.create_task.assert_called()

    def test_mcp_command(self):
        """Test mcp command"""
        with patch("sys.argv", ["mailtool", "mcp"]):
            # Since we mocked sys.modules["mailtool.mcp.server"], we need to set up the return value
            # of main in that mock
            mock_module = sys.modules["mailtool.mcp.server"]
            main()
            mock_module.main.assert_called_with(default_account=None)
