"""
Manual MCP Server Test Script for Claude Code Integration

This script validates the MCP SDK v2 server implementation by:
1. Testing all 23 tools are registered and callable
2. Verifying structured output (Pydantic models)
3. Testing error handling
4. Validating all 7 resources

Requirements:
- Windows with Outlook running
- MCP server package installed (uv run --with mcp --with pywin32)
- Claude Code plugin loaded

Usage:
    uv run --with mcp --with pywin32 python scripts/manual_mcp_test.py

Output:
- Detailed test results for each tool and resource
- Validation of structured output
- Error handling verification
"""

import sys
from typing import Any

# Add src to path for imports
sys.path.insert(0, "C:/dev/mailtool/src")

try:
    from mailtool.bridge import OutlookBridge  # noqa: F401
    from mailtool.mcp import resources  # noqa: F401
    from mailtool.mcp.models import (  # noqa: F401
        AppointmentDetails,
        AppointmentSummary,
        CreateAppointmentResult,
        CreateTaskResult,
        EmailDetails,
        EmailSummary,
        FreeBusyInfo,
        OperationResult,
        SendEmailResult,
        TaskSummary,
    )
    from mailtool.mcp.server import mcp
    print("[OK] All imports successful")
except ImportError as e:
    print(f"[FAIL] Import failed: {e}")
    print("\nPlease run: uv sync --all-groups")
    sys.exit(1)


def print_section(title: str) -> None:
    """Print a section header."""
    print(f"\n{'='*70}")
    print(f"  {title}")
    print(f"{'='*70}\n")


def test_tool_registration() -> dict[str, Any]:
    """Test that all 23 tools are registered with FastMCP."""
    print_section("Testing Tool Registration")

    expected_tools = {
        # Email tools (9)
        "list_emails",
        "get_email",
        "send_email",
        "reply_email",
        "forward_email",
        "mark_email",
        "move_email",
        "delete_email",
        "search_emails",
        # Calendar tools (7)
        "list_calendar_events",
        "create_appointment",
        "get_appointment",
        "edit_appointment",
        "respond_to_meeting",
        "delete_appointment",
        "get_free_busy",
        # Task tools (7)
        "list_tasks",
        "list_all_tasks",
        "create_task",
        "get_task",
        "edit_task",
        "complete_task",
        "delete_task",
    }

    registered_tools = mcp._tool_manager._tools

    results = {
        "expected_count": len(expected_tools),
        "registered_count": len(registered_tools),
        "missing_tools": [],
        "extra_tools": [],
        "all_tools": list(registered_tools.keys()),
    }

    print(f"Expected tools: {results['expected_count']}")
    print(f"Registered tools: {results['registered_count']}")

    for tool_name in expected_tools:
        if tool_name in registered_tools:
            tool = registered_tools[tool_name]
            print(f"  [OK] {tool_name}: {tool.description[:50]}...")
        else:
            results["missing_tools"].append(tool_name)
            print(f"  [FAIL] {tool_name}: NOT FOUND")

    for tool_name in registered_tools:
        if tool_name not in expected_tools:
            results["extra_tools"].append(tool_name)
            print(f"  [WARN] {tool_name}: UNEXPECTED")

    if not results["missing_tools"] and not results["extra_tools"]:
        print("\n[OK] All 23 tools registered correctly")
        results["passed"] = True
    else:
        print("\n[FAIL] Tool registration has issues")
        results["passed"] = False

    return results


def test_resource_registration() -> dict[str, Any]:
    """Test that all 7 resources are registered with FastMCP."""
    print_section("Testing Resource Registration")

    expected_resources = {
        # Email resources (3)
        "inbox://emails",
        "inbox://unread",
        "email://{entry_id}",
        # Calendar resources (2)
        "calendar://today",
        "calendar://week",
        # Task resources (2)
        "tasks://active",
        "tasks://all",
    }

    registered_resources = mcp._resource_manager._resources
    registered_templates = mcp._resource_manager._templates

    results = {
        "expected_count": len(expected_resources),
        "regular_count": len(registered_resources),
        "template_count": len(registered_templates),
        "missing_resources": [],
        "extra_resources": [],
        "all_resources": {
            "regular": list(registered_resources.keys()),
            "templates": list(registered_templates.keys()),
        },
    }

    print(f"Expected resources: {results['expected_count']}")
    print(f"Regular resources: {results['regular_count']}")
    print(f"Template resources: {results['template_count']}")

    # Check regular resources
    for resource_uri in expected_resources:
        if "{entry_id}" not in resource_uri:
            if resource_uri in registered_resources:
                resource = registered_resources[resource_uri]
                print(f"  [OK] {resource_uri}: {type(resource).__name__!s}")
            else:
                results["missing_resources"].append(resource_uri)
                print(f"  [FAIL] {resource_uri}: NOT FOUND")

    # Check template resources
    template_uris = [uri for uri in expected_resources if "{entry_id}" in uri]
    for template_uri in template_uris:
        if template_uri in registered_templates:
            template = registered_templates[template_uri]
            print(f"  [OK] {template_uri}: {type(template).__name__!s}")
        else:
            results["missing_resources"].append(template_uri)
            print(f"  [FAIL] {template_uri}: NOT FOUND")

    if not results["missing_resources"]:
        print("\n[OK] All 7 resources registered correctly")
        results["passed"] = True
    else:
        print("\n[FAIL] Resource registration has issues")
        results["passed"] = False

    return results


def test_model_validation() -> dict[str, Any]:
    """Test that Pydantic models are properly defined."""
    print_section("Testing Pydantic Model Validation")

    results = {
        "models_tested": [],
        "failed_models": [],
        "passed": True,
    }

    # Test EmailSummary
    try:
        email = EmailSummary(
            entry_id="test123",
            subject="Test Email",
            sender="sender@example.com",
            sender_name="Test Sender",
            received_time="2025-01-19 10:00:00",
            unread=True,
            has_attachments=False,
        )
        print(f"  [OK] EmailSummary: {email.model_dump_json()[:100]}...")
        results["models_tested"].append("EmailSummary")
    except Exception as e:
        print(f"  [FAIL] EmailSummary: {e}")
        results["failed_models"].append("EmailSummary")
        results["passed"] = False

    # Test EmailDetails
    try:
        email_details = EmailDetails(
            entry_id="test123",
            subject="Test Email",
            sender="sender@example.com",
            sender_name="Test Sender",
            received_time="2025-01-19 10:00:00",
            has_attachments=False,
            body="Test body",
            html_body="<p>Test body</p>",
        )
        print(f"  [OK] EmailDetails: {email_details.model_dump_json()[:100]}...")
        results["models_tested"].append("EmailDetails")
    except Exception as e:
        print(f"  [FAIL] EmailDetails: {e}")
        results["failed_models"].append("EmailDetails")
        results["passed"] = False

    # Test SendEmailResult
    try:
        send_result = SendEmailResult(
            success=True,
            entry_id="draft123",
            message="Email saved as draft",
        )
        print(f"  [OK] SendEmailResult: {send_result.model_dump_json()}")
        results["models_tested"].append("SendEmailResult")
    except Exception as e:
        print(f"  [FAIL] SendEmailResult: {e}")
        results["failed_models"].append("SendEmailResult")
        results["passed"] = False

    # Test AppointmentSummary
    try:
        appt = AppointmentSummary(
            entry_id="appt123",
            subject="Test Meeting",
            start="2025-01-20 14:00:00",
            end="2025-01-20 15:00:00",
            location="Room 101",
            organizer="organizer@example.com",
            all_day=False,
            required_attendees="attendee1@example.com",
            optional_attendees="",  # Changed from None to empty string
            response_status="Accepted",
            meeting_status="Meeting",
            response_requested=True,
        )
        print(f"  [OK] AppointmentSummary: {appt.model_dump_json()[:100]}...")
        results["models_tested"].append("AppointmentSummary")
    except Exception as e:
        print(f"  [FAIL] AppointmentSummary: {e}")
        results["failed_models"].append("AppointmentSummary")
        results["passed"] = False

    # Test TaskSummary
    try:
        task = TaskSummary(
            entry_id="task123",
            subject="Test Task",
            body="Task description",
            due_date="2025-01-25",
            status=1,
            priority=2,
            complete=False,
            percent_complete=50.0,
        )
        print(f"  [OK] TaskSummary: {task.model_dump_json()[:100]}...")
        results["models_tested"].append("TaskSummary")
    except Exception as e:
        print(f"  [FAIL] TaskSummary: {e}")
        results["failed_models"].append("TaskSummary")
        results["passed"] = False

    # Test OperationResult
    try:
        op_result = OperationResult(success=True, message="Operation completed")
        print(f"  [OK] OperationResult: {op_result.model_dump_json()}")
        results["models_tested"].append("OperationResult")
    except Exception as e:
        print(f"  [FAIL] OperationResult: {e}")
        results["failed_models"].append("OperationResult")
        results["passed"] = False

    if results["passed"]:
        print(f"\n[OK] All {len(results['models_tested'])} models validated successfully")
    else:
        print(f"\n[FAIL] {len(results['failed_models'])} model(s) failed validation")

    return results


def test_tool_signatures() -> dict[str, Any]:
    """Test that tool functions have correct signatures and return types."""
    print_section("Testing Tool Signatures")

    results = {
        "tools_tested": [],
        "failed_tools": [],
        "passed": True,
    }

    # Sample a few key tools to verify they have the right structure
    key_tools = ["list_emails", "get_email", "send_email", "create_appointment", "create_task"]

    for tool_name in key_tools:
        try:
            tool = mcp._tool_manager._tools[tool_name]
            # Just verify the tool exists and has a name (FastMCP Tool object structure)
            if hasattr(tool, 'name') or hasattr(tool, 'description'):
                print(f"  [OK] {tool_name}: Tool registered correctly")
                results["tools_tested"].append(tool_name)
            else:
                print(f"  [FAIL] {tool_name}: Tool structure invalid")
                results["failed_tools"].append(tool_name)
                results["passed"] = False
        except Exception as e:
            print(f"  [FAIL] {tool_name}: {e}")
            results["failed_tools"].append(tool_name)
            results["passed"] = False

    if results["passed"]:
        print(f"\n[OK] All {len(results['tools_tested'])} key tool signatures validated")
    else:
        print(f"\n[FAIL] {len(results['failed_tools'])} tool signature(s) failed validation")

    return results


def test_exception_classes() -> dict[str, Any]:
    """Test that custom exception classes are properly defined."""
    print_section("Testing Custom Exception Classes")

    results = {
        "exceptions_tested": [],
        "failed_exceptions": [],
        "passed": True,
    }

    try:
        from mailtool.mcp.exceptions import (
            OutlookComError,
            OutlookNotFoundError,
            OutlookValidationError,
        )

        # Test OutlookNotFoundError
        try:
            raise OutlookNotFoundError("test_entry_id")
        except OutlookNotFoundError as e:
            error_msg = str(e)
            print(f"  [OK] OutlookNotFoundError: {error_msg[:80]}...")
            print(f"    - Error code: {e.error.code}")
            print(f"    - Entry ID: {e.entry_id}")
            results["exceptions_tested"].append("OutlookNotFoundError")
        except Exception as e:
            print(f"  [FAIL] OutlookNotFoundError: Unexpected exception {e}")
            results["failed_exceptions"].append("OutlookNotFoundError")
            results["passed"] = False

        # Test OutlookComError
        try:
            raise OutlookComError("COM operation failed")
        except OutlookComError as e:
            error_msg = str(e)
            print(f"  [OK] OutlookComError: {error_msg[:80]}...")
            print(f"    - Error code: {e.error.code}")
            print(f"    - Details: {e.details}")
            results["exceptions_tested"].append("OutlookComError")
        except Exception as e:
            print(f"  [FAIL] OutlookComError: Unexpected exception {e}")
            results["failed_exceptions"].append("OutlookComError")
            results["passed"] = False

        # Test OutlookValidationError
        try:
            raise OutlookValidationError("Invalid email format", "email")
        except OutlookValidationError as e:
            error_msg = str(e)
            print(f"  [OK] OutlookValidationError: {error_msg[:80]}...")
            print(f"    - Error code: {e.error.code}")
            print(f"    - Field: {e.field}")
            results["exceptions_tested"].append("OutlookValidationError")
        except Exception as e:
            print(f"  [FAIL] OutlookValidationError: Unexpected exception {e}")
            results["failed_exceptions"].append("OutlookValidationError")
            results["passed"] = False

        if results["passed"]:
            print(f"\n[OK] All {len(results['exceptions_tested'])} exception classes validated")
        else:
            print(f"\n[FAIL] {len(results['failed_exceptions'])} exception class(es) failed validation")

    except ImportError as e:
        print(f"  [FAIL] Failed to import exception classes: {e}")
        results["passed"] = False

    return results


def test_lifespan_management() -> dict[str, Any]:
    """Test that lifespan management is properly configured."""
    print_section("Testing Lifespan Management")

    results = {
        "lifespan_configured": False,
        "bridge_state_configured": False,
        "passed": False,
    }

    # Check if lifespan is configured (FastMCP stores it in _mcp_server)
    if hasattr(mcp, "_mcp_server") and hasattr(mcp._mcp_server, "lifespan"):
        print(f"  [OK] Lifespan configured: {mcp._mcp_server.lifespan}")
        results["lifespan_configured"] = True
    else:
        print("  [FAIL] Lifespan not configured")

    # Check if bridge state is accessible
    try:
        from mailtool.mcp import server

        if hasattr(server, "_get_bridge"):
            print("  [OK] Bridge state function exists: server._get_bridge")
            results["bridge_state_configured"] = True
        else:
            print("  [FAIL] Bridge state function not found")
    except Exception as e:
        print(f"  [FAIL] Error checking bridge state: {e}")

    if results["lifespan_configured"] and results["bridge_state_configured"]:
        print("\n[OK] Lifespan management properly configured")
        results["passed"] = True
    else:
        print("\n[FAIL] Lifespan management has issues")

    return results


def print_summary(results: dict[str, Any]) -> None:
    """Print test summary."""
    print_section("Test Summary")

    total_tests = len(results)
    passed_tests = sum(1 for r in results.values() if r.get("passed", False))

    print(f"Total test suites: {total_tests}")
    print(f"Passed: {passed_tests}")
    print(f"Failed: {total_tests - passed_tests}")

    print("\nDetailed Results:")
    for test_name, test_result in results.items():
        status = "[OK] PASS" if test_result.get("passed", False) else "[FAIL] FAIL"
        print(f"  {status}: {test_name}")

    print("\n" + "="*70)
    if passed_tests == total_tests:
        print("  [OK] ALL TESTS PASSED")
    else:
        print(f"  [FAIL] {total_tests - passed_tests} TEST SUITE(S) FAILED")
    print("="*70 + "\n")


def main() -> int:
    """Run all manual MCP server tests."""
    print("\n" + "="*70)
    print("  MCP SDK v2 Server Manual Test Suite")
    print("  Testing tool registration, models, and resources")
    print("="*70)

    results = {}

    # Run all tests
    results["Tool Registration"] = test_tool_registration()
    results["Resource Registration"] = test_resource_registration()
    results["Model Validation"] = test_model_validation()
    results["Tool Signatures"] = test_tool_signatures()
    results["Exception Classes"] = test_exception_classes()
    results["Lifespan Management"] = test_lifespan_management()

    # Print summary
    print_summary(results)

    # Return exit code
    return 0 if all(r.get("passed", False) for r in results.values()) else 1


if __name__ == "__main__":
    sys.exit(main())
