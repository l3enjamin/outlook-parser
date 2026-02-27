#!/usr/bin/env python3
"""
Outlook COM Automation Bridge
Runs on Windows, callable from WSL2
Provides CLI interface to Outlook email and calendar via COM

Requirements (Windows):
    pip install pywin32

Usage:
    # List recent emails
    python outlook_com_bridge.py emails --limit 10

    # List calendar events
    python outlook_com_bridge.py calendar --days 7

    # Get email body
    python outlook_com_bridge.py email --id <entry_id>
"""

# Modified to test pre-commit hook

import contextlib
import logging
import sys
from datetime import datetime, timedelta

import win32com.client

logger = logging.getLogger(__name__)

# MAPI Property Tags
# Source: https://github.com/microsoft/microsoft-graph-docs/blob/main/api-reference/v1.0/resources/opentypeextension.md (and other MAPI docs)
# We use the DASL property tag format for Table.Columns.Add
PR_ENTRYID = "http://schemas.microsoft.com/mapi/proptag/0x0FFF0102"
PR_SUBJECT = "http://schemas.microsoft.com/mapi/proptag/0x0037001E"
PR_SENDER_NAME = "http://schemas.microsoft.com/mapi/proptag/0x0C1A001E"
PR_SENT_REPRESENTING_NAME = "http://schemas.microsoft.com/mapi/proptag/0x0042001E"
PR_SENDER_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x5D01001E"
PR_MESSAGE_DELIVERY_TIME = "http://schemas.microsoft.com/mapi/proptag/0x0E060040"
PR_HASATTACH = "http://schemas.microsoft.com/mapi/proptag/0x0E1B000B"
PR_UNREAD = "http://schemas.microsoft.com/mapi/proptag/0x0E09000B"
PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"


class OutlookBridge:
    """Bridge to Outlook application via COM"""

    @staticmethod
    def _safe_get_attr(obj, attr, default=None):
        """
        Safely get an attribute from a COM object, handling COM errors gracefully

        Args:
            obj: COM object
            attr: Attribute name to get
            default: Default value if attribute access fails

        Returns:
            Attribute value or default
        """
        try:
            return getattr(obj, attr, default)
        # Catch pywintypes.com_error if available, otherwise fall back to Exception
        except Exception:
            return default

    def _get_emails_from_table(self, table, limit):
        """
        Efficiently extract email data from a Table object

        Args:
            table: Outlook Table object
            limit: Maximum number of emails to return

        Returns:
            List of email dictionaries
        """
        try:
            # Add required columns
            # Note: Columns MUST be added in the order we access them
            table.Columns.Add(PR_ENTRYID)
            table.Columns.Add(PR_SUBJECT)
            table.Columns.Add(PR_SENT_REPRESENTING_NAME)
            table.Columns.Add(PR_SENDER_SMTP_ADDRESS)
            table.Columns.Add(PR_MESSAGE_DELIVERY_TIME)
            table.Columns.Add(PR_UNREAD)
            table.Columns.Add(PR_HASATTACH)
        except Exception as e:
            logger.error(f"Error adding columns to table: {e}")
            return []

        emails = []

        while not table.EndOfTable:
            if len(emails) >= limit:
                break

            remaining = limit - len(emails)
            batch_size = min(remaining, 50)

            try:
                # Fetch a batch of rows
                rows = table.GetArray(batch_size)
            except Exception as e:
                 logger.error(f"Error getting array from table: {e}")
                 break

            if not rows:
                break

            for row in rows:
                try:
                    # Map row values to our dictionary structure
                    entry_id_raw = row[0]
                    if isinstance(entry_id_raw, (bytes, bytearray)):
                        entry_id = entry_id_raw.hex().upper()
                    else:
                        entry_id = str(entry_id_raw)

                    subject = row[1] if row[1] else "(No Subject)"
                    sender_name = row[2] if row[2] else ""
                    sender_smtp = row[3] if row[3] else ""

                    received_time = row[4]
                    formatted_time = None
                    if received_time:
                        if hasattr(received_time, "strftime"):
                             formatted_time = received_time.strftime("%Y-%m-%d %H:%M:%S")
                        else:
                             formatted_time = str(received_time)

                    is_unread = bool(row[5])
                    has_attachments = bool(row[6])

                    email = {
                        "entry_id": entry_id,
                        "subject": subject,
                        "sender_name": sender_name,
                        "sender": sender_smtp,
                        "received_time": formatted_time,
                        "unread": is_unread,
                        "has_attachments": has_attachments,
                    }
                    emails.append(email)
                except Exception as e:
                    logger.warning(f"Error parsing email row from table: {e}")
                    continue

        return emails

    def __init__(self, default_account: str | None = None):
        """
        Connect to running Outlook instance or start it

        Launch Logic:
        1. Try GetActiveObject first (if Outlook is running)
        2. Fall back to Dispatch if Outlook is closed
        """
        try:
            self.outlook = win32com.client.GetActiveObject("Outlook.Application")
        except Exception as e:
            # Outlook might not be running - try to launch it
            try:
                self.outlook = win32com.client.Dispatch("Outlook.Application")
            except Exception:
                logger.error("Could not connect to or launch Outlook.")
                logger.error(f"Details: {e}")
                logger.error(
                    "Hint: Make sure Outlook is installed and you can launch it manually."
                )
                logger.exception("Full traceback:")
                sys.exit(1)

        self.namespace = self.outlook.GetNamespace("MAPI")
        # Default account name and root folder (set by set_default_account or via init param)
        self.default_account_name = None
        self.default_root_folder = None

        # If provided, attempt to set the default account/store on init
        if default_account:
            with contextlib.suppress(Exception):
                self.set_default_account(default_account)

    # -- Helper methods for account/folder resolution -----------------
    def _find_root_by_name(self, acc_name: str):
        """Find and return the root folder object for an account by name (case-insensitive).

        Returns None if not found.
        """
        if not acc_name:
            return None
        try:
            count = self.namespace.Folders.Count
        except Exception:
            count = None

        if count and count > 0:
            for i in range(1, count + 1):
                try:
                    root = self.namespace.Folders.Item(i)
                    if str(root.Name).strip().lower() == str(acc_name).strip().lower():
                        return root
                except Exception:
                    continue

        # Fallback: try a reasonable range if Count isn't available
        for i in range(1, 10):
            try:
                root = self.namespace.Folders.Item(i)
                if str(root.Name).strip().lower() == str(acc_name).strip().lower():
                    return root
            except Exception:
                continue

        return None

    def _get_root(self):
        """Return the active root folder to use (default account root if set, else the first mailbox/root)."""
        if self.default_root_folder:
            return self.default_root_folder
        # Try DefaultStore if set
        try:
            default_store = getattr(self.namespace, "DefaultStore", None)
            if default_store:
                try:
                    root = default_store.GetRootFolder()
                    return root
                except Exception:
                    pass
        except Exception:
            pass

        # Fallback: first available root
        try:
            return self.namespace.Folders.Item(1)
        except Exception:
            # try a small range as a last resort
            for i in range(1, 10):
                try:
                    return self.namespace.Folders.Item(i)
                except Exception:
                    continue
        return None

    def _find_account_by_name(self, name: str):
        """Find an Outlook Account object by SMTP address or display name (case-insensitive)."""
        if not name:
            return None
        try:
            accounts = self.namespace.Accounts
            for acc in accounts:
                try:
                    if (
                        hasattr(acc, "SmtpAddress")
                        and str(acc.SmtpAddress).strip().lower()
                        == str(name).strip().lower()
                    ):
                        return acc
                    if (
                        hasattr(acc, "DisplayName")
                        and str(acc.DisplayName).strip().lower()
                        == str(name).strip().lower()
                    ):
                        return acc
                except Exception:
                    continue
        except Exception:
            pass
        return None

    def set_default_account(self, acc_name: str):
        """
        Set the default account by name

        Args:
            acc_name: Account name to set as default

        Returns:
            True if successful, False otherwise
        """
        root = self._find_root_by_name(acc_name)
        if not root:
            return False
        try:
            # Set attributes for bridge usage
            self.default_account_name = acc_name
            self.default_root_folder = root
            # Also set DefaultStore to help other COM calls that rely on it
            with contextlib.suppress(Exception):
                self.namespace.DefaultStore = root.Store
            return True
        except Exception:
            return False

    def get_inbox(self):
        """Get the inbox folder"""
        # Prefer default account root when set
        root = self._get_root()
        if root:
            try:
                # Try case-sensitive first
                return root.Folders["Inbox"]
            except Exception:
                # Try case-insensitive search across root subfolders
                try:
                    for f in root.Folders:
                        if str(f.Name).strip().lower() == "inbox":
                            return f
                except Exception:
                    pass

        # Fallback to namespace default
        try:
            return self.namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
        except Exception:
            return None

    def list_folders(self, acc_name: str | None = None) -> dict[str, list[dict]]:
        """
        Recursively list all folders for all accounts.
        Args:
            acc_name: Specific account name to list folders for (optional)
        Returns:
            Dict with account names as keys and list of folder info dicts as values

        """

        def retrieve_folder_details(folder, parent_folder, depth):
            logger.info(f"{'  ' * depth}- {folder.Name} (Items: {folder.Items.Count})")
            all_items = []
            cur_folder_data = {
                "name": folder.Name,
                "id": folder.EntryID,
                "parent_name": parent_folder.Name if parent_folder else None,
                "parent_id": parent_folder.EntryID if parent_folder else None,
                "number_of_items": folder.Items.Count,
                "path": folder.FolderPath,
                "depth": depth,
                "account": parent_folder.Name if parent_folder else None,
            }
            all_items.append(cur_folder_data)
            for subfolder in folder.Folders:
                all_items.extend(retrieve_folder_details(subfolder, folder, depth + 1))
            return all_items

        final = {}
        for i in range(1, 7):
            try:
                parent_folder = self.namespace.Folders.Item(i)
                if acc_name and parent_folder.Name != acc_name:
                    logger.info(
                        f"acc_name arg does not match found account: {parent_folder.Name}\n  skipping..."
                    )
                    continue
                logger.info(f"Account: {parent_folder.Name}")
            except Exception:
                logger.info(f"Finished listing accounts. Total accounts: {i - 1}")
                break  # No more accounts

            final[parent_folder.Name] = retrieve_folder_details(parent_folder, None, 0)

        return final

    def get_calendar(self):
        """Get the calendar folder"""
        # Prefer default account root when set
        root = self._get_root()
        if root:
            try:
                # direct access
                return root.Folders["Calendar"]
            except Exception:
                # case-insensitive search
                try:
                    for f in root.Folders:
                        if str(f.Name).strip().lower() == "calendar":
                            return f
                except Exception:
                    pass

        # Fallback: search all accounts by name
        try:
            count = self.namespace.Folders.Count
        except Exception:
            count = None

        if count and count > 0:
            for i in range(1, count + 1):
                try:
                    parent_folder = self.namespace.Folders.Item(i)
                    try:
                        cal = parent_folder.Folders["Calendar"]
                        return cal
                    except Exception:
                        # try case-insensitive
                        for f in parent_folder.Folders:
                            if str(f.Name).strip().lower() == "calendar":
                                return f
                except Exception:
                    continue

        return None

    def get_tasks(self):
        """Get the tasks folder"""
        root = self._get_root()
        if root:
            try:
                return root.Folders["Tasks"]
            except Exception:
                try:
                    for f in root.Folders:
                        if str(f.Name).strip().lower() == "tasks":
                            return f
                except Exception:
                    pass

        try:
            return self.namespace.GetDefaultFolder(13)  # 13 = olFolderTasks
        except Exception:
            return None

    def get_folder_by_name(self, folder_name):
        """
        Get a folder by name (e.g., "Sent Items", "Archive", etc.)

        Args:
            folder_name: Name of the folder

        Returns:
            Folder object or None
        """
        # Try default account root first
        if not folder_name:
            return None

        root = self._get_root()
        if root:
            try:
                return root.Folders[folder_name]
            except Exception:
                try:
                    for f in root.Folders:
                        if (
                            str(f.Name).strip().lower()
                            == str(folder_name).strip().lower()
                        ):
                            return f
                except Exception:
                    pass

        # Search across all account roots
        try:
            count = self.namespace.Folders.Count
        except Exception:
            count = None

        if count and count > 0:
            for i in range(1, count + 1):
                try:
                    parent = self.namespace.Folders.Item(i)
                    try:
                        return parent.Folders[folder_name]
                    except Exception:
                        # case-insensitive search in this parent
                        try:
                            for f in parent.Folders:
                                if (
                                    str(f.Name).strip().lower()
                                    == str(folder_name).strip().lower()
                                ):
                                    return f
                        except Exception:
                            pass
                except Exception:
                    continue

        # Last resort: try the first root's children
        try:
            root = self.namespace.Folders.Item(1)
            try:
                return root.Folders[folder_name]
            except Exception:
                for f in root.Folders:
                    if str(f.Name).strip().lower() == str(folder_name).strip().lower():
                        return f
        except Exception:
            pass

        return None

    def get_item_by_id(self, entry_id):
        """
        Get any Outlook item by EntryID (O(1) direct access)

        Args:
            entry_id: Outlook EntryID

        Returns:
            Outlook item (MailItem, AppointmentItem, TaskItem, etc.) or None
        """
        try:
            return self.namespace.GetItemFromID(entry_id)
        except Exception:
            return None

    def resolve_smtp_address(self, mail_item):
        """
        Get SMTP address from Exchange address (EX type)

        Args:
            mail_item: Outlook MailItem

        Returns:
            SMTP email address string
        """
        try:
            if (
                (
                    hasattr(mail_item, "SenderEmailType")
                    and mail_item.SenderEmailType == "EX"
                )
                and hasattr(mail_item, "Sender")
                and hasattr(mail_item.Sender, "GetExchangeUser")
            ):
                exchange_user = mail_item.Sender.GetExchangeUser()
                if hasattr(exchange_user, "PrimarySmtpAddress"):
                    return exchange_user.PrimarySmtpAddress
            return (
                mail_item.SenderEmailAddress
                if hasattr(mail_item, "SenderEmailAddress")
                else ""
            )
        except Exception:
            return (
                mail_item.SenderEmailAddress
                if hasattr(mail_item, "SenderEmailAddress")
                else ""
            )
