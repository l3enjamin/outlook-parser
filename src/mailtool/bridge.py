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

    def list_emails(self, limit=10, folder="Inbox"):
        """
        List emails from the specified folder

        Args:
            limit: Maximum number of emails to return
            folder: Folder name (default: Inbox)

        Returns:
            List of email dictionaries
        """
        # Use get_inbox() for the default Inbox to ensure correct account
        if folder == "Inbox":
            inbox = self.get_inbox()
        else:
            inbox = self.get_folder_by_name(folder)
            if not inbox:
                inbox = self.get_inbox()

        items = inbox.Items

        # Sort by received time, most recent first
        items.Sort("[ReceivedTime]", True)

        emails = []
        count = 0
        for item in items:
            if count >= limit:
                break

            try:
                email = {
                    "entry_id": item.EntryID,
                    "subject": item.Subject,
                    "sender": self.resolve_smtp_address(item),
                    "sender_name": item.SenderName,
                    "received_time": item.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
                    if item.ReceivedTime
                    else None,
                    "unread": item.Unread,
                    "has_attachments": item.Attachments.Count > 0,
                }
                emails.append(email)
                count += 1
            except Exception:
                # Skip items that can't be accessed
                continue

        return emails

    def get_email_body(self, entry_id):
        """
        Get the full body of an email by entry ID (O(1) direct access)

        Args:
            entry_id: Outlook EntryID of the email

        Returns:
            Email dictionary with body
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                return {
                    "entry_id": item.EntryID,
                    "subject": item.Subject,
                    "sender": self.resolve_smtp_address(item),
                    "sender_name": item.SenderName,
                    "body": item.Body,
                    "html_body": item.HTMLBody,
                    "received_time": item.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
                    if item.ReceivedTime
                    else None,
                    "has_attachments": item.Attachments.Count > 0,
                }
            except Exception:
                return None
        return None

    def get_email_parsed(
        self,
        entry_id,
        remove_quoted=False,
        deduplication_tier="none",
        strip_html=True,
    ):
        """
        Get structured email object using mail-parser (O(1) direct access)

        Args:
            entry_id: Outlook EntryID of the email
            remove_quoted: DEPRECATED - use deduplication_tier="low" instead.
                           If True, sets deduplication_tier="low".
            deduplication_tier: Strategy for deduplicating quoted content.
                                "none" - No deduplication (default)
                                "low" - Strip using mail-parser-reply (standard)
                                "medium" - Strip ONLY if parent email found via References/In-Reply-To
                                "high" - Strip ONLY if parent email found via Subject/Content (Not Implemented)
            strip_html: If True, remove HTML code from body and clear text_html field (default: True)

        Returns:
            Dictionary matching EmailParsed model structure
        """
        # Backwards compatibility for remove_quoted bool
        if remove_quoted and deduplication_tier == "none":
            deduplication_tier = "low"

        import os
        import tempfile

        try:
            import mailparser
        except ImportError:
            logger.warning(
                "mail-parser not installed, falling back to basic parsing"
            )
            item = self.get_item_by_id(entry_id)
            if not item:
                return None
            return self._fallback_parsed_model(
                item, deduplication_tier, strip_html=strip_html
            )

        item = self.get_item_by_id(entry_id)
        if not item:
            return None

        # Create temp file
        fd, temp_path = tempfile.mkstemp(suffix=".msg")
        os.close(fd)

        try:
            # Save as MSG
            # olMSG = 3
            try:
                item.SaveAs(temp_path, 3)
            except Exception as e:
                logger.error(f"Error saving to .msg: {e}")
                return self._fallback_parsed_model(
                    item, deduplication_tier, strip_html=strip_html
                )

            # Parse
            try:
                mail = mailparser.parse_from_file_msg(temp_path)
                return self._convert_to_parsed_model(
                    mail, item, deduplication_tier, strip_html=strip_html
                )
            except Exception as e:
                logger.error(f"Error parsing .msg with mail-parser: {e}")
                return self._fallback_parsed_model(
                    item, deduplication_tier, strip_html=strip_html
                )

        finally:
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except Exception:
                    pass

    def _check_parent_exists(self, mail_obj=None, item=None, tier="low"):
        """
        Check if the parent/quoted email exists in Outlook.
        Returns True if found, False otherwise.
        """
        try:
            # Tier LOW: Check via In-Reply-To header
            # Note: Outlook Object Model doesn't expose headers easily on MailItem without PropertyAccessor
            # We use the parsed mail object for headers if available
            in_reply_to = None
            if mail_obj and hasattr(mail_obj, "headers"):
                in_reply_to = mail_obj.headers.get("In-Reply-To")

            # If we only have item (fallback mode), try PropertyAccessor
            if not in_reply_to and item:
                try:
                    # PR_INTERNET_REFERENCES = http://schemas.microsoft.com/mapi/proptag/0x1039001E
                    # PR_IN_REPLY_TO_ID = http://schemas.microsoft.com/mapi/proptag/0x1042001E
                    prop_accessor = item.PropertyAccessor
                    in_reply_to = prop_accessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1042001E")
                except Exception:
                    pass

            if in_reply_to:
                # Search by InternetMessageId
                # Clean up ID (remove < >)
                msg_id = in_reply_to.strip("<> ")
                if msg_id:
                    # Search inbox (or all folders? Restrict is usually folder-bound)
                    # For performance/simplicity, check Inbox first.
                    inbox = self.get_inbox()
                    if inbox:
                        # PR_INTERNET_MESSAGE_ID = http://schemas.microsoft.com/mapi/proptag/0x1035001E
                        # Jet query: [InternetMessageId] = 'id' (doesn't always work reliably)
                        # DASL query is better
                        dasl_filter = f"@SQL=\"http://schemas.microsoft.com/mapi/proptag/0x1035001E\" = '{msg_id.replace("'", "''")}'"
                        found_items = inbox.Items.Restrict(dasl_filter)
                        if found_items.Count > 0:
                            return True

            # Tier MEDIUM: Check by Subject (if low failed or not possible)
            if tier == "medium":
                subject = getattr(mail_obj, "subject", None) or (item.Subject if item else None)
                if subject:
                    # Strip RE/FW prefixes
                    import re
                    clean_subject = re.sub(r"^((re|fw|fwd):\s*)+", "", subject, flags=re.IGNORECASE).strip()
                    if clean_subject:
                        inbox = self.get_inbox()
                        if inbox:
                            # Search for matching subject
                            # Use simplified check
                            # Just try to find "Conversations" -> GetConversation() is best but complex
                            # Let's try restrictive search
                            # Try strict subject match of CLEAN subject
                            # This is heuristic.
                            # Just check if ANY item has this subject
                            items = inbox.Items.Restrict(f"[Subject] = '{clean_subject.replace("'", "''")}'")
                            if items.Count > 0:
                                return True

            return False

        except Exception as e:
            logger.error(f"Error checking parent existence: {e}")
            return False

    def _extract_latest_reply(self, text_body):
        """Extract latest reply using mail-parser-reply"""
        if not text_body:
            return None
        try:
            from mailparser_reply import EmailReplyParser

            parsed = EmailReplyParser.read(text_body)
            # return the latest reply text
            return parsed.latest_reply
        except ImportError:
            logger.warning(
                "mail-parser-reply not installed, skipping reply extraction"
            )
            return None
        except Exception as e:
            logger.error(f"Error extracting reply: {e}")
            return None

    def _convert_to_parsed_model(
        self, mail, item, deduplication_tier="none", strip_html=True
    ):
        """Convert mail-parser object to dict matching EmailParsed model"""
        # mail-parser returns list of tuples for from_, to, cc, bcc
        # mail.date is a datetime object

        # Helper to convert list of tuples to expected format
        def safe_list_tuples(val):
            if isinstance(val, list):
                return [tuple(x) for x in val]
            return []

        # received headers
        received = []
        if hasattr(mail, "received"):
            for r in mail.received:
                # convert to dict if it's not
                if isinstance(r, dict):
                    received.append(r)
                else:
                    try:
                        received.append(dict(r))
                    except Exception:
                        pass

        # attachments
        # User said "leave out the attachment for now", so I will return metadata only or empty
        # mail-parser attachments is list of dicts.
        attachments = []
        if hasattr(mail, "attachments"):
            for att in mail.attachments:
                # filter fields? Just keep what mail-parser gives
                # binary payload might be large.
                # User said "leave out the attachment for now".
                # I'll include metadata but maybe strip payload if huge?
                # mail-parser 'payload' is base64 string.
                # I'll strip 'payload' key to save bandwidth as requested "leave out attachment".
                # But keep filename, etc.
                a_copy = att.copy()
                if "payload" in a_copy:
                    del a_copy["payload"]
                attachments.append(a_copy)

        latest_reply = None
        parent_found = None

        should_strip = False

        if deduplication_tier != "none":
            # Extract reply text candidate
            latest_reply = self._extract_latest_reply(mail.body)

            if deduplication_tier == "low":
                # Low: Just strip (matches original remove_quoted=True behavior/default requirement)
                # "low (default) - use only outlook specific In-Reply-To... metadata"
                # This implies checking logic.
                # If parent exists -> Strip.
                # If not -> Keep? Or assumes if In-Reply-To is present, it's a reply?
                # User said "use only... metadata... medium... tries to deduplicate with same title".
                # I will interpret "Low" as: Check In-Reply-To. If found in DB, strip.
                parent_found = self._check_parent_exists(mail_obj=mail, item=item, tier="low")
                if parent_found:
                    should_strip = True
                else:
                    # If we can't find parent, we preserve full body
                    should_strip = False

            elif deduplication_tier == "medium":
                # Medium: Check In-Reply-To OR Subject
                parent_found = self._check_parent_exists(mail_obj=mail, item=item, tier="medium")
                if parent_found:
                    should_strip = True

            elif deduplication_tier == "high":
                # High: Same as medium for now (content search is slow/complex via COM)
                parent_found = self._check_parent_exists(mail_obj=mail, item=item, tier="medium")
                if parent_found:
                    should_strip = True

        # If stripping is active and we have a reply extracted
        if should_strip and latest_reply:
            final_body = latest_reply
        else:
            final_body = mail.body

        text_html = mail.text_html

        if strip_html:
            # Check if we need to convert HTML to text
            # If body is empty or looks like HTML, and we have HTML content
            import re

            is_html = bool(re.search(r"<[a-z][\s\S]*>", final_body, re.IGNORECASE))
            if is_html or (not final_body and text_html):
                try:
                    from bs4 import BeautifulSoup

                    # Use text_html if body is empty
                    source = (
                        final_body
                        if final_body and is_html
                        else (text_html[0] if text_html else "")
                    )
                    if source:
                        soup = BeautifulSoup(source, "html.parser")
                        final_body = soup.get_text(separator="\n").strip()
                except ImportError:
                    logger.warning(
                        "beautifulsoup4 not installed, skipping HTML stripping"
                    )
                except Exception as e:
                    logger.error(f"Error stripping HTML: {e}")

            # Clear text_html to save context
            text_html = []

        return {
            "entry_id": item.EntryID,
            "subject": mail.subject,
            "from": safe_list_tuples(mail.from_),
            "to": safe_list_tuples(mail.to),
            "cc": safe_list_tuples(mail.cc),
            "bcc": safe_list_tuples(mail.bcc),
            "date": mail.date.isoformat()
            if mail.date
            else (
                item.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
                if item.ReceivedTime
                else None
            ),
            "message_id": mail.message_id,
            "headers": mail.headers,
            "text_plain": mail.text_plain,
            "text_html": text_html,
            "body": final_body,
            "attachments": attachments,
            "received": received,
            "latest_reply": latest_reply,
            "deduplication_tier": deduplication_tier,
            "parent_found": parent_found,
        }

    def _fallback_parsed_model(
        self, item, deduplication_tier="none", strip_html=True
    ):
        """Fallback when mail-parser fails, using COM properties"""
        sender_email = self.resolve_smtp_address(item)
        sender_name = item.SenderName

        # Parse recipients
        to_list = []
        cc_list = []
        bcc_list = []

        # This is a basic approximation
        if hasattr(item, "To") and item.To:
            # item.To is a string "Name; Name". No emails usually.
            # Best effort: use Recipients collection if possible, but that iterates.
            # For fallback, just putting name in tuple is okay or split string.
            # Or just use empty list if we can't reliably get emails.
            # I'll use the strings.
            for name in item.To.split(";"):
                if name.strip():
                    to_list.append((name.strip(), ""))

        if hasattr(item, "CC") and item.CC:
            for name in item.CC.split(";"):
                if name.strip():
                    cc_list.append((name.strip(), ""))

        received_time = (
            item.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
            if item.ReceivedTime
            else None
        )

        # Attachments metadata (basic)
        attachments = []
        try:
            if item.Attachments.Count > 0:
                for i in range(1, item.Attachments.Count + 1):
                    att = item.Attachments.Item(i)
                    attachments.append({
                        "filename": att.FileName,
                        "size": att.Size,
                        "content_id": self._safe_get_attr(att, "ContentID"),
                    })
        except Exception:
            pass

        latest_reply = None
        parent_found = None
        should_strip = False

        if deduplication_tier != "none":
            latest_reply = self._extract_latest_reply(item.Body)

            # For fallback, we only have 'item'
            if deduplication_tier in ["low", "medium", "high"]:
                parent_found = self._check_parent_exists(item=item, tier=deduplication_tier)
                if parent_found:
                    should_strip = True

        if should_strip and latest_reply:
            final_body = latest_reply
        else:
            final_body = item.Body

        text_html = [item.HTMLBody]

        if strip_html:
            # If body is empty, try to get from HTML
            # Note: item.Body is usually plain text in Outlook Object Model
            if not final_body and item.HTMLBody:
                try:
                    from bs4 import BeautifulSoup

                    soup = BeautifulSoup(item.HTMLBody, "html.parser")
                    final_body = soup.get_text(separator="\n").strip()
                except ImportError:
                    pass
                except Exception:
                    pass

            # Clear text_html
            text_html = []

        return {
            "entry_id": item.EntryID,
            "subject": item.Subject,
            "from": [(sender_name, sender_email)],
            "to": to_list,
            "cc": cc_list,
            "bcc": bcc_list,
            "date": received_time,
            "message_id": "",
            "headers": {},
            "text_plain": [item.Body],
            "text_html": text_html,
            "body": final_body,
            "attachments": attachments,
            "received": [],
            "latest_reply": latest_reply,
            "deduplication_tier": deduplication_tier,
            "parent_found": parent_found,
        }

    def list_calendar_events(self, days=7, all_events=False):
        """
        List calendar events for the next N days

        Args:
            days: Number of days ahead to look
            all_events: If True, return all events without date filtering

        Returns:
            List of event dictionaries
        """
        calendar = self.get_calendar()
        items = calendar.Items

        # CRITICAL: Filter to only appointment items before any other operations
        # This prevents COM errors when encountering meeting requests/responses
        items = items.Restrict(
            "[MessageClass] >= 'IPM.Appointment' AND [MessageClass] < 'IPM.Appointment{'"
        )

        # CRITICAL: Enable recurrence expansion BEFORE sorting
        # Must sort ascending for recurrence to work properly
        items.IncludeRecurrences = True
        items.Sort("[Start]")  # Ascending for recurrence

        # CRITICAL FIX: Apply Restrict BEFORE iterating to avoid "Calendar Bomb"
        # Without this, recurring meetings without end dates generate infinite items
        if not all_events:
            start_date = datetime.now()
            end_date = start_date + timedelta(days=days)
            # Jet SQL format for dates: MM/DD/YYYY HH:MM
            # Use Restrict to filter at COM level before Python iteration
            filter_str = (
                f"[Start] <= '{end_date.strftime('%m/%d/%Y %H:%M')}' "
                f"AND [End] >= '{start_date.strftime('%m/%d/%Y %H:%M')}'"
            )
            items = items.Restrict(filter_str)

        events = []
        for item in items:
            try:
                # Use safe attribute access to handle COM errors
                start = self._safe_get_attr(item, "Start")
                end = self._safe_get_attr(item, "End")

                # Skip if no start time
                if not start:
                    continue

                # Additional Python-level filtering for safety (in case Restrict wasn't applied)
                if not all_events:
                    start_date = datetime.now()
                    end_date = start_date + timedelta(days=days)
                    # Normalize COM/pywintypes datetimes to naive Python datetimes for comparison
                    start_dt = None
                    try:
                        if isinstance(start, datetime):
                            # drop tzinfo if present to compare with datetime.now()
                            start_dt = datetime(
                                start.year,
                                start.month,
                                start.day,
                                start.hour,
                                start.minute,
                                start.second,
                            )
                        else:
                            start_dt = datetime(
                                start.Year,
                                start.Month,
                                start.Day,
                                start.Hour,
                                start.Minute,
                                start.Second,
                            )
                    except Exception:
                        # If normalization fails, skip this item
                        continue

                    if not (start_dt >= start_date and start_dt <= end_date):
                        continue

                # Get attendees (safe access)
                required_attendees = self._safe_get_attr(item, "RequiredAttendees", "")
                optional_attendees = self._safe_get_attr(item, "OptionalAttendees", "")

                # Get meeting status
                # ResponseStatus: 0=None, 1=Organizer, 2=Tentative, 3=Accepted, 4=Declined, 5=NotResponded
                response_status = self._safe_get_attr(item, "ResponseStatus")
                response_status_map = {
                    0: "None",
                    1: "Organizer",
                    2: "Tentative",
                    3: "Accepted",
                    4: "Declined",
                    5: "NotResponded",
                }

                # MeetingStatus: 0=Non-meeting, 1=Meeting, 2=Received, 3=Canceled
                meeting_status = self._safe_get_attr(item, "MeetingStatus")
                meeting_status_map = {
                    0: "NonMeeting",
                    1: "Meeting",
                    2: "Received",
                    3: "Canceled",
                }

                event = {
                    "entry_id": self._safe_get_attr(item, "EntryID", ""),
                    "subject": self._safe_get_attr(item, "Subject", "(No Subject)"),
                    "start": start.strftime("%Y-%m-%d %H:%M:%S") if start else None,
                    "end": end.strftime("%Y-%m-%d %H:%M:%S") if end else None,
                    "location": self._safe_get_attr(item, "Location", ""),
                    "organizer": self._safe_get_attr(item, "Organizer"),
                    "all_day": self._safe_get_attr(item, "AllDayEvent", False),
                    "required_attendees": required_attendees,
                    "optional_attendees": optional_attendees,
                    "response_status": response_status_map.get(
                        response_status, "Unknown"
                    ),
                    "meeting_status": meeting_status_map.get(meeting_status, "Unknown"),
                    "response_requested": self._safe_get_attr(
                        item, "ResponseRequested", False
                    ),
                }
                events.append(event)
            except (Exception, BaseException):
                # Skip items that cause errors (including COM fatal errors)
                continue

        return events

    def _create_mail_item(self, save_draft=False):
        """
        Create a new mail item, optionally in the Drafts folder

        Args:
            save_draft: If True, attempt to create in the Drafts folder

        Returns:
            Outlook MailItem
        """
        if save_draft:
            drafts = self.get_folder_by_name("Drafts")
            if drafts:
                try:
                    return drafts.Items.Add()
                except Exception:
                    pass

        return self.outlook.CreateItem(0)  # 0 = olMailItem

    def _add_attachments(self, mail_item, file_paths):
        """
        Add attachments to a mail item

        Args:
            mail_item: Outlook MailItem
            file_paths: List of file paths to attach
        """
        if not file_paths:
            return

        for file_path in file_paths:
            with contextlib.suppress(Exception):
                mail_item.Attachments.Add(file_path)

    def _set_sender_account(self, mail_item):
        """
        Set the sender account for a mail item based on the default account

        Args:
            mail_item: Outlook MailItem
        """
        try:
            acc = None
            if self.default_account_name:
                acc = self._find_account_by_name(self.default_account_name)

            # If DefaultStore was set, try to find account by matching store owner
            if not acc:
                try:
                    accounts = self.namespace.Accounts
                    for a in accounts:
                        try:
                            if hasattr(a, "SmtpAddress") and a.SmtpAddress in (
                                self.default_account_name or ""
                            ):
                                acc = a
                                break
                        except Exception:
                            continue
                except Exception:
                    pass

            if acc:
                with contextlib.suppress(Exception):
                    mail_item.SendUsingAccount = acc
        except Exception:
            pass

    def send_email(
        self,
        to,
        subject,
        body,
        cc=None,
        bcc=None,
        html_body=None,
        file_paths=None,
        save_draft=False,
    ):
        """
        Send an email (or save as draft)

        Args:
            to: Recipient email address
            subject: Email subject
            body: Email body (plain text)
            cc: CC recipients (optional)
            bcc: BCC recipients (optional)
            html_body: HTML body (optional)
            file_paths: List of file paths to attach (optional)
            save_draft: If True, save to Drafts instead of sending

        Returns:
            Draft entry ID if saved, True if sent, False if failed
        """
        try:
            mail = self._create_mail_item(save_draft=save_draft)

            mail.To = to
            mail.Subject = subject
            if html_body:
                mail.HTMLBody = html_body
            else:
                mail.Body = body
            if cc:
                mail.CC = cc
            if bcc:
                mail.BCC = bcc

            # Add attachments
            self._add_attachments(mail, file_paths)

            # Ensure sending uses the default account when set
            self._set_sender_account(mail)

            if save_draft:
                mail.Save()
                return mail.EntryID
            else:
                mail.Send()
                return True
        except Exception as e:
            logger.error(f"Error sending email: {e}")
            return False

    def reply_email(self, entry_id, body, reply_all=False):
        """
        Reply to an email (O(1) direct access)

        Args:
            entry_id: Email entry ID
            body: Reply body
            reply_all: True to reply all, False to reply sender only

        Returns:
            True if successful
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                if reply_all:
                    reply = item.ReplyAll()
                else:
                    reply = item.Reply()
                reply.Body = body
                # try to enforce default send account
                try:
                    if self.default_account_name:
                        acc = self._find_account_by_name(self.default_account_name)
                        if acc:
                            with contextlib.suppress(Exception):
                                reply.SendUsingAccount = acc
                except Exception:
                    pass
                reply.Send()
                return True
            except Exception as e:
                logger.error(f"Error replying to email: {e}")
                return False
        return False

    def forward_email(self, entry_id, to, body=""):
        """
        Forward an email (O(1) direct access)

        Args:
            entry_id: Email entry ID
            to: Recipient to forward to
            body: Optional additional body text

        Returns:
            True if successful
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                forward = item.Forward()
                forward.To = to
                if body:
                    forward.Body = body + "\n\n" + forward.Body
                try:
                    if self.default_account_name:
                        acc = self._find_account_by_name(self.default_account_name)
                        if acc:
                            with contextlib.suppress(Exception):
                                forward.SendUsingAccount = acc
                except Exception:
                    pass
                forward.Send()
                return True
            except Exception as e:
                logger.error(f"Error forwarding email: {e}")
                return False
        return False

    def mark_email_read(self, entry_id, unread=False):
        """
        Mark an email as read or unread (O(1) direct access)

        Args:
            entry_id: Email entry ID
            unread: True to mark as unread, False to mark as read

        Returns:
            True if successful
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                item.Unread = unread
                item.Save()
                return True
            except Exception:
                return False
        return False

    def move_email(self, entry_id, folder_name):
        """
        Move an email to a different folder (O(1) direct access)

        Args:
            entry_id: Email entry ID
            folder_name: Target folder name

        Returns:
            True if successful
        """
        try:
            item = self.get_item_by_id(entry_id)
            if not item:
                return False

            target_folder = self.get_folder_by_name(folder_name)
            if not target_folder:
                logger.error(f"Folder '{folder_name}' not found")
                return False

            item.Move(target_folder)
            return True
        except Exception as e:
            logger.error(f"Error moving email: {e}")
            return False

    def delete_email(self, entry_id):
        """
        Delete an email (O(1) direct access)

        Args:
            entry_id: Email entry ID

        Returns:
            True if successful
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                item.Delete()
                return True
            except Exception:
                return False
        return False

    def download_attachments(self, entry_id, download_dir):
        """
        Download all attachments from an email

        Args:
            entry_id: Email entry ID
            download_dir: Directory to save attachments

        Returns:
            List of downloaded file paths
        """
        item = self.get_item_by_id(entry_id)
        if not item or item.Attachments.Count == 0:
            return []

        downloaded = []
        try:
            import os

            os.makedirs(download_dir, exist_ok=True)

            for i in range(item.Attachments.Count):
                attachment = item.Attachments.Item(i + 1)  # COM is 1-indexed

                # Security Fix: Sanitize filename to prevent path traversal
                # Get just the filename part, handling both / and \ regardless of platform
                raw_filename = attachment.FileName or f"attachment_{i + 1}"
                filename = os.path.basename(raw_filename.replace("\\", "/"))

                # Ensure filename is not empty or special directory markers
                if not filename or filename in (".", ".."):
                    filename = f"attachment_{i + 1}"

                # Construct full path and ensure it's within download_dir
                abs_download_dir = os.path.abspath(download_dir)
                filepath = os.path.abspath(os.path.join(abs_download_dir, filename))

                if not filepath.startswith(abs_download_dir):
                    logger.warning(
                        f"Skipping suspicious attachment filename: {raw_filename}"
                    )
                    continue

                attachment.SaveAsFile(filepath)
                downloaded.append(filepath)
            return downloaded
        except Exception as e:
            logger.error(f"Error downloading attachments: {e}")
            return []

    def create_appointment(
        self,
        subject,
        start,
        end,
        location="",
        body="",
        all_day=False,
        required_attendees=None,
        optional_attendees=None,
    ):
        """
        Create a calendar appointment

        Args:
            subject: Appointment subject
            start: Start time (YYYY-MM-DD HH:MM:SS)
            end: End time (YYYY-MM-DD HH:MM:SS)
            location: Location
            body: Appointment body/description
            all_day: True for all-day event
            required_attendees: Semicolon-separated list of required attendees
            optional_attendees: Semicolon-separated list of optional attendees

        Returns:
            Appointment entry ID if successful
        """
        try:
            # Prefer creating the appointment in the default account's Calendar folder
            calendar = self.get_calendar()
            if calendar:
                try:
                    appointment = calendar.Items.Add()
                except Exception:
                    appointment = self.outlook.CreateItem(1)  # fallback
            else:
                appointment = self.outlook.CreateItem(1)  # 1 = olAppointmentItem
            appointment.Subject = subject
            appointment.Start = datetime.strptime(start, "%Y-%m-%d %H:%M:%S")
            appointment.End = datetime.strptime(end, "%Y-%m-%d %H:%M:%S")
            appointment.Location = location
            appointment.Body = body
            appointment.AllDayEvent = all_day
            if required_attendees:
                appointment.RequiredAttendees = required_attendees
            if optional_attendees:
                appointment.OptionalAttendees = optional_attendees
            appointment.Save()
            return appointment.EntryID
        except Exception as e:
            logger.error(f"Error creating appointment: {e}")
            return None

    def edit_appointment(
        self,
        entry_id,
        required_attendees=None,
        optional_attendees=None,
        subject=None,
        start=None,
        end=None,
        location=None,
        body=None,
    ):
        """
        Edit an existing appointment

        Args:
            entry_id: Appointment entry ID
            required_attendees: Comma-separated list of required attendees
            optional_attendees: Comma-separated list of optional attendees
            subject: New subject (optional)
            start: New start time (optional)
            end: New end time (optional)
            location: New location (optional)
            body: New body (optional)

        Returns:
            True if successful
        """
        try:
            calendar = self.get_calendar()
            for item in calendar.Items:
                if item.EntryID == entry_id:
                    if required_attendees:
                        item.RequiredAttendees = required_attendees
                    if optional_attendees:
                        item.OptionalAttendees = optional_attendees
                    if subject:
                        item.Subject = subject
                    if start:
                        item.Start = datetime.strptime(start, "%Y-%m-%d %H:%M:%S")
                    if end:
                        item.End = datetime.strptime(end, "%Y-%m-%d %H:%M:%S")
                    if location is not None:
                        item.Location = location
                    if body is not None:
                        item.Body = body
                    item.Save()
                    return True
            return False
        except Exception as e:
            logger.error(f"Error editing appointment: {e}")
            return False

    def get_appointment(self, entry_id):
        """
        Get full appointment details by entry ID (O(1) direct access)

        Args:
            entry_id: Appointment entry ID

        Returns:
            Appointment dictionary with full details
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                required_attendees = (
                    item.RequiredAttendees if hasattr(item, "RequiredAttendees") else ""
                )
                optional_attendees = (
                    item.OptionalAttendees if hasattr(item, "OptionalAttendees") else ""
                )
                response_status = (
                    item.ResponseStatus if hasattr(item, "ResponseStatus") else None
                )
                response_status_map = {
                    0: "None",
                    1: "Organizer",
                    2: "Tentative",
                    3: "Accepted",
                    4: "Declined",
                    5: "NotResponded",
                }
                meeting_status = (
                    item.MeetingStatus if hasattr(item, "MeetingStatus") else None
                )
                meeting_status_map = {
                    0: "NonMeeting",
                    1: "Meeting",
                    2: "Received",
                    3: "Canceled",
                }

                return {
                    "entry_id": item.EntryID,
                    "subject": item.Subject
                    if hasattr(item, "Subject")
                    else "(No Subject)",
                    "start": item.Start.strftime("%Y-%m-%d %H:%M:%S")
                    if hasattr(item, "Start") and item.Start
                    else None,
                    "end": item.End.strftime("%Y-%m-%d %H:%M:%S")
                    if hasattr(item, "End") and item.End
                    else None,
                    "location": item.Location if hasattr(item, "Location") else "",
                    "organizer": item.Organizer if hasattr(item, "Organizer") else None,
                    "body": item.Body if hasattr(item, "Body") else "",
                    "all_day": item.AllDayEvent
                    if hasattr(item, "AllDayEvent")
                    else False,
                    "required_attendees": required_attendees,
                    "optional_attendees": optional_attendees,
                    "response_status": response_status_map.get(
                        response_status, "Unknown"
                    ),
                    "meeting_status": meeting_status_map.get(meeting_status, "Unknown"),
                    "response_requested": item.ResponseRequested
                    if hasattr(item, "ResponseRequested")
                    else False,
                }
            except Exception:
                pass
        return None

    def respond_to_meeting(self, entry_id, response):
        """
        Respond to a meeting invitation (O(1) direct access)

        Args:
            entry_id: Appointment entry ID
            response: Response - "accept", "decline", "tentative"

        Returns:
            True if successful
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                response_map = {
                    "accept": 3,  # olResponseAccepted
                    "decline": 4,  # olResponseDeclined
                    "tentative": 2,  # olResponseTentative
                }
                if response.lower() in response_map:
                    item.Response(response_map[response.lower()])
                    item.Send()
                    return True
            except Exception:
                pass
        return False

    def delete_appointment(self, entry_id):
        """
        Delete an appointment (O(1) direct access)

        Args:
            entry_id: Appointment entry ID

        Returns:
            True if successful
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                item.Delete()
                return True
            except Exception:
                pass
        return False

    def list_tasks(self, include_completed=False):
        """
        List tasks (only incomplete by default)

        Args:
            include_completed: If True, return all tasks. If False (default), return only incomplete tasks.

        Returns:
            List of task dictionaries
        """
        tasks_folder = self.get_tasks()
        items = tasks_folder.Items

        tasks = []

        # Optimization: Filter on server side if possible
        if not include_completed:
            with contextlib.suppress(Exception):
                items = items.Restrict("[Complete] = False")

        for item in items:
            try:
                # Skip completed tasks unless include_completed is True
                if not include_completed and item.Complete:
                    continue

                task = {
                    "entry_id": item.EntryID,
                    "subject": item.Subject
                    if hasattr(item, "Subject")
                    else "(No Subject)",
                    "body": item.Body if hasattr(item, "Body") else "",
                    "due_date": item.DueDate.strftime("%Y-%m-%d")
                    if hasattr(item, "DueDate") and item.DueDate
                    else None,
                    "status": item.Status if hasattr(item, "Status") else None,
                    "priority": item.Importance
                    if hasattr(item, "Importance")
                    else None,
                    "complete": item.Complete if hasattr(item, "Complete") else False,
                    "percent_complete": item.PercentComplete
                    if hasattr(item, "PercentComplete")
                    else 0,
                }
                tasks.append(task)
            except Exception:
                continue

        return tasks

    def list_all_tasks(self):
        """
        List all tasks including completed ones

        Returns:
            List of task dictionaries
        """
        return self.list_tasks(include_completed=True)

    def create_task(self, subject, body="", due_date=None, importance=1):
        """
        Create a new task

        Args:
            subject: Task subject
            body: Task description
            due_date: Due date (YYYY-MM-DD) or None
            importance: 0=Low, 1=Normal, 2=High

        Returns:
            Task entry ID if successful
        """
        try:
            # Prefer creating the task in the default account's Tasks folder
            tasks_folder = self.get_tasks()
            if tasks_folder:
                try:
                    task = tasks_folder.Items.Add()
                except Exception:
                    task = self.outlook.CreateItem(3)
            else:
                task = self.outlook.CreateItem(3)  # 3 = olTaskItem
            task.Subject = subject
            task.Body = body
            if due_date:
                # Use noon to avoid timezone boundary issues
                task.DueDate = datetime.strptime(
                    f"{due_date} 12:00:00", "%Y-%m-%d %H:%M:%S"
                )
            task.Importance = importance
            task.Save()
            return task.EntryID
        except Exception as e:
            logger.error(f"Error creating task: {e}")
            return None

    def get_task(self, entry_id):
        """
        Get full task details by entry ID (O(1) direct access)

        Args:
            entry_id: Task entry ID

        Returns:
            Task dictionary with full details
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                return {
                    "entry_id": item.EntryID,
                    "subject": item.Subject
                    if hasattr(item, "Subject")
                    else "(No Subject)",
                    "body": item.Body if hasattr(item, "Body") else "",
                    "due_date": item.DueDate.strftime("%Y-%m-%d")
                    if hasattr(item, "DueDate") and item.DueDate
                    else None,
                    "status": item.Status if hasattr(item, "Status") else None,
                    "priority": item.Importance
                    if hasattr(item, "Importance")
                    else None,
                    "complete": item.Complete if hasattr(item, "Complete") else False,
                    "percent_complete": item.PercentComplete
                    if hasattr(item, "PercentComplete")
                    else 0,
                }
            except Exception:
                pass
        return None

    def edit_task(
        self,
        entry_id,
        subject=None,
        body=None,
        due_date=None,
        importance=None,
        percent_complete=None,
        complete=None,
    ):
        """
        Edit an existing task (O(1) direct access)

        Args:
            entry_id: Task entry ID
            subject: New subject (optional)
            body: New body (optional)
            due_date: New due date YYYY-MM-DD (optional)
            importance: New importance 0=Low, 1=Normal, 2=High (optional)
            percent_complete: New percent complete 0-100 (optional)
            complete: Mark complete/incomplete True/False (optional)

        Returns:
            True if successful
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                if subject:
                    item.Subject = subject
                if body is not None:
                    item.Body = body
                if due_date:
                    # Use noon to avoid timezone boundary issues
                    item.DueDate = datetime.strptime(
                        f"{due_date} 12:00:00", "%Y-%m-%d %H:%M:%S"
                    )
                if importance is not None:
                    item.Importance = importance
                if percent_complete is not None:
                    item.PercentComplete = percent_complete
                    # Update status based on percent_complete
                    if percent_complete == 100:
                        item.Status = 2  # Complete
                        item.Complete = True
                    elif percent_complete == 0:
                        item.Status = 0  # Not started
                    else:
                        item.Status = 1  # In progress
                if complete is not None:
                    item.Complete = complete
                    if complete:
                        item.PercentComplete = 100
                        item.Status = 2
                    else:
                        item.PercentComplete = 0
                        item.Status = 0
                item.Save()
                return True
            except Exception as e:
                logger.error(f"Error editing task: {e}")
            pass
        return False

    def complete_task(self, entry_id):
        """
        Mark a task as complete (O(1) direct access)

        Args:
            entry_id: Task entry ID

        Returns:
            True if successful
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                item.Complete = True
                item.PercentComplete = 100
                item.Status = 2  # olTaskComplete
                item.Save()
                return True
            except Exception:
                pass
        return False

    def delete_task(self, entry_id):
        """
        Delete a task (O(1) direct access)

        Args:
            entry_id: Task entry ID

        Returns:
            True if successful
        """
        item = self.get_item_by_id(entry_id)
        if item:
            try:
                item.Delete()
                return True
            except Exception:
                pass
        return False

    def _search_emails_raw(self, filter_query, limit=100, folder="Inbox"):
        """
        Internal: Search emails using raw Outlook Restriction filter.

        Args:
            filter_query: SQL query string for filtering
            limit: Max results to return
            folder: Folder to search in (default: Inbox)

        Returns:
            List of email dictionaries
        """
        try:
            # Get folder
            if folder == "Inbox":
                mail_folder = self.get_inbox()
            else:
                mail_folder = self.get_folder_by_name(folder)
                if not mail_folder:
                    mail_folder = self.get_inbox()

            items = mail_folder.Items
            # Apply restriction filter
            items = items.Restrict(filter_query)

            # Sort by received time, most recent first
            items.Sort("[ReceivedTime]", True)

            emails = []
            count = 0
            for item in items:
                if count >= limit:
                    break

                try:
                    email = {
                        "entry_id": item.EntryID,
                        "subject": item.Subject,
                        "sender": self.resolve_smtp_address(item),
                        "sender_name": item.SenderName,
                        "received_time": item.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
                        if item.ReceivedTime
                        else None,
                        "unread": item.Unread,
                        "has_attachments": item.Attachments.Count > 0,
                    }
                    emails.append(email)
                    count += 1
                except Exception:
                    # Skip items that can't be accessed
                    continue

            return emails
        except Exception as e:
            logger.error(f"Error searching emails: {e}")
            return []

    def search_emails(
        self,
        filter_query=None,
        limit=100,
        subject=None,
        sender=None,
        body=None,
        unread=None,
        has_attachments=None,
        folder="Inbox",
    ):
        """
        Search emails using structured criteria.

        Args:
            filter_query: Raw SQL/DASL query (Unsafe, legacy support)
            limit: Max results
            subject: Subject substring to match
            sender: Sender substring to match
            body: Body substring to match
            unread: Filter by unread status (True/False)
            has_attachments: Filter by attachment presence (True/False)
            folder: Folder to search in (default: Inbox)

        Returns:
            List of email dictionaries
        """
        if filter_query:
            return self._search_emails_raw(filter_query, limit, folder=folder)

        filters = []

        if subject:
            # Escape single quotes
            safe_subject = subject.replace("'", "''")
            filters.append(
                f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{safe_subject}%'"
            )

        if body:
            safe_body = body.replace("'", "''")
            filters.append(
                f"@SQL=\"urn:schemas:httpmail:textdescription\" LIKE '%{safe_body}%'"
            )

        if sender:
            # Match either name or email
            safe_sender = sender.replace("'", "''")
            filters.append(
                f"(\"@SQL=\"urn:schemas:httpmail:fromname\" LIKE '%{safe_sender}%' OR "
                f"\"urn:schemas:httpmail:fromemail\" LIKE '%{safe_sender}%'\")"
            )

        if unread is not None:
            filters.append(f"[Unread] = {'True' if unread else 'False'}")

        if has_attachments is not None:
            filters.append(f"[HasAttachments] = {'True' if has_attachments else 'False'}")

        query = " AND ".join(filters) if filters else ""

        if not query:
            # Just list recent emails if no filters provided
            return self.list_emails(limit=limit, folder=folder)

        return self._search_emails_raw(query, limit, folder=folder)

    def search_by_sender(self, sender_email, limit=100, folder="Inbox"):
        """
        Search emails by sender email address (handles Exchange addresses).

        This method properly handles both SMTP and Exchange email addresses.
        For Exchange users (internal emails), it resolves the Exchange address
        to SMTP address before matching.

        Args:
            sender_email: Email address to search for
            limit: Max results to return (default: 100)
            folder: Folder name to search in (default: "Inbox")

        Returns:
            List of email dictionaries matching the sender
        """
        try:
            # Get the folder
            if folder == "Inbox":
                mail_folder = self.get_inbox()
            else:
                mail_folder = self.get_folder_by_name(folder)
                if not mail_folder:
                    mail_folder = self.get_inbox()

            items = mail_folder.Items
            # Sort by received time, most recent first
            items.Sort("[ReceivedTime]", True)

            emails = []
            count = 0
            for item in items:
                if count >= limit:
                    break

                try:
                    # Resolve SMTP address (handles Exchange addresses)
                    smtp_address = self.resolve_smtp_address(item)

                    # Case-insensitive email match
                    if smtp_address.lower() == sender_email.lower():
                        email = {
                            "entry_id": item.EntryID,
                            "subject": item.Subject,
                            "sender": smtp_address,
                            "sender_name": item.SenderName,
                            "received_time": item.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")
                            if item.ReceivedTime
                            else None,
                            "unread": item.Unread,
                            "has_attachments": item.Attachments.Count > 0,
                        }
                        emails.append(email)
                        count += 1
                except Exception:
                    # Skip items that can't be accessed
                    continue

            return emails
        except Exception as e:
            logger.error(f"Error searching emails by sender: {e}")
            return []

    def get_free_busy(
        self, email_address=None, start_date=None, end_date=None, entry_id=None
    ):
        """
        Get free/busy status for an email address

        Args:
            email_address: Email address to check (optional, defaults to current user)
            start_date: Start date (YYYY-MM-DD) or datetime object (optional, defaults to today)
            end_date: End date (YYYY-MM-DD) or datetime object (optional, defaults to start + 1 day)
            entry_id: DEPRECATED - Appointment entry ID (legacy, use email_address instead)

        Returns:
            Dictionary with free/busy information
        """
        try:
            # Handle legacy entry_id parameter (extract first required attendee)
            if entry_id and not email_address:
                item = self.get_item_by_id(entry_id)
                if (
                    item
                    and hasattr(item, "RequiredAttendees")
                    and item.RequiredAttendees
                ):
                    attendees = item.RequiredAttendees.split(";")
                    if attendees:
                        email_address = attendees[0].strip()

            # Default to current user if no email provided
            if not email_address:
                email_address = self.namespace.CurrentUser.Address

            # Default to today if no dates provided
            if not start_date:
                start_date = datetime.now()
            elif isinstance(start_date, str):
                start_date = datetime.strptime(start_date, "%Y-%m-%d")

            if not end_date:
                end_date = start_date + timedelta(days=1)
            elif isinstance(end_date, str):
                end_date = datetime.strptime(end_date, "%Y-%m-%d")

            # Create recipient and get free/busy
            recipient = self.namespace.CreateRecipient(email_address)
            if recipient.Resolve():
                # FreeBusy returns a string with time slots and status
                # 0=Free, 1=Tentative, 2=Busy, 3=Out of Office, 4=Working Elsewhere
                freebusy = recipient.FreeBusy(
                    start_date, 60 * 24
                )  # 1440 minutes = 1 day
                return {
                    "email": email_address,
                    "start_date": start_date.strftime("%Y-%m-%d"),
                    "end_date": end_date.strftime("%Y-%m-%d"),
                    "free_busy": freebusy,
                    "resolved": True,
                }
            else:
                return {
                    "email": email_address,
                    "start_date": start_date.strftime("%Y-%m-%d"),
                    "end_date": end_date.strftime("%Y-%m-%d"),
                    "error": "Could not resolve email address",
                    "resolved": False,
                }
        except Exception as e:
            return {
                "email": email_address if email_address else "unknown",
                "error": str(e),
                "resolved": False,
            }
