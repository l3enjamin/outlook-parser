
import sys
import time
from unittest.mock import MagicMock
from datetime import datetime

# Define constants for MAPI properties to match what will be in the implementation
PR_SENDER_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x5D01001E"
PR_SENT_REPRESENTING_NAME = "http://schemas.microsoft.com/mapi/proptag/0x0042001E"
PR_SUBJECT = "http://schemas.microsoft.com/mapi/proptag/0x0037001E"
PR_ENTRYID = "http://schemas.microsoft.com/mapi/proptag/0x0FFF0102"
PR_MESSAGE_DELIVERY_TIME = "http://schemas.microsoft.com/mapi/proptag/0x0E060040"
PR_HASATTACH = "http://schemas.microsoft.com/mapi/proptag/0x0E1B000B"
PR_UNREAD = "http://schemas.microsoft.com/mapi/proptag/0x0E09000B"

class MockCOMObject(MagicMock):
    """Base mock for COM objects to track attribute access"""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._access_count = 0

    def __getattribute__(self, name):
        if not name.startswith("_") and name not in ["_access_count"]:
            if hasattr(self, "_access_count"):
                self._access_count += 1
        return super().__getattribute__(name)

class MockExchangeUser(MockCOMObject):
    def __init__(self, smtp_address):
        super().__init__()
        self.PrimarySmtpAddress = smtp_address

class MockAddressEntry(MockCOMObject):
    def __init__(self, smtp_address, is_exchange=True):
        super().__init__()
        self._smtp_address = smtp_address
        self._is_exchange = is_exchange

    def GetExchangeUser(self):
        if self._is_exchange:
            return MockExchangeUser(self._smtp_address)
        return None

class MockMailItem(MockCOMObject):
    def __init__(self, i):
        super().__init__()
        self.EntryID = f"EntryID_{i}"
        self.Subject = f"Subject {i}"
        self.SenderName = f"Sender Name {i}"
        self.ReceivedTime = datetime.now()
        self.Unread = False
        self.Attachments = MagicMock()
        self.Attachments.Count = 0

        # Simulate Exchange sender
        self.SenderEmailType = "EX"
        self.Sender = MockAddressEntry(f"sender{i}@example.com")
        self.SenderEmailAddress = f"/O=ORG/OU=EXCHANGE/CN=RECIPIENTS/CN=sender{i}"

class MockTable(MockCOMObject):
    def __init__(self, items_count=0, **kwargs):
        super().__init__(**kwargs)
        self._items_count = items_count
        self._current_index = 0
        self.EndOfTable = False
        # To verify we are using the new implementation
        self._columns_added = []

    def GetNextRow(self):
        if self._current_index >= self._items_count:
            self.EndOfTable = True
            return None
        return MagicMock()

    def GetArray(self, max_rows):
        # Simulate fetching a batch of rows
        count = min(max_rows, self._items_count - self._current_index)
        if count <= 0:
            self.EndOfTable = True # Ensure loop terminates
            return []

        rows = []
        for _ in range(count):
            i = self._current_index
            # Return list of values simulating a row in the array
            # The order MUST match what _get_emails_from_table expects:
            # 1. PR_ENTRYID
            # 2. PR_SUBJECT
            # 3. PR_SENT_REPRESENTING_NAME
            # 4. PR_SENDER_SMTP_ADDRESS
            # 5. PR_MESSAGE_DELIVERY_TIME
            # 6. PR_UNREAD
            # 7. PR_HASATTACH

            # EntryID as bytes (simulating COM behavior roughly, or just string)
            # The code handles both. Let's return string for simplicity as per our previous mock

            rows.append([
                f"EntryID_{i}",
                f"Subject {i}",
                f"Sender Name {i}",
                f"sender{i}@example.com",
                datetime.now(),
                False,
                False
            ])
            self._current_index += 1

        if self._current_index >= self._items_count:
            self.EndOfTable = True

        return rows

    @property
    def Columns(self):
        # Mock Columns.Add
        cols = MagicMock()
        cols.Add = self._columns_added.append
        return cols

class MockFolder(MockCOMObject):
    def __init__(self, items_count):
        super().__init__()
        self._items_count = items_count
        self.Items = MagicMock()
        self.Items.Count = items_count

        # Setup iterator for Items
        self._items_list = [MockMailItem(i) for i in range(items_count)]
        self.Items.__iter__.return_value = iter(self._items_list)

    def GetTable(self, filter_str="", content_table_provider=0):
        # Return a NEW table each time
        return MockTable(self._items_count)

class BenchmarkRunner:
    def __init__(self, num_emails=100):
        self.num_emails = num_emails
        self.folder = MockFolder(num_emails)
        self.bridge = None

    def setup(self):
        # Mock modules before importing bridge
        sys.modules["win32com"] = MagicMock()
        sys.modules["win32com.client"] = MagicMock()

        # Import here to avoid early import issues
        from mailtool.bridge import OutlookBridge
        self.bridge = OutlookBridge(default_account=None)

        # Mock the get_inbox method to return our mock folder
        self.bridge.get_inbox = MagicMock(return_value=self.folder)

    def run_actual_implementation(self):
        """Run the actual list_emails method which should now use GetTable"""
        start_time = time.time()

        # This will call our MockFolder.GetTable()
        emails = self.bridge.list_emails(limit=self.num_emails, folder="Inbox")

        end_time = time.time()
        return end_time - start_time, len(emails)

if __name__ == "__main__":
    try:
        runner = BenchmarkRunner(num_emails=100)
        runner.setup()

        print(f"Benchmarking Actual Implementation with {runner.num_emails} emails...")

        # Reset tracking
        runner.folder = MockFolder(runner.num_emails)
        runner.bridge.get_inbox = MagicMock(return_value=runner.folder)

        # We need to spy on the Table created by GetTable to count ops
        # Since GetTable is called inside the method, we wrap it
        original_get_table = runner.folder.GetTable
        tables_created = []

        def get_table_spy(*args, **kwargs):
            t = original_get_table(*args, **kwargs)
            tables_created.append(t)
            return t

        runner.folder.GetTable = get_table_spy

        duration, count = runner.run_actual_implementation()

        # Verify we actually used the table
        if not tables_created:
            print("ERROR: GetTable was NOT called! Logic might have fallen back to Items.")
        else:
            table = tables_created[0]
            # Verify columns were added
            if len(table._columns_added) < 7:
                 print(f"WARNING: Only {len(table._columns_added)} columns added to table.")

            # Count ops roughly
            # 1 GetTable
            # 7 Columns.Add
            # 10 GetArray calls (for 100 items, batch 50? Wait, batch size is min(limit, 50))
            # 100 / 50 = 2 calls

            ops_count = 1 + len(table._columns_added) + (count // 50 + 1)

            print(f"Actual Implementation: {count} emails processed")
            print(f"  Time: {duration:.6f}s")
            print(f"  Approx COM Ops (Table calls): {ops_count}")
            print(f"  Ops per email: {ops_count / count if count else 0:.2f}")

            # Compare with the theoretical baseline of 1500 ops
            ops_baseline = 1500
            print(f"\nReduction in COM Operations vs Baseline ({ops_baseline}): {(ops_baseline - ops_count) / ops_baseline * 100:.1f}%")

    except Exception as e:
        print(f"Benchmark failed: {e}")
        import traceback
        traceback.print_exc()
