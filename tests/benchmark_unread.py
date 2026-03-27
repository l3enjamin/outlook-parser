import sys
import time
from unittest.mock import MagicMock
from datetime import datetime

class MockCOMObject(MagicMock):
    """Base mock for COM objects to track attribute access"""
    def __init__(self, *args, **kwargs):
        if "name" in kwargs:
            kwargs.pop("name")
        super().__init__(*args, **kwargs)
        self._access_count = 0

    def __getattribute__(self, name):
        if not name.startswith("_") and name not in ["_access_count"]:
            if hasattr(self, "_access_count"):
                self._access_count += 1
        return super().__getattribute__(name)

class MockTable(MockCOMObject):
    def __init__(self, items_count=0, filter_str="", **kwargs):
        super().__init__(**kwargs)
        self._items_count = items_count
        self._filter_str = filter_str
        self._current_index = 0
        self.EndOfTable = False
        self._columns_added = []

    def GetNextRow(self):
        if self._current_index >= self._items_count:
            self.EndOfTable = True
            return None
        return MagicMock()

    def GetArray(self, max_rows):
        count = min(max_rows, self._items_count - self._current_index)
        if count <= 0:
            self.EndOfTable = True
            return []

        rows = []
        for _ in range(count):
            i = self._current_index

            # If a filter is applied (e.g., Unread=True), all returned items should be unread
            # Otherwise, every 10th item is unread
            is_unread = True if self._filter_str else (i % 10 == 0)

            # entry_id, subject, sender_name, sender_smtp, time, unread, hasattach
            rows.append([
                f"EntryID_{i}",
                f"Subject {i}",
                f"Sender Name {i}",
                f"sender{i}@example.com",
                datetime.now(),
                is_unread,
                False
            ])
            self._current_index += 1

        if self._current_index >= self._items_count:
            self.EndOfTable = True

        return rows

    @property
    def Columns(self):
        cols = MagicMock()
        cols.Add = self._columns_added.append
        cols.RemoveAll = MagicMock()
        return cols

class MockFolder(MockCOMObject):
    def __init__(self, items_count=1000, **kwargs):
        super().__init__(**kwargs)
        self._items_count = items_count
        self.Items = MagicMock()
        self.Items.Count = items_count

    def GetTable(self, filter_str="", content_table_provider=0):
        if "Unread" in str(filter_str):
            return MockTable(self._items_count // 10, filter_str=filter_str)
        return MockTable(self._items_count)

    def __bool__(self):
        return True

def setup_mocks(num_emails=1000):
    sys.modules["win32com"] = MagicMock()
    sys.modules["win32com.client"] = MagicMock()

    from mailtool.bridge import OutlookBridge
    bridge = OutlookBridge(default_account=None)

    # We create a real MockFolder rather than having MagicMock auto-create one
    folder = MockFolder(num_emails)
    bridge.get_folder_by_name = MagicMock(return_value=folder)
    bridge.get_inbox = MagicMock(return_value=folder)
    return bridge

def run_baseline(bridge, limit=10, folder_name="TestFolder"):
    start = time.time()
    items = bridge.list_emails(limit=limit * 2, folder=folder_name)
    result = [e for e in items if e["unread"]][:limit]
    end = time.time()
    return end - start, len(result)

def run_optimized(bridge, limit=10, folder_name="TestFolder"):
    start = time.time()
    result = bridge.search_emails(unread=True, limit=limit, folder=folder_name)
    end = time.time()
    return end - start, len(result)

if __name__ == "__main__":
    bridge = setup_mocks(num_emails=1000)

    baseline_time = 0
    optimized_time = 0

    for _ in range(100):
        t, c = run_baseline(bridge, limit=50)
        baseline_time += t

        # Reset mocks
        bridge = setup_mocks(num_emails=1000)
        t, c = run_optimized(bridge, limit=50)
        optimized_time += t
        bridge = setup_mocks(num_emails=1000)

    print(f"Baseline Time (100 runs): {baseline_time:.6f}s")
    print(f"Optimized Time (100 runs): {optimized_time:.6f}s")

    if baseline_time > 0:
        improvement = (baseline_time - optimized_time) / baseline_time * 100
        print(f"Improvement: {improvement:.1f}%")
