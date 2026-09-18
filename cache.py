"""A small in-memory cache for results worked out from the same inputs (Task 17).

One run used to clean each workbook five times (the summary table, the Excel file, the AI reuse
check, the deck and the memo) and work out its metrics up to eight times. clean.py and metrics.py
now keep their results here, keyed on a hash of their inputs, so a run does that work once.

Three rules keep it from ever changing an output:
- the key is a hash of the inputs themselves, never a file name: edited inputs miss the cache;
- only results are kept, never errors: a problem is found again every time, with today's message;
- every result is copied on the way in and on the way out, so a caller that changes a table it
  was given can't change what the next caller gets.
It keeps at most MAX_ENTRIES results, dropping the one used longest ago, because the web page
runs for days and every upload is a new key.
"""

import copy
import threading
from collections import OrderedDict

MAX_ENTRIES = 32   # a batch of 3 companies needs 3 per cache; 32 leaves room for a day of uploads


class ResultCache:
    """Remembers up to max_entries results by key, and hands out copies of them."""

    def __init__(self, max_entries=MAX_ENTRIES):
        self.max_entries = max_entries
        self._results = OrderedDict()     # oldest use first
        self._lock = threading.Lock()     # main.py --workers runs companies in threads

    def get(self, key, work_out):
        """The result saved for key, or work_out()'s result, saved for next time. Always a copy."""
        with self._lock:
            if key in self._results:
                self._results.move_to_end(key)   # just used: the last to be dropped
                return copy.deepcopy(self._results[key])
        result = work_out()   # outside the lock: two threads may both work it out, which is harmless
        with self._lock:
            self._results[key] = copy.deepcopy(result)
            while len(self._results) > self.max_entries:
                self._results.popitem(last=False)   # drop the one used longest ago
        return result

    def clear(self):
        """Forget everything (benchmark.py and the tests start each run from nothing)."""
        with self._lock:
            self._results.clear()

    def __len__(self):
        return len(self._results)
