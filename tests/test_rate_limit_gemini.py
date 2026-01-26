import sys
from unittest.mock import MagicMock

# Mock google.generativeai before importing the module under test
# This is necessary because the environment might not have the correct package installed.
# We only mock the specific module that is missing.
sys.modules["google.generativeai"] = MagicMock()

import unittest
from unittest.mock import patch
import time
from collections import deque
import book_generator.llm_api

class TestGeminiRateLimit(unittest.TestCase):
    def setUp(self):
        if hasattr(book_generator.llm_api, "_GEMINI_CALL_TIMESTAMPS"):
             book_generator.llm_api._GEMINI_CALL_TIMESTAMPS.clear()

    @patch("time.sleep")
    @patch("time.time")
    def test_rate_limit_flow(self, mock_time, mock_sleep):
        # Check if the function exists
        if not hasattr(book_generator.llm_api, "_enforce_gemini_rate_limit"):
             self.fail("Function _enforce_gemini_rate_limit not implemented yet.")

        from book_generator.llm_api import _enforce_gemini_rate_limit, _GEMINI_CALL_TIMESTAMPS
        _GEMINI_CALL_TIMESTAMPS.clear()

        # fill up 29 slots
        t0 = 1000.0
        mock_time.return_value = t0
        for _ in range(29):
            _enforce_gemini_rate_limit()

        # Call 30:
        # 1. current_time check -> 1005.0
        # 2. (Sleep happens)
        # 3. post-sleep timestamp -> 1061.0

        mock_time.side_effect = [1005.0, 1061.0]

        _enforce_gemini_rate_limit()

        # Check sleep
        # Oldest is 1000.0. Wait = 1000 + 60 - 1005 = 55.
        mock_sleep.assert_called_with(55.0)

        # Check deque
        # Since we slept past the expiry of all previous timestamps (1000.0 vs 1061.0 - 60 = 1001.0),
        # all previous 29 timestamps are removed. Only the new one remains.
        self.assertEqual(len(_GEMINI_CALL_TIMESTAMPS), 1)
        # The only one should be the post-sleep time
        self.assertEqual(_GEMINI_CALL_TIMESTAMPS[-1], 1061.0)

        # 31st call
        mock_time.side_effect = [1062.0]
        _enforce_gemini_rate_limit()

        # Should not sleep
        # mock_sleep was called once before.
        self.assertEqual(mock_sleep.call_count, 1)

if __name__ == "__main__":
    unittest.main()
