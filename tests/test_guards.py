from __future__ import annotations

import time
import unittest

from mcp_office_writer.guards import with_timeout


class GuardsTimeoutTest(unittest.TestCase):
    def test_returns_value_before_timeout(self) -> None:
        self.assertEqual(with_timeout(lambda: 42, 5, -1), 42)

    def test_returns_default_on_timeout(self) -> None:
        def slow() -> int:
            time.sleep(2)
            return 1

        started = time.monotonic()
        result = with_timeout(slow, 0.2, -1)
        elapsed = time.monotonic() - started
        self.assertEqual(result, -1)
        self.assertLess(elapsed, 1.5, "超时后应立即返回默认值，不等待后台任务")

    def test_propagates_function_exception(self) -> None:
        def boom() -> int:
            raise ValueError("boom")

        with self.assertRaisesRegex(ValueError, "boom"):
            with_timeout(boom, 5, -1)

    def test_passes_arguments_through(self) -> None:
        def add(a: int, b: int) -> int:
            return a + b

        self.assertEqual(with_timeout(add, 5, -1, 1, 2), 3)


if __name__ == "__main__":
    unittest.main()
