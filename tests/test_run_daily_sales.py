from __future__ import annotations

import unittest
from datetime import date


class RunDailySalesTest(unittest.TestCase):
    def test_month_to_yesterday_uses_first_day_through_previous_day(self) -> None:
        from run_daily_sales import resolve_business_dates

        self.assertEqual(
            resolve_business_dates(None, True, None, None, today=date(2026, 9, 22)),
            [date(2026, 9, day) for day in range(1, 22)],
        )

    def test_explicit_range_is_inclusive(self) -> None:
        from run_daily_sales import resolve_business_dates

        self.assertEqual(
            resolve_business_dates(None, False, "2026-09-01", "2026-09-03"),
            [date(2026, 9, 1), date(2026, 9, 2), date(2026, 9, 3)],
        )


if __name__ == "__main__":
    unittest.main()
