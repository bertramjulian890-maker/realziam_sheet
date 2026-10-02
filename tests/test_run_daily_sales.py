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

    def test_explicit_month_uses_every_day_of_that_month(self) -> None:
        from run_daily_sales import resolve_business_dates

        dates = resolve_business_dates(None, False, None, None, month="2026-09")

        self.assertEqual(dates[0], date(2026, 9, 1))
        self.assertEqual(dates[-1], date(2026, 9, 30))
        self.assertEqual(len(dates), 30)

    def test_explicit_month_handles_leap_year(self) -> None:
        from run_daily_sales import resolve_business_dates

        dates = resolve_business_dates(None, False, None, None, month="2024-02")

        self.assertEqual(dates[-1], date(2024, 2, 29))
        self.assertEqual(len(dates), 29)

    def test_explicit_month_cannot_mix_with_other_date_modes(self) -> None:
        from run_daily_sales import resolve_business_dates

        with self.assertRaisesRegex(ValueError, "--month 不能"):
            resolve_business_dates(None, True, None, None, month="2026-09")

    def test_explicit_month_requires_year_and_month_format(self) -> None:
        from run_daily_sales import resolve_business_dates

        with self.assertRaisesRegex(ValueError, "YYYY-MM"):
            resolve_business_dates(None, False, None, None, month="2026-9")


if __name__ == "__main__":
    unittest.main()
