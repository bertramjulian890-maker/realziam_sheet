from __future__ import annotations

import tempfile
import unittest
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook


class DailySalesSheetTest(unittest.TestCase):
    def test_filters_yesterday_and_writes_only_four_columns(self) -> None:
        from daily_sales_sheet import filter_export, save_filtered_workbook

        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "export.xlsx"
            output = Path(temp_dir) / "filtered.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["其他", "店铺号", "店铺名称", "销售日期", "销售总金额"])
            sheet.append(["x", "EIGB101N01", "店铺甲", date(2026, 9, 14), 120.5])
            sheet.append(["y", "DL4101", "店铺乙", "2026/9/15", 400])
            sheet.append(["z", "CA11-005Z001", "店铺丙", "2026-09-14", 0])
            workbook.save(source)

            rows = filter_export(source, date(2026, 9, 14))
            save_filtered_workbook(rows, output)
            result = load_workbook(output, read_only=True, data_only=True)
            values = list(result.active.iter_rows(values_only=True))
            result.close()

        self.assertEqual(rows, [
            {"店铺号": "EIGB101N01", "店铺名称": "店铺甲", "销售日期": "2026-09-14", "销售总金额": 120.5},
            {"店铺号": "CA11-005Z001", "店铺名称": "店铺丙", "销售日期": "2026-09-14", "销售总金额": 0},
        ])
        self.assertEqual(values[0], ("店铺号", "店铺名称", "销售日期", "销售总金额"))
        self.assertEqual(values[1], ("EIGB101N01", "店铺甲", "2026-09-14", 120.5))

    def test_duplicate_shop_for_same_day_is_rejected(self) -> None:
        from daily_sales_sheet import filter_export

        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "export.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["店铺号", "店铺名称", "销售日期", "销售总金额"])
            sheet.append(["DL4101", "甲", "2026/9/14", 100])
            sheet.append(["DL4101", "甲", "2026/9/14", 200])
            workbook.save(source)
            with self.assertRaisesRegex(ValueError, "重复店铺号.*DL4101"):
                filter_export(source, date(2026, 9, 14))

    def test_matches_shop_id_and_existing_date_column(self) -> None:
        from daily_sales_sheet import plan_updates

        headers = ["店铺号", "店铺名称"] + [f"2026/9/{day}" for day in range(1, 20)]
        updates = plan_updates(
            "sheet-1",
            headers,
            ["CA11-005Z001", "DL4101", "EIGB101N01"],
            [{"店铺号": "EIGB101N01", "店铺名称": "甲", "销售日期": "2026-09-14", "销售总金额": 120.5}],
            date(2026, 9, 14),
        )
        self.assertEqual(updates, [{"range": "sheet-1!P4:P4", "values": [[120.5]]}])

    def test_creates_next_date_header_when_missing(self) -> None:
        from daily_sales_sheet import plan_updates

        updates = plan_updates(
            "sheet-1",
            ["店铺号", "店铺名称", "2026/9/14", ""],
            ["DL4101"],
            [{"店铺号": "DL4101", "店铺名称": "甲", "销售日期": "2026-09-15", "销售总金额": 45}],
            date(2026, 9, 15),
        )
        self.assertEqual(updates, [
            {"range": "sheet-1!D1:D1", "values": [["2026/9/15"]]},
            {"range": "sheet-1!D2:D2", "values": [[45]]},
        ])


if __name__ == "__main__":
    unittest.main()
