from __future__ import annotations

import tempfile
import time
import unittest
from pathlib import Path


class MmsExportTest(unittest.TestCase):
    def test_browser_actions_do_not_change_date_filters(self) -> None:
        from mms_export import POINTS

        self.assertEqual(set(POINTS), {"export_data", "export_history", "download_latest"})

    def test_coordinates_scale_from_reference_resolution(self) -> None:
        from mms_export import scaled_point

        self.assertEqual(scaled_point(1848, 304, 1920, 1080), (1848, 304))
        self.assertEqual(scaled_point(1848, 304, 2560, 1440), (2464, 405))

    def test_new_download_ignores_existing_and_partial_files(self) -> None:
        from mms_export import download_snapshot, find_new_xlsx

        with tempfile.TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            old = folder / "old.xlsx"
            old.write_bytes(b"old")
            before = download_snapshot(folder)
            started = time.time()
            partial = folder / "new.xlsx.crdownload"
            partial.write_bytes(b"partial")
            self.assertIsNone(find_new_xlsx(folder, before, started))
            complete = folder / "new.xlsx"
            complete.write_bytes(b"complete")
            self.assertEqual(find_new_xlsx(folder, before, started), complete)


if __name__ == "__main__":
    unittest.main()
