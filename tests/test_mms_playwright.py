import tempfile
import importlib.util
import unittest
from pathlib import Path
from unittest.mock import patch
from urllib.parse import quote

import mms_playwright


@unittest.skipUnless(importlib.util.find_spec('playwright'), '需要可选 Playwright 依赖及 Chromium')
class BrowserExportTest(unittest.TestCase):
    def test_new_export_is_downloaded_instead_of_old_record(self):
        html = '''<button onclick="document.querySelector('#entries').insertAdjacentHTML('afterbegin',
          '<div><span>零售额台账__2026-10-01至2026-10-15_120001</span><a download=latest.xlsx href=data:application/octet-stream;base64,bmV3>下载</a></div>')">导出数据</button>
          <button>导出历史</button><div id=entries><div>
          <span>零售额台账__2026-10-01至2026-10-15_100001</span>
          <a download=old.xlsx href=data:application/octet-stream;base64,b2xk>下载</a></div></div>'''
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            with patch.object(mms_playwright, 'ROOT', root):
                result = mms_playwright.export_sales(
                    source_url='data:text/html;charset=utf-8,' + quote(html),
                    download_dir=root / 'downloads', timeout_seconds=15)
            self.assertTrue(result.name.endswith('latest.xlsx'))
            self.assertEqual(result.read_bytes(), b'new')


if __name__ == '__main__':
    unittest.main()
