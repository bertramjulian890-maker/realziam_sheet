import io
import contextlib
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from openpyxl import Workbook

import mms_api_export as api


def reply(data, code="0", status=200):
    return Mock(status_code=status, json=Mock(return_value={"code": code, "data": data}))


def task(identity, state="FINISH"):
    return {"businessId": identity, "projectNo": "20058", "taskType": "PROJECT_SALES_EXPORT",
            "subTaskType": "DATE_TOTAL", "taskRangeBeginTime": "2026-10-01",
            "taskRangeEndTime": "2026-10-23", "taskState": state,
            "downloadUrl": "https://storagecloud-api.saas.crland.com.cn/api.action?id=test"}


class ApiExportTest(unittest.TestCase):
    def test_config_credentials_work_without_set_environment(self):
        with patch.dict("os.environ", {}, clear=True), \
             patch("requests.Session") as factory, \
             patch.object(api, "login_with_ca", return_value={"data": {"token": "test-token"}}) as login:
            session = api.create_session("20071", api.DEFAULT_POSITION,
                                         auth={"caName": "example", "caPin": "pin", "certificateIndex": 1})
        self.assertIs(session, factory.return_value)
        self.assertEqual(login.call_args.kwargs["name"], "EXAMPLE")
        self.assertEqual(login.call_args.kwargs["pin"], "pin")
        self.assertEqual(login.call_args.kwargs["identity_index"], 1)
        self.assertEqual(session.headers.update.call_args.args[0]["mixc-position-id"], api.DEFAULT_POSITION)

    def test_defaults_use_wanxianghui_position(self):
        with patch.dict("os.environ", {}, clear=True):
            self.assertEqual(api.project_settings({}), ("20071", "cd1f2432c46b4dbeac7b13ef4e1aaf87"))

    def test_other_project_requires_its_own_position(self):
        with patch.dict("os.environ", {}, clear=True):
            with self.assertRaisesRegex(ValueError, "positionId"):
                api.project_settings({"mmsApi": {"projectCode": "20058"}})

    def test_monthly_entry_passes_date_range_and_stops_before_cloud_on_export_failure(self):
        import run_daily_sales

        config = {"sourceUrl": "https://mms.crland.com.cn/?projectCode=20071",
                  "mmsApi": {"positionId": "position"},
                  "spreadsheetToken": "test-token",
                  "monthlySheetIds": {"2026-09": "test-sheet"}}
        with patch("sys.argv", ["run_daily_sales.py", "--export-engine", "api", "--month", "2026-09", "--dry-run"]), \
             patch.object(run_daily_sales, "configure_logging", return_value=Path("test.log")), \
             patch("daily_sales_sheet.load_config", return_value=config), \
             patch.object(api, "export_sales", side_effect=RuntimeError("export failed")) as export, \
             patch("daily_sales_sheet.sync_file") as sync, \
             self.assertLogs("daily_sales", level="ERROR"), \
             contextlib.redirect_stderr(io.StringIO()):
            self.assertEqual(run_daily_sales.main(), 1)
        self.assertEqual(str(export.call_args.kwargs["start_date"]), "2026-09-01")
        self.assertEqual(str(export.call_args.kwargs["end_date"]), "2026-09-30")
        self.assertEqual(export.call_args.kwargs["project_code"], "20071")
        sync.assert_not_called()

    def test_only_new_task_is_signed_and_downloaded(self):
        session = Mock()
        old = task("old")
        session.request.side_effect = [reply({"records": [old], "pages": 1}), reply(True),
                                      reply({"records": [task("new", "GENERATING"), old], "pages": 1}),
                                      reply({"records": [task("new"), old], "pages": 1}),
                                      reply(old["downloadUrl"] + "&f_token=fresh")]
        with patch.object(api, "download_file", return_value=Path("new.xlsx")) as download:
            result = api.export_sales(project_code="20058", position_id="position",
                                     start_date="2026-10-01", end_date="2026-10-23",
                                     download_dir="downloads", session=session, poll_seconds=0)
        self.assertEqual(result, Path("new.xlsx"))
        calls = session.request.call_args_list
        self.assertEqual(calls[1].kwargs["json"]["dealDateFrom"], "2026-10-01")
        self.assertEqual(calls[-1].kwargs["params"], {"url": old["downloadUrl"]})
        self.assertIn("f_token=fresh", download.call_args.args[0])

    def test_ambiguous_new_tasks_stop_without_downloading(self):
        session = Mock()
        session.request.side_effect = [reply({"records": [], "pages": 1}), reply(True),
                                      reply({"records": [task("one"), task("two")], "pages": 1})]
        with patch.object(api, "download_file") as download:
            with self.assertRaisesRegex(RuntimeError, "多条"):
                api.export_sales(project_code="20058", position_id="p", start_date="2026-10-01",
                                 end_date="2026-10-23", download_dir="downloads", session=session)
        download.assert_not_called()

    def test_business_failure_stops(self):
        session = Mock()
        session.request.return_value = reply(None, code="30000")
        with self.assertRaisesRegex(RuntimeError, "业务响应失败"):
            api.records(session, "20058")

    def test_records_collect_all_pages(self):
        session = Mock()
        session.request.side_effect = [reply({"records": [task("one")], "pages": 2}),
                                      reply({"records": [task("two")], "pages": 2})]
        self.assertEqual(len(api.records(session, "20058")), 2)
        self.assertEqual(session.request.call_args.kwargs["params"]["pageIndex"], 2)

    def test_wrong_project_config_stops(self):
        with self.assertRaisesRegex(ValueError, "不一致"):
            api.project_settings({"sourceUrl": "https://mms.crland.com.cn/?projectCode=20071",
                                  "mmsApi": {"projectCode": "20058", "positionId": "p"}})

    def test_download_rejects_error_body_and_cleans_partial_file(self):
        with tempfile.TemporaryDirectory() as directory, patch("requests.get") as get:
            response = get.return_value.__enter__.return_value
            response.status_code = 200
            response.iter_content.return_value = [b'{"code":30000}']
            with self.assertRaisesRegex(RuntimeError, "有效销售 XLSX"):
                api.download_file(task("x")["downloadUrl"], directory, "2026-10-01", "2026-10-23")
            self.assertEqual(list(Path(directory).iterdir()), [])
            self.assertNotIn("headers", get.call_args.kwargs)

    def test_download_accepts_sales_workbook(self):
        workbook = Workbook()
        workbook.active.append(["店铺号", "店铺名称", "销售日期", "销售总金额"])
        buffer = io.BytesIO()
        workbook.save(buffer)
        workbook.close()
        with tempfile.TemporaryDirectory() as directory, patch("requests.get") as get:
            response = get.return_value.__enter__.return_value
            response.status_code = 200
            response.iter_content.return_value = [buffer.getvalue()]
            output = api.download_file(task("x")["downloadUrl"], directory, "2026-10-01", "2026-10-23")
            self.assertTrue(output.is_file())
            self.assertEqual(len(list(Path(directory).iterdir())), 1)
