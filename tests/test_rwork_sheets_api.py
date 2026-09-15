from __future__ import annotations

import json
import threading
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse


class Handler(BaseHTTPRequestHandler):
    requests: list[dict] = []

    def log_message(self, _format: str, *_args: object) -> None:
        pass

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        self.__class__.requests.append({"method": "GET", "path": parsed.path, "query": parse_qs(parsed.query)})
        if parsed.path == "/api/sheets/query":
            payload = {"code": 0, "data": {"sheets": [{"sheet_id": "sheet-1", "title": "9月销售"}]}}
        else:
            payload = {"code": 0, "data": {"valueRange": {"values": [["店铺号", "店铺名称", "2026/9/14"]]}}}
        self._send(payload)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        length = int(self.headers.get("Content-Length") or 0)
        body = json.loads(self.rfile.read(length).decode("utf-8"))
        self.__class__.requests.append({"method": "POST", "path": parsed.path, "query": parse_qs(parsed.query), "body": body})
        if isinstance(body.get("valueRange"), list):
            self._send({"code": 500, "message": "写入失败: 'list' object has no attribute 'get'"})
            return
        self._send({"code": 0, "data": {"updatedCells": 1}})

    def _send(self, payload: dict) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


class RworkSheetsApiTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()
        host, port = cls.server.server_address
        cls.base_url = f"http://{host}:{port}"

    @classmethod
    def tearDownClass(cls) -> None:
        cls.server.shutdown()
        cls.server.server_close()
        cls.thread.join(timeout=2)

    def setUp(self) -> None:
        Handler.requests = []

    def test_nonadjacent_cells_do_not_overwrite_gap(self) -> None:
        from rwork_sheets_api import RworkSheetsClient

        RworkSheetsClient(self.base_url).write_values("test", [
            {"range": "sheet-1!C2:C2", "values": [[0]]},
            {"range": "sheet-1!C4:C4", "values": [[45]]},
        ])
        self.assertEqual(
            [r["body"]["valueRange"] for r in Handler.requests],
            [
                {"range": "sheet-1!C2:C2", "values": [[0]]},
                {"range": "sheet-1!C4:C4", "values": [[45]]},
            ],
        )

    def test_uses_documented_query_ranges_and_values_routes(self) -> None:
        from rwork_sheets_api import RworkSheetsClient

        client = RworkSheetsClient(self.base_url)
        self.assertEqual(client.sheet_id("shtk-test"), "sheet-1")
        self.assertEqual(client.read_range("shtk-test", "sheet-1!A1:C1"), [["店铺号", "店铺名称", "2026/9/14"]])
        client.write_values("shtk-test", [
            {"range": "sheet-1!C2:C2", "values": [[120.5]]},
            {"range": "sheet-1!C3:C3", "values": [[45]]},
        ])
        self.assertEqual(Handler.requests[-1], {
            "method": "POST",
            "path": "/api/sheets/values",
            "query": {"spreadsheet_token": ["shtk-test"]},
            "body": {"valueRange": {"range": "sheet-1!C2:C3", "values": [[120.5], [45]]}},
        })
        self.assertEqual(sum(item["method"] == "POST" for item in Handler.requests), 1)


if __name__ == "__main__":
    unittest.main()
