from __future__ import annotations

import json
import re
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen


class RworkSheetsApiError(RuntimeError):
    pass


class RworkSheetsClient:
    def __init__(self, base_url: str, timeout_seconds: int = 60) -> None:
        self.base_url = base_url.rstrip("/")
        self.timeout_seconds = timeout_seconds

    def sheet_id(self, spreadsheet_token: str) -> str:
        payload = self._request(
            "GET", "/api/sheets/query", {"spreadsheet_token": spreadsheet_token}
        )
        data = payload.get("data")
        candidates: list[dict[str, Any]] = []
        if isinstance(data, list):
            candidates = [item for item in data if isinstance(item, dict)]
        elif isinstance(data, dict):
            for key in ("sheets", "sheet_info", "sheetInfo", "items", "worksheets"):
                value = data.get(key)
                if isinstance(value, list):
                    candidates = [item for item in value if isinstance(item, dict)]
                    break
            if not candidates and any(key in data for key in ("sheet_id", "sheetId", "id")):
                candidates = [data]

        ids = [
            str(item.get("sheet_id") or item.get("sheetId") or item.get("id") or "").strip()
            for item in candidates
        ]
        ids = [value for value in ids if value]
        if len(ids) == 1:
            return ids[0]
        if not ids:
            raise RworkSheetsApiError("电子表格未返回可用的工作表 ID")
        raise RworkSheetsApiError("电子表格包含多个工作表，请通过 --sheet-id 明确指定")

    def read_range(self, spreadsheet_token: str, cell_range: str) -> list[list[Any]]:
        payload = self._request(
            "GET",
            "/api/sheets/ranges",
            {"spreadsheet_token": spreadsheet_token, "range": cell_range},
        )
        data = payload.get("data")
        if isinstance(data, dict):
            value_range = data.get("valueRange") or data.get("value_range")
            if isinstance(value_range, dict) and isinstance(value_range.get("values"), list):
                return value_range["values"]
            if isinstance(data.get("values"), list):
                return data["values"]
        raise RworkSheetsApiError(f"读取范围 {cell_range} 时响应中没有 values")

    def write_values(self, spreadsheet_token: str, value_ranges: list[dict[str, Any]]) -> dict[str, Any]:
        # The endpoint accepts one valueRange object, not a batch array.
        # Combine adjacent cells only; never fill gaps with empty values.
        blocks: list[dict[str, Any]] = []
        previous = None
        for item in value_ranges:
            match = re.fullmatch(r"(.+)!([A-Z]+)(\d+):\2(\d+)", item["range"])
            if (match and previous and (match[1], match[2]) == (previous[1], previous[2])
                    and int(match[3]) == int(previous[4]) + 1):
                blocks[-1]["range"] = blocks[-1]["range"].split(":")[0] + f":{match[2]}{match[4]}"
                blocks[-1]["values"].extend(item["values"])
            else:
                blocks.append({"range": item["range"], "values": list(item["values"])})
            previous = match
        result: dict[str, Any] = {}
        for block in blocks:
            result = self._request(
                "POST", "/api/sheets/values",
                {"spreadsheet_token": spreadsheet_token},
                {"valueRange": block},
            )
        return result

    def _request(
        self,
        method: str,
        path: str,
        query: dict[str, str],
        body: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        url = f"{self.base_url}{path}?{urlencode(query)}"
        raw_body = None if body is None else json.dumps(body, ensure_ascii=False).encode("utf-8")
        request = Request(
            url,
            data=raw_body,
            method=method,
            headers={"Content-Type": "application/json", "Accept": "application/json"},
        )
        try:
            with urlopen(request, timeout=self.timeout_seconds) as response:
                raw = response.read().decode("utf-8", errors="replace")
        except HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")
            raise RworkSheetsApiError(f"HTTP {exc.code}: {detail}") from exc
        except URLError as exc:
            raise RworkSheetsApiError(f"无法连接润工作接口：{exc.reason}") from exc

        try:
            payload = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise RworkSheetsApiError(f"接口返回的不是 JSON：{raw[:200]}") from exc
        if not isinstance(payload, dict):
            raise RworkSheetsApiError("接口响应格式异常")
        code = payload.get("code")
        if code not in (None, 0, "0", 200, "200"):
            raise RworkSheetsApiError(str(payload.get("message") or payload))
        return payload
