from __future__ import annotations

import argparse
import json
from datetime import date
from pathlib import Path

from daily_sales_sheet import (
    DEFAULT_API_BASE_URL,
    DEFAULT_SPREADSHEET_TOKEN,
    load_config,
    parse_business_date,
    sync_file,
)
from mms_export import DEFAULT_SOURCE_URL, export_sales


def main() -> int:
    parser = argparse.ArgumentParser(description="导出昨日销售、筛选四个字段并写入润工作电子表格")
    parser.add_argument("--date", help="业务日期，默认昨天")
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--use-file", help="跳过鼠标导出，直接使用已有 XLSX（用于测试或补跑）")
    parser.add_argument("--download-dir")
    parser.add_argument("--pause-for-login", action="store_true")
    parser.add_argument("--api-base-url")
    parser.add_argument("--spreadsheet-token")
    parser.add_argument("--sheet-id")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--timeout", type=int, default=180)
    args = parser.parse_args()

    config = load_config(args.config)
    business_date: date = parse_business_date(args.date)
    download_dir = args.download_dir or config.get("downloadDir") or str(Path.home() / "Downloads")
    source_file = Path(args.use_file) if args.use_file else export_sales(
        business_date=business_date,
        source_url=config.get("sourceUrl") or DEFAULT_SOURCE_URL,
        download_dir=download_dir,
        pause_for_login=args.pause_for_login,
        timeout_seconds=args.timeout,
    )
    result = sync_file(
        source_file,
        business_date,
        api_base_url=args.api_base_url or config.get("apiBaseUrl") or DEFAULT_API_BASE_URL,
        spreadsheet_token=args.spreadsheet_token or config.get("spreadsheetToken") or DEFAULT_SPREADSHEET_TOKEN,
        sheet_id=args.sheet_id or config.get("sheetId") or "",
        dry_run=args.dry_run,
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
