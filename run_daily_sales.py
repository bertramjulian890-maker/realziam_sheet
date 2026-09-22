from __future__ import annotations

import argparse
import json
from datetime import date, timedelta
from pathlib import Path

from daily_sales_sheet import (
    DEFAULT_API_BASE_URL,
    DEFAULT_SPREADSHEET_TOKEN,
    load_config,
    parse_business_date,
    filter_export,
    sync_file,
)
from mms_export import DEFAULT_SOURCE_URL, export_sales


def resolve_business_dates(
    single_date: str | None,
    month_to_yesterday: bool,
    start_date: str | None,
    end_date: str | None,
    *,
    today: date | None = None,
) -> list[date]:
    if single_date and (month_to_yesterday or start_date or end_date):
        raise ValueError("--date 不能与批量日期参数同时使用")
    if month_to_yesterday and (start_date or end_date):
        raise ValueError("--month-to-yesterday 不能与 --start-date/--end-date 同时使用")
    if bool(start_date) != bool(end_date):
        raise ValueError("--start-date 与 --end-date 必须同时填写")
    if single_date:
        return [parse_business_date(single_date)]
    if month_to_yesterday:
        current = today or date.today()
        end = current - timedelta(days=1)
        start = end.replace(day=1)
    elif start_date:
        start, end = parse_business_date(start_date), parse_business_date(end_date)
    else:
        return [parse_business_date(None)]
    if end < start:
        raise ValueError("结束日期不能早于开始日期")
    return [start + timedelta(days=offset) for offset in range((end - start).days + 1)]


def main() -> int:
    parser = argparse.ArgumentParser(description="导出昨日销售、筛选四个字段并写入润工作电子表格")
    parser.add_argument("--date", help="单个业务日期，默认昨天")
    parser.add_argument("--month-to-yesterday", action="store_true", help="本月 1 日至昨天全部覆盖更新")
    parser.add_argument("--start-date", help="批量开始日期 YYYY-MM-DD")
    parser.add_argument("--end-date", help="批量结束日期 YYYY-MM-DD")
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--use-file", help="跳过鼠标导出，直接使用已有 XLSX（用于测试或补跑）")
    parser.add_argument("--download-dir")
    parser.add_argument("--pause-for-login", action="store_true", help="兼容旧命令；现在自动继续")
    parser.add_argument("--api-base-url")
    parser.add_argument("--spreadsheet-token")
    parser.add_argument("--sheet-id")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--timeout", type=int, default=180)
    args = parser.parse_args()

    config = load_config(args.config)
    business_dates = resolve_business_dates(
        args.date, args.month_to_yesterday, args.start_date, args.end_date
    )
    download_dir = args.download_dir or config.get("downloadDir") or str(Path.home() / "Downloads")
    source_file = Path(args.use_file) if args.use_file else export_sales(
        business_date=business_dates[-1],
        source_url=config.get("sourceUrl") or DEFAULT_SOURCE_URL,
        download_dir=download_dir,
        pause_for_login=args.pause_for_login,
        timeout_seconds=args.timeout,
    )
    # Validate every target date before any cloud write, so a partial export
    # cannot produce a partially refreshed month.
    for business_date in business_dates:
        filter_export(source_file, business_date)

    common = {
        "api_base_url": args.api_base_url or config.get("apiBaseUrl") or DEFAULT_API_BASE_URL,
        "spreadsheet_token": args.spreadsheet_token or config.get("spreadsheetToken") or DEFAULT_SPREADSHEET_TOKEN,
        "sheet_id": args.sheet_id or config.get("sheetId") or "",
        "dry_run": args.dry_run,
    }
    results = [sync_file(source_file, business_date, **common) for business_date in business_dates]
    output = results[0] if len(results) == 1 else {
        "sourceFile": str(source_file.resolve()),
        "startDate": business_dates[0].isoformat(),
        "endDate": business_dates[-1].isoformat(),
        "dayCount": len(results),
        "dryRun": args.dry_run,
        "days": results,
    }
    print(json.dumps(output, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
