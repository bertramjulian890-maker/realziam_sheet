from __future__ import annotations

import argparse
import logging
import json
import re
import sys
from calendar import monthrange
from datetime import date, timedelta
from datetime import datetime
from pathlib import Path

LOGGER = logging.getLogger("daily_sales")
DEFAULT_SOURCE_URL = "https://mms.crland.com.cn/bmp/saleData?projectName=汕头万象汇&projectCode=20071&projectId=266"
DEFAULT_API_BASE_URL = "http://10.90.10.66:7055/"
DEFAULT_SPREADSHEET_TOKEN = "shtk99v6k5T9IvGMQjKnUbxZqub"


def configure_logging(project_dir: Path) -> Path:
    log_dir = project_dir / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / f"run_daily_sales_{datetime.now():%Y%m%d_%H%M%S}.log"
    formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(formatter)
    console_handler = logging.StreamHandler(sys.stderr)
    console_handler.setFormatter(formatter)
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    root.addHandler(file_handler)
    root.addHandler(console_handler)
    return log_path


def resolve_business_dates(
    single_date: str | None,
    month_to_yesterday: bool,
    start_date: str | None,
    end_date: str | None,
    month: str | None = None,
    *,
    today: date | None = None,
) -> list[date]:
    if month:
        if single_date or month_to_yesterday or start_date or end_date:
            raise ValueError("--month 不能与其他日期参数同时使用")
        match = re.fullmatch(r"(\d{4})-(\d{2})", month)
        if not match:
            raise ValueError(f"月份格式错误：{month}，需要 YYYY-MM")
        year, month_number = map(int, match.groups())
        try:
            start = date(year, month_number, 1)
            end = date(year, month_number, monthrange(year, month_number)[1])
        except ValueError as exc:
            raise ValueError(f"月份无效：{month}，需要 YYYY-MM") from exc
        return [start + timedelta(days=offset) for offset in range((end - start).days + 1)]
    if single_date and (month_to_yesterday or start_date or end_date):
        raise ValueError("--date 不能与批量日期参数同时使用")
    if month_to_yesterday and (start_date or end_date):
        raise ValueError("--month-to-yesterday 不能与 --start-date/--end-date 同时使用")
    if bool(start_date) != bool(end_date):
        raise ValueError("--start-date 与 --end-date 必须同时填写")
    if single_date:
        return [_parse_date(single_date)]
    if month_to_yesterday:
        current = today or date.today()
        end = current - timedelta(days=1)
        start = end.replace(day=1)
    elif start_date:
        start, end = _parse_date(start_date), _parse_date(end_date)
    else:
        return [date.today() - timedelta(days=1)]
    if end < start:
        raise ValueError("结束日期不能早于开始日期")
    return [start + timedelta(days=offset) for offset in range((end - start).days + 1)]


def _parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value)
    except ValueError as exc:
        raise ValueError(f"日期格式错误：{value}，需要 YYYY-MM-DD") from exc


def main() -> int:
    parser = argparse.ArgumentParser(description="导出销售数据、筛选四个字段并写入润工作电子表格")
    parser.add_argument("--date", help="单个业务日期，默认昨天")
    parser.add_argument("--month-to-yesterday", action="store_true", help="本月 1 日至昨天全部覆盖更新")
    parser.add_argument("--month", help="指定自然月全部覆盖更新，格式 YYYY-MM，例如 2026-09")
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

    project_dir = Path(__file__).resolve().parent
    log_path = configure_logging(project_dir)
    LOGGER.info("启动 run_daily_sales；工作目录=%s；参数=%s", Path.cwd(), sys.argv[1:])
    LOGGER.info("运行日志：%s", log_path)
    try:
        # Import the application modules after logging is active, so missing
        # packages on a fresh machine are recorded instead of ending silently.
        from daily_sales_sheet import filter_export, load_config, sync_file
        from mms_export import export_sales

        config_path = Path(args.config)
        if not config_path.is_absolute():
            config_path = project_dir / config_path
        config = load_config(config_path)
        LOGGER.info("配置文件=%s", config_path)
        business_dates = resolve_business_dates(
            args.date, args.month_to_yesterday, args.start_date, args.end_date, month=args.month
        )
        LOGGER.info("目标日期：%s 至 %s（%d 天）；dry_run=%s",
                    business_dates[0], business_dates[-1], len(business_dates), args.dry_run)
        download_dir = args.download_dir or config.get("downloadDir") or str(Path.home() / "Downloads")
        if args.use_file:
            source_file = Path(args.use_file).expanduser().resolve()
            LOGGER.info("使用指定导出文件：%s", source_file)
        else:
            LOGGER.info("阶段=网页导出；下载目录=%s", download_dir)
            source_file = export_sales(
                business_date=business_dates[-1],
                source_url=config.get("sourceUrl") or DEFAULT_SOURCE_URL,
                download_dir=download_dir,
                pause_for_login=args.pause_for_login,
                timeout_seconds=args.timeout,
            )
        LOGGER.info("导出文件就绪：%s (%d bytes)", source_file, source_file.stat().st_size)

        # Validate every target date before any cloud write, so a partial export
        # cannot produce a partially refreshed month.
        LOGGER.info("阶段=导出完整性检查")
        for business_date in business_dates:
            rows = filter_export(source_file, business_date)
            LOGGER.info("日期=%s 导出记录=%d", business_date, len(rows))

        common = {
            "api_base_url": args.api_base_url or config.get("apiBaseUrl") or DEFAULT_API_BASE_URL,
            "spreadsheet_token": args.spreadsheet_token or config.get("spreadsheetToken") or DEFAULT_SPREADSHEET_TOKEN,
            "sheet_id": args.sheet_id or config.get("sheetId") or "",
            "dry_run": args.dry_run,
        }
        results = []
        for index, business_date in enumerate(business_dates, start=1):
            LOGGER.info("阶段=云表覆盖；进度=%d/%d；日期=%s", index, len(business_dates), business_date)
            result = sync_file(source_file, business_date, **common)
            results.append(result)
            LOGGER.info("日期完成=%s；工作表=%s；列=%s；店铺=%d；清空=%d；合计单元格=%s；合计=%s；校验=%s",
                        business_date, result["sheetId"], result["targetColumn"],
                        result["rowCount"], result["clearedShopRows"], result["totalCell"],
                        result["totalAmount"], result["verified"])
        output = results[0] if len(results) == 1 else {
            "sourceFile": str(source_file.resolve()),
            "startDate": business_dates[0].isoformat(),
            "endDate": business_dates[-1].isoformat(),
            "dayCount": len(results),
            "dryRun": args.dry_run,
            "days": results,
        }
        print(json.dumps(output, ensure_ascii=False, indent=2))
        LOGGER.info("全部完成；处理天数=%d", len(results))
        return 0
    except Exception:
        LOGGER.exception("任务失败")
        print(f"任务失败，详细日志：{log_path}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
