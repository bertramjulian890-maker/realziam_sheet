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
        end = (today or date.today()) - timedelta(days=1)
        start = end.replace(day=1)
    if end < start:
        raise ValueError("结束日期不能早于开始日期")
    return [start + timedelta(days=offset) for offset in range((end - start).days + 1)]


def _parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value)
    except ValueError as exc:
        raise ValueError(f"日期格式错误：{value}，需要 YYYY-MM-DD") from exc


def failure_guidance(error: Exception) -> str:
    """Translate common operational failures into an actionable next step."""
    message = str(error)
    if "monthlySheetIds" in message:
        return "打开 config.json，在 monthlySheetIds 中填写报错月份对应的工作表 ID；其他未用月份可留空。"
    if "spreadsheetToken" in message:
        return "打开 config.json，填写共用电子表格 spreadsheetToken（链接 /sheets/ 后的部分）。"
    if "caName" in message or "caPin" in message or "MMS_CA_NAME" in message:
        return "打开 config.json，检查 mmsApi.caName 和 mmsApi.caPin 是否已填写。"
    if "证书客户端" in message or "wss://127.0.0.1:10022" in message:
        return "确认已登录 Windows，启动本机 CA/证书客户端并插好证书，再双击重试。"
    if "证书枚举" in message or "证书序号" in message or "证书签名" in message:
        return "检查证书是否可用、PIN 是否正确；若电脑有多个证书，可核对 config.json 的 mmsApi.certificateIndex。"
    if "CA 登录失败" in message or "认证中心" in message or "认证回调" in message or "良域登录" in message:
        return "证书步骤已开始但登录没有完成。检查内网、账号权限和本机证书客户端；完整阶段及 HTTP 状态见日志。"
    if "requirements-api.txt" in message or "No module named" in message:
        return "在项目文件夹打开 CMD，运行：python -m pip install -r requirements-api.txt，然后重试。"
    if "导出任务" in message or "导出记录" in message or "等待本次导出" in message:
        return "检查良域内网连接、账号的万象汇项目权限，以及导出任务是否完成；确认后再重试。"
    if "下载" in message or "XLSX" in message or "导出表" in message:
        return "检查下载文件是否完整、是否包含目标日期和销售字段；完整日志会指出失败阶段。"
    if "润工作" in message or "云表" in message or "工作表" in message or "写入后" in message or "回读校验" in message:
        return "检查润工作内网接口是否可访问、config.json 的表格令牌/工作表 ID 是否正确，再重试。"
    if isinstance(error, (json.JSONDecodeError, FileNotFoundError)) or "config.json" in message:
        return "检查项目目录是否存在 config.json，且内容是有效 JSON；可参考 config.example.json 补齐配置。"
    return "查看下方日志路径中的完整错误详情；日志按阶段记录问题，可把最后的 ERROR 和 traceback 发给维护人员。"


def main() -> int:
    parser = argparse.ArgumentParser(description="导出销售数据、筛选四个字段并写入润工作电子表格")
    parser.add_argument("--date", help="指定日期只处理当天；未指定日期范围时默认月度刷新")
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
    parser.add_argument("--export-engine", choices=["pyautogui", "playwright", "api"], help="取数方式；api 为纯接口取数")
    parser.add_argument("--headed", action="store_true", help="Playwright 显示浏览器，便于调试")
    parser.add_argument("--ca-login", action="store_true", help="Playwright 使用本机证书认证")
    args = parser.parse_args()

    project_dir = Path(__file__).resolve().parent
    log_path = configure_logging(project_dir)
    LOGGER.info("启动 run_daily_sales；工作目录=%s；参数=%s", Path.cwd(), sys.argv[1:])
    LOGGER.info("运行日志：%s", log_path)
    try:
        # Import the application modules after logging is active, so missing
        # packages on a fresh machine are recorded instead of ending silently.
        from daily_sales_sheet import filter_export, load_config, spreadsheet_token_for_month, sheet_id_for_month, sync_file
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
        months = sorted({business_date.strftime("%Y-%m") for business_date in business_dates})
        token_override = (args.spreadsheet_token or "").strip()
        tokens_by_month = {
            month: token_override or spreadsheet_token_for_month(config, month)
            for month in months
        }
        sheet_override = (args.sheet_id or "").strip()
        sheets_by_month = {
            month: sheet_override or sheet_id_for_month(config, month)
            for month in months
        }
        LOGGER.info("目标数据表月份：%s", "、".join(months))
        download_dir = args.download_dir or config.get("downloadDir") or str(Path.home() / "Downloads")
        if args.use_file:
            source_file = Path(args.use_file).expanduser().resolve()
            LOGGER.info("使用指定导出文件：%s", source_file)
        else:
            LOGGER.info("阶段=取数；下载目录=%s", download_dir)
            engine = args.export_engine or config.get("exportEngine", "api")
            if engine not in {"pyautogui", "playwright", "api"}:
                raise ValueError("exportEngine 只能是 pyautogui、playwright 或 api")
            if args.ca_login and engine not in {"playwright", "api"}:
                raise ValueError("--ca-login 需要选择 playwright 或 api")
            options = dict(
                source_url=config.get("sourceUrl") or DEFAULT_SOURCE_URL,
                download_dir=download_dir,
                timeout_seconds=args.timeout,
            )
            LOGGER.info("取数方式=%s", engine)
            if engine == "api":
                from mms_api_export import export_sales as api_export, project_settings
                project, position = project_settings({**config, "sourceUrl": config.get("sourceUrl") or DEFAULT_SOURCE_URL})
                source_file = api_export(project_code=project, position_id=position,
                                         start_date=business_dates[0], end_date=business_dates[-1],
                                         download_dir=download_dir, timeout_seconds=args.timeout,
                                         auth=config.get("mmsApi", {}))
            elif engine == "playwright":
                from mms_playwright import export_sales as browser_export
                source_file = browser_export(**options, headed=args.headed, ca_login=args.ca_login)
            else:
                source_file = export_sales(**options, business_date=business_dates[-1],
                                          pause_for_login=args.pause_for_login)
        LOGGER.info("导出文件就绪：%s (%d bytes)", source_file, source_file.stat().st_size)

        # Validate every target date before any cloud write, so a partial export
        # cannot produce a partially refreshed month.
        LOGGER.info("阶段=导出完整性检查")
        for business_date in business_dates:
            rows = filter_export(source_file, business_date)
            LOGGER.info("日期=%s 导出记录=%d", business_date, len(rows))

        common = {
            "api_base_url": args.api_base_url or config.get("apiBaseUrl") or DEFAULT_API_BASE_URL,
            "dry_run": args.dry_run,
        }
        results = []
        for index, business_date in enumerate(business_dates, start=1):
            LOGGER.info("阶段=云表覆盖；进度=%d/%d；日期=%s", index, len(business_dates), business_date)
            result = sync_file(source_file, business_date,
                               spreadsheet_token=tokens_by_month[business_date.strftime("%Y-%m")],
                               sheet_id=sheets_by_month[business_date.strftime("%Y-%m")],
                               **common)
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
    except Exception as exc:
        LOGGER.exception("任务失败")
        guidance = failure_guidance(exc)
        LOGGER.error("处理建议：%s", guidance)
        print(f"\n处理建议：{guidance}", file=sys.stderr)
        print(f"任务失败，详细日志：{log_path}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
