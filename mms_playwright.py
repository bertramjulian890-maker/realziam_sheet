"""Liangyu export using a dedicated persistent browser profile."""
from __future__ import annotations

import argparse
import logging
import re
import time
from datetime import datetime
from pathlib import Path

from mms_export import DEFAULT_SOURCE_URL

LOGGER = logging.getLogger("daily_sales.export")
ROOT = Path(__file__).resolve().parent
TITLE = re.compile(r"零售额台账.*\d{8}|零售额台账.*\d{4}-\d{2}-\d{2}")


def export_sales(*, source_url=DEFAULT_SOURCE_URL, download_dir=None,
                 timeout_seconds=180, login=False, headed=False, ca_login=False):
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise RuntimeError("请执行 python -m pip install playwright 和 python -m playwright install chromium") from exc
    folder = Path(download_dir or Path.home() / "Downloads").expanduser().resolve()
    folder.mkdir(parents=True, exist_ok=True)
    diagnostics = ROOT / "logs"
    diagnostics.mkdir(exist_ok=True)
    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            str(ROOT / ".browser-profile"), headless=not (login or headed),
            accept_downloads=True, viewport={"width": 1920, "height": 1080},
        )
        page = context.pages[0] if context.pages else context.new_page()
        page.set_default_timeout(timeout_seconds * 1000)
        try:
            if ca_login:
                from mms_ca_login import login_with_ca
                LOGGER.info("开始本机证书认证")
                result = login_with_ca(context, timeout_seconds=timeout_seconds)
                # The reference does not identify any frontend localStorage token key.
                LOGGER.info("良域登录接口已响应，检查页面会话；响应顶层字段=%s", ",".join(sorted(result)))
            page.goto(source_url, wait_until="domcontentloaded")
            if login:
                print("请在打开的专用浏览器中登录，进入零售额台账并设置导出日期；完成后回到此窗口按 Enter 保存。", flush=True)
                input()
                page.get_by_role("button", name="导出数据", exact=True).wait_for(state="visible", timeout=10000)
                LOGGER.info("登录配置已保存；未触发导出。后续仍可能因登录过期而需要重新登录。")
                return None

            export = page.get_by_role("button", name="导出数据", exact=True)
            try:
                export.wait_for(state="visible")
            except Exception:
                if ca_login:
                    raise RuntimeError("CA 登录后仍未找到导出按钮；接口 Cookie 可能不足以建立页面会话，需核对登录响应及前端令牌存储，详见失败截图") from None
                raise
            history = page.get_by_role("button", name="导出历史", exact=True)
            history.click()
            page.wait_for_timeout(1500)
            titles = page.get_by_text(TITLE)
            previous = set(titles.all_text_contents())
            page.keyboard.press("Escape")
            export.click()
            LOGGER.info("已请求导出，等待本次新增导出记录；保持网页现有日期条件")
            deadline = time.monotonic() + timeout_seconds
            while time.monotonic() < deadline:
                history.click()
                page.wait_for_timeout(1500)
                titles = page.get_by_text(TITLE)
                fresh = [text for text in titles.all_text_contents() if text not in previous]
                if len(fresh) > 1:
                    raise RuntimeError("出现多个新导出记录，无法确定本次文件；请避免同时手动导出")
                if fresh:
                    title = page.get_by_text(fresh[0], exact=True)
                    # Nearest entry containing a download action, rather than an old row.
                    entry = title.locator("xpath=ancestor::*[.//*[normalize-space(text())='下载']][1]")
                    download_link = entry.get_by_text("下载", exact=True)
                    if download_link.count() == 1 and download_link.is_visible():
                        with page.expect_download(timeout=max(1000, int((deadline-time.monotonic())*1000))) as event:
                            download_link.click()
                        download = event.value
                        if download.failure():
                            raise RuntimeError("浏览器下载失败：" + str(download.failure()))
                        name = Path(download.suggested_filename.replace("\\", "/")).name
                        if not name.lower().endswith(".xlsx"):
                            raise RuntimeError("导出文件不是 XLSX：" + name)
                        output = folder / (datetime.now().strftime("%Y%m%d_%H%M%S_%f_") + name)
                        download.save_as(str(output))
                        LOGGER.info("下载完成：%s", output)
                        return output
                page.keyboard.press("Escape")
                page.wait_for_timeout(2000)
            raise TimeoutError("未找到本次新增且可下载的导出记录；需要检查导出历史的页面结构")
        except Exception:
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            screenshot = diagnostics / f"playwright_failure_{stamp}.png"
            try:
                page.screenshot(path=str(screenshot), full_page=True, timeout=5000)
                LOGGER.error("失败页面截图：%s；若显示登录页，请重新执行 --login", screenshot)
            except Exception:
                LOGGER.warning("无法保存失败页面截图")
            raise
        finally:
            context.close()


def main():
    parser = argparse.ArgumentParser(description="Playwright 良域取数测试（不写云表）")
    parser.add_argument("--login", action="store_true")
    parser.add_argument("--headed", action="store_true")
    parser.add_argument("--ca-login", action="store_true", help="通过本机证书客户端自动认证（内网试用）")
    parser.add_argument("--config", default=str(ROOT / "config.json"))
    parser.add_argument("--timeout", type=int, default=180)
    args = parser.parse_args()
    from daily_sales_sheet import load_config
    from run_daily_sales import configure_logging
    configure_logging(ROOT)
    config = load_config(Path(args.config))
    try:
        result = export_sales(source_url=config.get("sourceUrl") or DEFAULT_SOURCE_URL,
                              download_dir=config.get("downloadDir"), login=args.login,
                              headed=args.headed, timeout_seconds=args.timeout, ca_login=args.ca_login)
        if result:
            print(result)
        return 0
    except Exception:
        LOGGER.exception("Playwright 取数失败")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
