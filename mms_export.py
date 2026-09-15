from __future__ import annotations

import argparse
import os
import time
import webbrowser
from datetime import date, timedelta
from pathlib import Path


REFERENCE_SIZE = (1920, 1080)
DEFAULT_SOURCE_URL = "https://mms.crland.com.cn/bmp/saleData?projectName=汕头万象汇&projectCode=20071&projectId=266"
POINTS = {
    "export_data": (1848, 304),
    "export_history": (1745, 304),
    "download_latest": (1427, 402),
}


def scaled_point(x: int, y: int, screen_width: int, screen_height: int) -> tuple[int, int]:
    return round(x * screen_width / REFERENCE_SIZE[0]), round(y * screen_height / REFERENCE_SIZE[1])


def download_snapshot(folder: str | Path) -> dict[Path, tuple[int, int]]:
    directory = Path(folder)
    directory.mkdir(parents=True, exist_ok=True)
    return {
        path.resolve(): (path.stat().st_mtime_ns, path.stat().st_size)
        for path in directory.glob("*.xlsx")
        if path.is_file()
    }


def find_new_xlsx(folder: str | Path, before: dict[Path, tuple[int, int]], started_at: float) -> Path | None:
    candidates: list[Path] = []
    for path in Path(folder).glob("*.xlsx"):
        if not path.is_file() or path.name.startswith("~$"):
            continue
        resolved = path.resolve()
        stat = path.stat()
        signature = (stat.st_mtime_ns, stat.st_size)
        if before.get(resolved) == signature:
            continue
        if stat.st_mtime < started_at - 1:
            continue
        candidates.append(path)
    return max(candidates, key=lambda item: item.stat().st_mtime_ns) if candidates else None


def export_sales(
    *,
    business_date: date,
    source_url: str = DEFAULT_SOURCE_URL,
    download_dir: str | Path,
    pause_for_login: bool = False,
    page_wait_seconds: int = 8,
    export_wait_seconds: int = 8,
    timeout_seconds: int = 180,
) -> Path:
    try:
        import pyautogui
    except ImportError as exc:
        raise RuntimeError("缺少 pyautogui，请先执行 pip install -r requirements.txt") from exc

    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.35
    folder = Path(download_dir).expanduser().resolve()
    before = download_snapshot(folder)
    started_at = time.time()

    webbrowser.open(source_url, new=2)
    time.sleep(page_wait_seconds)
    if os.name == "nt":
        pyautogui.hotkey("win", "up")
        time.sleep(1)

    width, height = pyautogui.size()
    points = {name: scaled_point(x, y, width, height) for name, (x, y) in POINTS.items()}
    pyautogui.click(*points["export_data"])
    time.sleep(export_wait_seconds)
    pyautogui.click(*points["export_history"])
    time.sleep(2)

    deadline = time.monotonic() + timeout_seconds
    last_click = 0.0
    candidate: Path | None = None
    stable_size = -1
    stable_checks = 0
    while time.monotonic() < deadline:
        if time.monotonic() - last_click >= 5:
            pyautogui.click(*points["download_latest"])
            last_click = time.monotonic()
        current = find_new_xlsx(folder, before, started_at)
        if current:
            size = current.stat().st_size
            if current == candidate and size == stable_size and size > 0:
                stable_checks += 1
            else:
                candidate, stable_size, stable_checks = current, size, 0
            if stable_checks >= 2:
                return current
        time.sleep(1)
    raise TimeoutError(f"{timeout_seconds} 秒内未在 {folder} 检测到新的 XLSX 文件")


def main() -> int:
    parser = argparse.ArgumentParser(description="用 PyAutoGUI 从良域零售额台账导出 XLSX")
    parser.add_argument("--date", default=(date.today() - timedelta(days=1)).isoformat())
    parser.add_argument("--url", default=DEFAULT_SOURCE_URL)
    parser.add_argument("--download-dir", default=str(Path.home() / "Downloads"))
    parser.add_argument("--pause-for-login", action="store_true", help="兼容旧命令；现在自动继续")
    parser.add_argument("--timeout", type=int, default=180)
    args = parser.parse_args()
    output = export_sales(
        business_date=date.fromisoformat(args.date),
        source_url=args.url,
        download_dir=args.download_dir,
        pause_for_login=args.pause_for_login,
        timeout_seconds=args.timeout,
    )
    print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
