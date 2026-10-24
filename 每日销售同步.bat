@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul
title 万象汇每日销售同步

echo ============================================================
echo   万象汇每日销售同步
echo   自动导出本月 1 日至昨天的数据并覆盖更新润工作云表
echo ============================================================
echo.
echo 注意：本次会更新云表中对应日期的销售金额和合计。
set /p "CONFIRM=确认继续？输入 Y 后按 Enter："
if /i not "%CONFIRM%"=="Y" (
  echo 已取消，没有运行同步。
  pause
  exit /b 0
)
echo.

if not exist "config.json" (
  echo [需要配置] 找不到 config.json。
  echo 请先复制 config.example.json 并命名为 config.json，
  echo 再填写 mmsApi.caName 和 mmsApi.caPin。
  echo 详细说明请查看 WINDOWS_DAILY_TASK.md。
  echo.
  pause
  exit /b 2
)

if not exist "logs" mkdir "logs" >nul 2>nul
where python >nul 2>nul
if errorlevel 1 goto try_py

python -u run_daily_sales.py --month-to-yesterday --config config.json
set "RUN_EXIT=%ERRORLEVEL%"
goto show_result

:try_py
where py >nul 2>nul
if errorlevel 1 (
  echo [环境问题] 找不到 Python。
  echo 请安装 Python 3 并勾选“Add Python to PATH”，然后运行：
  echo python -m pip install -r requirements-api.txt
  echo.
  pause
  exit /b 9009
)
py -3 -u run_daily_sales.py --month-to-yesterday --config config.json
set "RUN_EXIT=%ERRORLEVEL%"

:show_result
echo.
if "%RUN_EXIT%"=="0" (
  echo [完成] 月度数据已刷新。可在上方结果中查看日期范围、行数和回读校验。
) else (
  echo [未完成] 同步失败，错误建议和详细日志路径已显示在上方。
  echo 也可打开项目 logs 文件夹，查看最新的 run_daily_sales_*.log。
  echo 将日志末尾的 ERROR/Traceback 内容提供给维护人员即可定位。
)
echo.
pause
exit /b %RUN_EXIT%
