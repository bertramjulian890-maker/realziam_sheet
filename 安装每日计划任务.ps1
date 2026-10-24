param(
    [string]$AtTime = "09:00",
    [string]$TaskName = "万象汇每日销售同步"
)

$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConfigPath = Join-Path $ProjectDir "config.json"
$Python = (Get-Command python.exe -ErrorAction SilentlyContinue).Source
if (-not $Python) {
    $Python = (Get-Command python -ErrorAction SilentlyContinue).Source
}

Write-Host ""
Write-Host "万象汇每日销售同步 - 计划任务设置" -ForegroundColor Cyan
Write-Host "项目目录：$ProjectDir"
Write-Host ""

if (-not $Python) {
    throw "未找到 Python。先安装 Python 3，并在安装时勾选 Add Python to PATH。"
}
if (-not (Test-Path $ConfigPath)) {
    throw "找不到 config.json。请先复制 config.example.json 为 config.json，并填写 mmsApi.caName、mmsApi.caPin。"
}
try {
    $Config = Get-Content -Raw -Encoding UTF8 $ConfigPath | ConvertFrom-Json
} catch {
    throw "config.json 不是有效 JSON：$($_.Exception.Message)"
}
if (-not $Config.mmsApi.caName -or -not $Config.mmsApi.caPin) {
    throw "config.json 的 mmsApi.caName 或 mmsApi.caPin 为空。请补齐后重新运行本文件。"
}

$TimeValue = [datetime]::MinValue
if (-not [datetime]::TryParseExact($AtTime, "HH:mm", [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::None, [ref]$TimeValue)) {
    throw "运行时间格式无效：$AtTime。请用 24 小时制 HH:mm，例如 09:00。"
}

Write-Host "检查 Python 依赖..."
& $Python -c "import openpyxl, requests, websocket" 2>$null
if ($LASTEXITCODE -ne 0) {
    throw "Python 依赖未安装。请在项目目录打开 CMD，运行：python -m pip install -r requirements-api.txt"
}

$Arguments = "-u `"$(Join-Path $ProjectDir 'run_daily_sales.py')`" --month-to-yesterday --config `"$ConfigPath`""
$Action = New-ScheduledTaskAction -Execute $Python -Argument $Arguments -WorkingDirectory $ProjectDir
$Trigger = New-ScheduledTaskTrigger -Daily -At $TimeValue
$Principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited
$Settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1) -RestartCount 2 -RestartInterval (New-TimeSpan -Minutes 5)
$Task = New-ScheduledTask -Action $Action -Trigger $Trigger -Principal $Principal -Settings $Settings `
    -Description "每天刷新万象汇本月 1 日至昨天销售数据；需要当前用户已登录且本机 CA 证书客户端可用。日志保存在项目 logs 文件夹。"

Register-ScheduledTask -TaskName $TaskName -InputObject $Task -Force | Out-Null
Write-Host ""
Write-Host "计划任务已创建：$TaskName" -ForegroundColor Green
Write-Host "每天运行时间：$($TimeValue.ToString('HH:mm'))"
Write-Host "运行方式：当前 Windows 用户登录后运行；锁屏状态可运行，注销状态不运行。"
Write-Host "失败详情：$ProjectDir\logs\run_daily_sales_*.log"
Write-Host ""
Write-Host "可在任务计划程序中找到该任务，右键“运行”进行验证。"
