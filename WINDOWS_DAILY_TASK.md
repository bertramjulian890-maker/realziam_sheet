# Windows 每日销售同步

每日流程会通过良域接口重新认证，导出万象汇当月 1 日至昨天的数据，按店铺覆盖润工作云表，并回读核对金额和合计。手动双击与计划任务使用同一个程序入口。

## 一次性准备

在项目目录打开 CMD，安装接口版依赖：

```cmd
python -m pip install -r requirements-api.txt
```

若还没有 `config.json`，复制 `config.example.json` 并重命名为 `config.json`。在 `mmsApi` 中填写 `caName` 和 `caPin`；在 `monthlySpreadsheetTokens` 中按 `YYYY-MM` 填入每个月对应的电子表格 token。未用月份可以留空，但本次运行涉及的月份 token 为空时会立即停止，并提示需要填写的月份；不会先调用良域导出，也不会写云表。检查工作表 ID 和下载目录。`config.json` 是本地文件，不会提交到 Git。

运行前请确认 Windows 用户已登录、本机 CA 证书客户端正在运行、证书可用，并能访问良域和润工作内网接口。

## 双击立即运行

双击项目目录中的 `每日销售同步.bat`。窗口会显示执行日期、同步进度和结果，结束后暂停，方便查看提示。成功时检查 `verified` 和各日行数；失败时先看窗口中的“处理建议”，再查看最新的 `logs/run_daily_sales_*.log`。

常见错误提示：

- 找不到 `config.json`：从示例配置复制，并填写账号、证书 PIN。
- 无法连接本机证书客户端：登录 Windows，启动 CA 客户端并检查证书。
- 证书签名失败：检查 PIN；多个证书时核对 `certificateIndex`。
- 导出任务超时或失败：检查内网和良域项目权限，之后可重新运行。
- 下载或日期校验失败：检查导出是否完整；脚本会在开始云端覆盖前校验所有目标日期。
- 云表写入/回读失败：检查润工作内网接口、表格 token、工作表 ID；结合日志定位具体阶段。

## 建立每日计划任务

在项目目录双击 PowerShell（或打开 PowerShell 后切换到项目目录），运行：

```powershell
powershell -ExecutionPolicy Bypass -File ".\安装每日计划任务.ps1"
```

默认每天 09:00 运行。可自定义时间：

```powershell
powershell -ExecutionPolicy Bypass -File ".\安装每日计划任务.ps1" -AtTime "08:30"
```

任务以当前 Windows 用户的交互登录身份运行，不要求管理员权限。选择此模式是为了沿用该用户会话中的本机 CA 证书客户端。电脑锁屏时通常可以运行；用户注销后不会运行。任务不会弹出命令窗口，成功/失败日志都写入项目 `logs` 文件夹。

在“任务计划程序”中找到“万象汇每日销售同步”，右键“运行”可立即验证；“上次运行结果”`0x0` 表示脚本成功，其他结果请查看最新日志。若需要改时间，再次运行安装脚本并传入新的 `-AtTime`。

## 日期规则

- 每日计划任务和双击入口：本月 1 日至昨天。
- 月初第一天运行：刷新上一个完整自然月，并使用上月对应的 `YYYY-MM` token；每月 2 日起使用当月 token。
- 手动刷新日期区间时，区间跨越月份会逐日选用各自月份的 token；缺少任一月份 token 会在导出前停止。
- 要刷新指定整月：

```cmd
python run_daily_sales.py --month 2026-09
```

- 仅运行某一天：

```cmd
python run_daily_sales.py --date 2026-10-22
```
