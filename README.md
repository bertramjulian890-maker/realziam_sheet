# 良域销售日报 → 润工作电子表格

本项目独立完成以下流程：打开良域“零售额台账”、导出销售明细、只保留昨天的 `店铺号 / 店铺名称 / 销售日期 / 销售总金额`，最后按店铺号把金额写入润工作电子表格对应日期列。

每次运行都会在项目目录的 `logs` 文件夹生成独立日志，记录启动参数、网页导出、日期完整性检查、逐日云表覆盖、合计和完整异常堆栈。计划任务的“上次运行结果”显示 `0x1` 时，先查看最新的 `logs/run_daily_sales_*.log`。

## 初始化（Windows CMD）

```cmd
git clone https://github.com/bertramjulian890-maker/realziam_sheet.git
cd realziam_sheet
python -m pip install -r requirements.txt
copy config.example.json config.json
```

示例配置已填写接口地址、良域页面和下载目录。`spreadsheetToken` 是共用电子表格 token；在 `monthlySheetIds` 中填写各月工作表 ID，未使用月份可以留空。链接 `/sheets/` 后为 token，`?sheet=` 后为工作表 ID。

## 安全测试

先用已有导出文件测试筛选和匹配，不写云表：

```cmd
python run_daily_sales.py --use-file "C:\Users\Administrator\Downloads\实际导出文件.xlsx" --date 2026-09-14 --dry-run
```

确认输出中的日期、行数和筛选文件无误后，去掉 `--dry-run` 正式写入。

## 鼠标自动导出并写入

先确保浏览器已登录良域，然后执行：

```cmd
python run_daily_sales.py
```

补跑本月 1 日至昨天时，脚本只下载一次最新导出文件，再依次覆盖每个日期列：

```cmd
python run_daily_sales.py --month-to-yesterday
```

导出页面本身需已包含本月 1 日至昨天的完整数据。脚本在写入前先检查每一天都有导出记录；任意日期缺失会停止，不会开始覆盖。指定任意日期区间：

```cmd
python run_daily_sales.py --start-date 2026-09-01 --end-date 2026-09-21
```

刷新某个已经结束的自然月（例如今天是 10 月，需要补刷 9 月）时，使用 `--month YYYY-MM`。它只校验并覆盖指定月份，不会把今天所在月份的日期混进来：

```cmd
python run_daily_sales.py --month 2026-09
```

如果 9 月导出文件已经下载好了，可直接使用该文件先做不写入云端的检查：

```cmd
python run_daily_sales.py --month 2026-09 --use-file "C:\Users\Administrator\Downloads\9月导出.xlsx" --dry-run
```

检查输出日期范围为 `2026-09-01` 至 `2026-09-30`、每天都有导出记录后，去掉 `--dry-run` 正式覆盖。网页导出文件必须包含指定月份的完整数据；脚本不会自动更改网页上的日期筛选。

脚本打开页面、等待加载后自动继续，无需按 Enter。旧参数 `--pause-for-login` 保留兼容，但不再暂停。网页日期条件保持原样，昨天的筛选在下载后的 XLSX 中完成。日常直接执行：

```cmd
python run_daily_sales.py
```

脚本默认刷新昨天所在月份的 1 日至昨天（月初 1 日运行时刷新上个月整月）。指定日期时只处理当天：

```cmd
python run_daily_sales.py --date 2026-09-14
```

## 纯接口取数（Windows CMD）

已接入已确认的销售导出接口：证书登录 → 创建导出 → 轮询本次任务 → `tokenUrl` 获取临时下载地址 → 下载并验证销售 Excel。该模式不启动浏览器，不依赖桌面鼠标。

安装接口模式依赖即可，无需安装 Playwright 浏览器：

```cmd
python -m pip install -r requirements-api.txt
```

万象汇项目 `20071` 和“业务专员”职位 ID 已设为默认。认证信息一次性写入本地 `config.json` 的 `mmsApi`，后续不需要输入 `set`。现有配置增加如下对象即可；仅需填写实际账号和证书 PIN：

```json
"mmsApi": {
  "projectCode": "20071",
  "positionId": "cd1f2432c46b4dbeac7b13ef4e1aaf87",
  "caName": "你的认证账号",
  "caPin": "你的证书PIN",
  "certificateIndex": 0
}
```

云表共用 `spreadsheetToken`，按 `monthlySheetIds` 的 `YYYY-MM` 键选择工作表。每月 1 日刷新上月整月，使用上月工作表；每月 2 日起使用当月工作表。跨月区间逐日选择对应工作表。所需月份 ID 为空时，在良域导出前停止并提示。旧的 `monthlySpreadsheetTokens` 和顶层 `sheetId` 请替换为以下结构：

```json
"spreadsheetToken": "shtk94liHum47bU7iGxeFGhvSld",
"monthlySheetIds": {
  "2026-09": "n1af7R",
  "2026-10": "BRy3Zv",
  "2026-11": ""
}
```

`certificateIndex` 是首个设备内的证书序号，默认 0；与业务职位 ID 不同。`config.json` 已被 Git 忽略，账号和 PIN 不进入仓库。原来的环境变量仍可作为未填配置时的备用来源。若 `sourceUrl` 与 `mmsApi.projectCode` 不一致，刷新入口会停止。

只取数，不访问云表，默认下载本月 1 日至昨天：

```cmd
python mms_api_export.py
```

正式月度流程默认使用接口取数；若原配置已写入其他 `exportEngine`，请改为 `"exportEngine": "api"`：

```cmd
python run_daily_sales.py
```

指定日期只跑当天（只取数和完整流程分别为）：

```cmd
python mms_api_export.py --date 2026-10-15
python run_daily_sales.py --date 2026-10-15
```

仍支持 `--month YYYY-MM` 和 `--start-date/--end-date` 补跑。下载其他项目需显式提供对应项目及职位；万象汇的默认职位不会自动用于其他项目。

先预览本月（访问云表进行匹配，不写入）：

```cmd
python run_daily_sales.py --dry-run
```

确认记录和目标云表后，去掉 `--dry-run` 执行正式覆盖。当前网页控制台草稿的取数入口尚未接入此模式。

程序默认每次经证书重新认证；联调时也支持临时设置 `MMS_ACCESS_TOKEN`，此时使用它并跳过证书认证，不应在代码中固定保存。每次动态请求下载凭证，不复用历史签名链接。遇到登录过期、业务失败、任务归属不明确或非有效销售 Excel 时停止；日志位于 `logs`。本地证书服务在注销或不同计划任务账户下是否可用，需要台式机验证。

## Playwright 取数试用（备用）

这是备用的浏览器取数模式，尚需在良域内网实测页面定位和登录续期；主流程默认使用纯接口取数。

Windows CMD 安装：

```cmd
python -m pip install -r requirements-playwright.txt
python -m playwright install chromium
python mms_playwright.py --login
```

首次会打开独立浏览器，请登录、进入零售额台账，设置日期范围后回到命令窗口按 Enter 保存。登录保存在项目 `.browser-profile`（已排除 Git）；不是日常 Edge/Chrome 的登录状态，不要同时运行多个使用此配置的任务。

先只测试取数，不写入云表：

```cmd
python mms_playwright.py --headed
python mms_playwright.py
```

前者显示浏览器便于检查，后者无界面运行。成功后可用后一条测试 Windows 锁屏计划任务。任务应使用首次登录时相同的 Windows 账户。登录过期时需重新执行 `--login`；无界面运行仍依赖内网、登录认证和浏览器环境。

接入月度流程先预览：

```cmd
python run_daily_sales.py --export-engine playwright --month-to-yesterday --dry-run
```

确认目标云表配置和日期范围正确后去掉 `--dry-run`。也可在配置中添加 `"exportEngine": "playwright"`，供命令行流程默认选择。当前网页控制台草稿的取数入口仍使用鼠标方式。

新模式不更改网页日期条件：独立浏览器可能不保留筛选日期，应先核对下载文件的实际日期范围；月度脚本仍在云表写入前检查所有目标日期。程序只选择本次新增导出记录，找不到或出现多个新记录会报错。错误日志和失败页面截图保存在 `logs`，截图可能包含业务内容，请按需分享。

## 鼠标坐标与计划任务

### 可选 CA 自动登录（待内网验证）

参考现有展销提醒脚本，新增通过本机证书客户端 `wss://127.0.0.1:10022/` 签名、统一认证、良域登录的流程。要求本机已安装对应证书客户端且有可用证书。凭据从当前进程环境变量读取，不写入代码或配置文件。

```cmd
set MMS_CA_NAME=你的认证账号
set MMS_CA_PIN=你的证书PIN
set MMS_CA_POSITION=0
python mms_playwright.py --ca-login --headed
```

`MMS_CA_POSITION` 是第一个证书设备内的证书序号，从 0 开始（默认 0）。以上仅下载文件，不写云表。成功后可去掉 `--headed` 测试无界面模式；月度入口支持 `--export-engine playwright --ca-login --month-to-yesterday --dry-run`。

这条路径尚未在公司内网验证。参考脚本没有说明登录返回令牌如何写入页面，因此当前先复用浏览器请求共享的 Cookie 并检查导出按钮是否出现；若接口认证成功但页面仍要求登录，会明确报错并留截图，需要继续核对前端登录响应处理。环境变量仅在当前 CMD 有效，计划任务需自行提供相同账户的运行环境。

坐标来自 1920×1080 截图，并按当前屏幕分辨率同比缩放：导出数据 `(1848,304)`、导出历史 `(1745,304)`、最新下载 `(1427,402)`。浏览器需最大化，页面缩放建议保持 100%。

PyAutoGUI 依赖可交互桌面。Windows 任务计划程序应选择“仅当用户登录时运行”；锁屏或无桌面会话下，鼠标点击通常无法可靠执行。

任务计划程序操作参数示例：程序填写 Python 的完整路径，参数填写 `run_daily_sales.py --month-to-yesterday`，起始于填写项目目录。示例：

```text
程序：C:\Users\Administrator\AppData\Local\Programs\Python\Python310\python.exe
参数：run_daily_sales.py --month-to-yesterday
起始于：D:\git\realziam_sheet
```

## 云表写入规则

- A 列必须是 `店铺号`，B 列必须是 `店铺名称`。
- 固定模板：第 1 行为字段名，第 2 行从 L 列起为日期，第 3 行开始为店铺。
- 日期支持“9月14日”、完整年月日及电子表格日期序列值；云表专用于当前业务年份，列匹配只比较月、日，忽略表头底层年份。导出数据仍按完整业务年月日筛选。
- 找到业务日期列后，整列覆盖 A 列店铺号对应的金额：最新导出中存在的店铺写入最新金额，未出现的店铺清空当天旧金额。日期列缺失或重复时停止，不新增列。
- C 至 K 列、星期表头保留；没有匹配到的店铺金额保持原值。
- 根据 A 列最后一个店铺号定位末行（跳过“合计/总计/小计/求和”标签），读取当天各店铺金额，由脚本计算合计数字并写入下一行。例如末家店铺在第 101 行，当天为 Z 列，则计算 Z3 至 Z101，数字写入 Z102，并回读核验。接口目前会把公式当文字存储，因此不发送公式。每次运行更新合计；手动修改店铺金额后需重新运行更新汇总。输出 totalBasis 表示求和范围、totalAmount 表示合计数字（dry-run 时为 null）。
- 写入后回读核对每个金额。结果中的 targetColumn 和 dateHeaderCell 显示目标列和日期表头位置，verified 表示回读校验通过；dry-run 时为 false。
- 本地店铺号在云表中缺失、重复，或导出表同日店铺号重复时立即停止，不做部分写入。
- 不删除云端任何行列。
