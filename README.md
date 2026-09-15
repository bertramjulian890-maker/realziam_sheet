# 良域销售日报 → 润工作电子表格

本项目独立完成以下流程：打开良域“零售额台账”、导出销售明细、只保留昨天的 `店铺号 / 店铺名称 / 销售日期 / 销售总金额`，最后按店铺号把金额写入润工作电子表格对应日期列。

## 初始化（Windows CMD）

```cmd
git clone https://github.com/bertramjulian890-maker/realziam_sheet.git
cd realziam_sheet
python -m pip install -r requirements.txt
copy config.example.json config.json
```

默认配置已经填写接口地址、目标电子表格令牌、良域页面和下载目录。若目标电子表格只有一个工作表，`sheetId` 可以留空；多个工作表时必须填写实际 ID。

## 安全测试

先用已有导出文件测试筛选和匹配，不写云表：

```cmd
python run_daily_sales.py --use-file "C:\Users\Administrator\Downloads\实际导出文件.xlsx" --date 2026-09-14 --dry-run
```

确认输出中的日期、行数和筛选文件无误后，去掉 `--dry-run` 正式写入。

## 鼠标自动导出并写入

第一次运行或登录失效时：

```cmd
python run_daily_sales.py --pause-for-login
```

浏览器打开后完成登录、进入零售额台账，再回到 CMD 按 Enter。登录状态稳定后，日常可直接执行：

```cmd
python run_daily_sales.py
```

脚本默认处理昨天。也可补跑指定日期：

```cmd
python run_daily_sales.py --date 2026-09-14
```

## 鼠标坐标基准

坐标来自 1920×1080 截图，并按当前屏幕分辨率同比缩放：开始日期 `(680,250)`、结束日期 `(870,250)`、导出数据 `(1848,304)`、导出历史 `(1745,304)`、最新下载 `(1427,402)`。浏览器需最大化，页面缩放建议保持 100%。

PyAutoGUI 依赖可交互桌面。Windows 任务计划程序应选择“仅当用户登录时运行”；锁屏或无桌面会话下，鼠标点击通常无法可靠执行。

## 云表写入规则

- A 列必须是 `店铺号`，B 列必须是 `店铺名称`。
- 找到业务日期列后，按店铺号逐行更新金额。
- 日期列不存在时，在最后一个非空表头右侧创建日期列。
- 本地店铺号在云表中缺失、重复，或导出表同日店铺号重复时立即停止，不做部分写入。
- 不删除云端任何行列。
