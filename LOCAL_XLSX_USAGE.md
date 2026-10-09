# 本地 XLSX 整月刷新

本模式不连接润工作接口。它把良域导出的销售明细按整月覆盖到本地年度工作簿：

- 默认目标文件：`D:\git\realziam_sheet\2026年汕汇全场业绩手工表.xlsx`
- 匹配导出表字段：`店铺号`、`店铺名称`、`销售日期`、`销售总金额`
- 读取目标工作簿活动工作表：A1/B1 为店铺号/店铺名称，第 2 行 L 列起为日期，第 3 行起为门店。
- 日期表头匹配月/日，允许电子表格序列底层年份不同。
- 覆盖该月每天所有门店金额；最新导出缺少的门店会清空旧金额；日合计写到最后一个门店行的下一行。
- 导出缺少某天、存在重复店铺号或目标表没有对应店铺/日期时，脚本停止且不改目标文件。
- 正式写入前在目标文件旁生成带时间戳的备份；写入采用临时文件替换，避免半写入损坏原表。

## 台式机 CMD

先更新仓库并安装依赖（首次使用）：

```cmd
git pull
python -m pip install -r requirements.txt
```

先预览，不改目标工作簿：

```cmd
python local_month_refresh.py --month 2026-10 --export-file "C:\Users\Administrator\Downloads\10月销售导出.xlsx" --dry-run
```

核对月份、天数、门店数、清空数量和日合计无误后，正式刷新：

```cmd
python local_month_refresh.py --month 2026-10 --export-file "C:\Users\Administrator\Downloads\10月销售导出.xlsx"
```

若目标工作簿不在默认路径，增加 `--target "完整路径.xlsx"`。刷新后打开目标文件抽查几天金额和合计；若需回退，使用旁边自动生成的 `*.backup_时间戳.xlsx` 备份。
