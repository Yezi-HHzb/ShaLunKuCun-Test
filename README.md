# ShaLunKuCun-Test

这是砂轮库存二维码自动化功能的独立测试仓库，不会修改正式仓库。

## 手动同步

1. 打开仓库的 **Actions** 页面。
2. 选择 **Manual Sync Feishu Data and QR Codes**。
3. 点击 **Run workflow**。
4. 任务完成后下载仓库根目录的 `E10库存二维码.xlsx`。

同步任务会读取飞书多维表格，更新 `data.json`，为新品号生成二维码，并从 `E10品号-规格-30B套圈.xlsx` 生成 Excel。仅处理“最新”工作表：F 列品号存在于飞书时，在 J 列放入二维码；不存在时填写“E10库存二维码”。

历史二维码保留在 `qrcodes/`，状态记录在 `qr_manifest.json`。
