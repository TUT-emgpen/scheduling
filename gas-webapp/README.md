# GAS Web App

這個資料夾是 `Google Apps Script Web App` 版本。

目前有兩種用途：

1. 當成 Apps Script 內嵌版前端 + 後端
2. 當成 GitHub Pages 前端的「安全同步橋接後端」

## 檔案
- `Code.gs`: Apps Script 後端，負責讀寫 Google 試算表、自動備份、GitHub 安全同步橋接
- `Index.html`: Web App 前端頁面
- `appsscript.json`: Apps Script manifest

## 工作表規則
每個科系會自動建立：

- `科系名稱_學生名單`
- `科系名稱_自動排程`
- `科系名稱_設定`

系統也會建立隱藏工作表：

- `_系統備份`

## 佈署
詳見 [GOOGLE_SYNC_SETUP.md](/Users/james/Documents/4p/GOOGLE_SYNC_SETUP.md)
