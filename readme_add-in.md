# Ｍacaca Excel Add-in 使用指南

## 介紹

這份指南將說明如何使用我們的 Excel Add-in，並介紹如何執行 `selectRangeToJson` 功能。此 Add-in 是使用 Yeoman Office Add-in 工具建立的，並且需要使用 npm 來進行開發和運行。

## 安裝和設定

1. **安裝必要工具**： 首先，您需要安裝 Node.js 和 npm。請確保您已安裝最新版本的 Node.js (version 8.0.0 or later)。
2. **運行專案**： 開啟專案，運行以下命令在本地啟用add-in server：

   `npm start`
3. 開啟Excel檔案：開啟您要編輯的Excel檔案，如果您在上一個步驟成功啟用server，您會在右上角看到taskpane的圖示，點選taskpane即可開始使用本add-in。

## 使用 Add-in

### 主要功能：選擇範圍並轉換為 JSON

在 Add-in 中，我們提供了 `selectRangeToJson` 功能，可以將指定範圍的資料轉換為 JSON。以下是使用此功能的步驟：

1. **輸入範圍**： 在文本框中輸入您想要轉換的範圍（例如，`A1:C3`）。
2. **按下按鈕**： 點擊 `Get Json` 按鈕。
3. **查看結果**： 轉換結果將顯示在一個彈窗中。
