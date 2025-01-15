# Feedback Sheet Generator

A Google Apps Script project that helps process and analyze participant feedback for educational events.

## Features

1. **Create Individual Feedback Sheets**
   - Generates separate spreadsheets for each speaker
   - Organizes participant feedback in a structured format
   - Maintains participant anonymity for certain feedback types

2. **NPS (Net Promoter Score) Processing**
   - Calculates NPS for each speaker
   - Generates personalized feedback messages
   - Updates processing status automatically

3. **Duplicate Feedback Detection**
   - Identifies duplicate feedback entries
   - Reports duplicate locations by row numbers
   - Processes all speakers' feedback sheets automatically

## How to Use

1. Run `executeCreateFeedbackSheets()` to generate individual feedback sheets
2. Run `executeNPSProcess()` to process NPS data
3. Run `executeCheckDuplicateFeedback()` to check for duplicate feedback

---

# 回饋表產生器

這是一個用於處理和分析教育活動參與者回饋的 Google Apps Script 專案。

## 功能

1. **建立個別回饋表**
   - 為每位講者產生獨立的試算表
   - 以結構化格式整理參與者回饋
   - 對特定類型的回饋維持參與者匿名性

2. **NPS（淨推薦值）處理**
   - 計算每位講者的 NPS
   - 產生個人化的回饋訊息
   - 自動更新處理狀態

3. **重複回饋檢查**
   - 識別重複的回饋內容
   - 以列數標示重複位置
   - 自動處理所有講者的回饋表

## 使用方式

1. 執行 `executeCreateFeedbackSheets()` 以產生個別回饋表
2. 執行 `executeNPSProcess()` 以處理 NPS 資料
3. 執行 `executeCheckDuplicateFeedback()` 以檢查重複回饋

