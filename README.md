# circerio

Software tooling, automation, and graphics/compatibility engineering.

我目前開放自由接案，主要處理 **Excel / VBA / Power Query / Python 自動化、資料整理與報表流程、Windows 小工具，以及既有程式的分析與除錯**。

## 接案方向 / Freelance Focus

- **Excel VBA / Macro 自動化**：批次匯入、欄位轉換、資料比對、報表產生、按鈕式操作流程
- **多來源 Excel 資料整合**：將不同客戶或不同格式的工作簿轉成統一資料結構
- **Power Query / Python 資料處理**：清洗、轉換、彙總、驗證與例外紀錄
- **Forecast / Actual / Master Data 比對**：依客戶、料號、月份等 Key 進行整合與差異分析
- **工作流程自動化**：把重複性的人工整理流程改成可重複執行的工具
- **Windows 小工具與技術問題排查**：既有工具整合、測試、Debug 與效能分析
- **AI-assisted development**：用 AI 加速原型、程式撰寫與測試，但以可驗證、可維護的交付結果為準

### Excel 自動化的典型架構

```text
不同來源 Excel
      ↓
Import / Mapping Layer
      ↓
Standardized Data
      ↓
Validation & Reconciliation
      ↓
Report / Dashboard / History
```

對於格式不一致的檔案，會優先採用 **Mapping / Adapter** 的方式處理，而不是把固定儲存格位置全部寫死，讓後續格式調整與新增客戶時比較容易維護。

## 可交付內容

依案件需求可包含：

- 可直接操作的 `.xlsm` 工具
- Power Query / VBA / Python 程式碼
- 客戶 / 欄位 Mapping 設定
- 匯入檢核與錯誤紀錄
- 自動彙總與差異報表
- 簡易按鈕、下拉選單與使用介面
- 測試資料與操作說明
- 後續格式調整與功能擴充

## Public Projects

### [Excel Forecast Automation Demo](portfolio/excel-forecast-automation-demo)
Synthetic manufacturing-data portfolio demo showing **different customer forecast formats → standardized schema → baseline / actual comparison → dashboard / validation**. Includes a downloadable Excel workbook, VBA adapter/reference source, mapping examples, and sample input files.

### [OBS Canvas Rescaler](https://github.com/circerio/obs-canvas-rescaler)
Browser-based local tool for comparing effective OBS Virtual Camera resolutions and browser-side rescaling behavior. Includes a complete local workflow, presets, launch scripts, and troubleshooting documentation.

### [OptiScaler engineering branch](https://github.com/circerio/OptiScaler)
Public engineering work around game graphics compatibility, frame-generation integration, DX11/DX12 paths, instrumentation, regression testing, and reproducible checkpoints.

### [RenoDX engineering branch](https://github.com/circerio/renodx)
Public engineering work related to HDR / rendering compatibility and integration testing alongside the graphics pipeline above.

## 工作方式

我偏好先用少量真實樣本確認資料結構與例外情況，再定義標準格式、驗收條件與可擴充方式。對企業內部工具，會特別注意：

- 不讓使用者需要懂程式才能操作
- 對缺欄位、日期格式、未知料號等異常提供清楚提示
- 保留來源與更新日期，方便追查資料
- 儘量避免難以維護的大量硬編碼
- 先建立可驗證的 MVP，再逐步增加報表、歷史版本與 UI

---

**Available for freelance automation / tooling projects.**

若是 Excel、VBA、資料整理、報表自動化或小型工具案件，可以直接透過接案平台提供需求與樣本檔，再依實際資料評估工期與範圍。
