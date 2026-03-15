# vercel-fill · Excel 回填服務（來源 → 目標）

使用 Next.js + Serverless Function，在 **後端以 Excel 序列日 + `yyyy/mm/dd` 格式**寫回目標，**避免瀏覽器時區造成日期退一天**。

## 功能
- 來源：掃描所有分頁逐列分析，車號欄（候選：`車號/最小成本單位`、`車號`、`最小成本單位`），時間欄（候選：`檢修日期`、`檢查時間`、`複查時間`、`時間`、`紀錄時間`）。**同列多時間欄取最大**，再對**每車號**取**最新**日期。
- 目標：預設分頁 `鋼輪計算_115`，按「車號」回寫「檢修日期」。以 **Excel 序列日**（自動偵測 1900/1904）+ 格式 `yyyy/mm/dd` 寫入，**僅 date-only**，不帶時分秒，**永不使用 JS Date 寫回**。
- 參數：`sheet`（分頁名）、`overwrite`（是否覆蓋公式）。

## 為什麼能解決日期錯誤
- 第 1 步匯出的來源檔**不需要**在前端轉 JS Date；本服務在 **後端**解析字串/序列，並寫入 **序列日 + 格式**，Excel 以內建日期系統解碼，**不受時區影響**。

## 快速開始

### 本機開發
```bash
npm i
npm run dev
# 瀏覽 http://localhost:3000
```

### 部署到 Vercel（推薦）
1. 在 Vercel Dashboard 選 **Add New → Project**。
2. 連動此 repo，框架選 Next.js，按 Deploy。
3. 完成後開啟專案網址，上傳 `source.xlsx` 與 `target.xlsx` 測試。

> 若檔案很大或處理時間偏長，可在 `vercel.json` 調整 `maxDuration`，或改用 Vercel Blob / S3 收檔後由 URL 抓取（可再擴充）。

## API
`POST /api/fill`，`multipart/form-data`
- `source`: 來源 Excel（第 1 步匯出）
- `target`: 目標 Excel（原表）
- `sheet`: 目標分頁名稱（預設 `鋼輪計算_115`）
- `overwrite`: `true|false`（`true`＝覆蓋公式；`false`＝保留公式、略過該列）

**回應**：`filled.xlsx` 檔案串流。

### 範例（curl）
```bash
curl -X POST   -F "source=@source.xlsx"   -F "target=@target.xlsx"   -F "sheet=鋼輪計算_115"   -F "overwrite=true"   -o filled.xlsx   https://YOUR-VERCEL-APP.vercel.app/api/fill
```

## 寫入規則細節
- 來源日期解析：
  - 14 碼 `yyyyMMddHHmmss`
  - Excel 序列日（含 1900/1904）
  - `yyyy/mm/dd[ HH:MM[:SS]]`、`yyyy-mm-dd[ HH:MM[:SS]]`
  - 在地化字串（剔除 `上午/下午/AM/PM/年/月/日` 後解析）
- 目標回寫：「檢修日期」只寫 **date-only**；採 Excel **序列數 + `yyyy/mm/dd`**，避免時區退日；若 `overwrite=false` 且儲存格含公式則略過。

## 後續擴充（可選）
- 新增 `keywords` 與 533～548 分頁寫回 `檢查結果`（鍵：`車號 + DM1/M2/T + 檢查項目` 取最新）。
- 輸出 `debug_來源日期比對.csv` 用於稽核（列出 rawTimes 與 picked）。

---

> 版本：2026-03-15T17:08:36.056711Z
