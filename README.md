# ebird-kenny

> eBird 非官方 Tampermonkey 工具腳本 — 清單批次下載、行程匯出、群組名片夾、GPS 軌跡

[![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)](ebird-kenny.user.js)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](#授權)
[![Tampermonkey](https://img.shields.io/badge/Tampermonkey-✓-brightgreen.svg)](https://www.tampermonkey.net/)
[![eBird API v2](https://img.shields.io/badge/eBird%20API-v2-orange.svg)](https://documenter.getpostman.com/view/664302/S1ENwy59)
[![Tests](https://img.shields.io/badge/E2E%20Tests-23%20passed-success.svg)](tests/e2e)

[安裝](#安裝) · [初始設定](#初始設定) · [功能說明](#功能說明) · [匯出欄位](#匯出欄位) · [開發](#開發)

---

## 功能說明

### 批次下載清單 — `/mychecklists`

在「我的清單」頁面，每筆紀錄前注入勾選框。勾選後按下載按鈕，腳本逐筆呼叫 eBird API 取得詳細資料，匯整輸出為 `.xlsx`。

- 可一次選取任意數量的清單
- 欄位涵蓋物種、數量、地點座標、觀察方式、繁殖代碼等（[完整欄位列表](#匯出欄位)）

---

### 下載行程報告 — `/mytripreports`

在「我的行程報告」頁面，每筆行程旁注入「下載」按鈕。一鍵抓取行程下所有清單並合併輸出為 `.xlsx`。

---

### 群組名片夾 — `/checklist/*`

在清單分享頁面，側欄新增「群組通訊錄」區塊：

| 操作 | 效果 |
|------|------|
| 點擊「全部成員」| 一鍵填入所有已知聯絡人 |
| 點擊群組名稱 | 填入該群組的 eBird ID（不重複附加）|
| 新增 / 編輯 / 刪除群組 | 即時儲存至 `localStorage` |

---

### GPS 軌跡下載 — `/checklist/*`

在含有 GPS 軌跡的清單頁，軌跡圖旁出現「下載軌跡」按鈕。輸出標準 `.kml` 檔，可直接匯入：

- Google Earth / Google 我的地圖
- Garmin BaseCamp / Connect
- OsmAnd、CalTopo 等 GIS 工具

---

## 安裝

1. 安裝瀏覽器擴充功能 [Tampermonkey](https://www.tampermonkey.net/)（支援 Chrome、Firefox、Edge、Safari）
2. 點擊下方連結，Tampermonkey 會自動彈出安裝確認視窗

   **[點此一鍵安裝腳本](https://github.com/KennyChou/tampermonkey-ebird-kenny-dev/raw/refs/heads/main/ebird-kenny.user.js)**

> 也可手動建立新腳本，將 `ebird-kenny.user.js` 全文貼入並儲存。

---

## 初始設定

首次使用需完成以下兩步（在「我的清單」或「我的行程報告」頁面操作）：

### 步驟一：輸入 eBird API Key

1. 前往 [ebird.org/api/keygen](https://ebird.org/api/keygen) 申請金鑰
2. 將金鑰貼入頁面上方欄位，按「儲存」

### 步驟二：下載鳥名錄

1. 選擇名稱顯示語言（中文俗名 / 英文俗名）
2. 按「下載鳥名錄」（約數秒，快取於 `localStorage`）

> 設定完成後即可正常使用。若需切換語言，清除瀏覽器 `localStorage` 中的 `sp_info` 後重新下載即可。

---

## 匯出欄位

Excel 輸出欄位對應 eBird 官方匯出格式：

| 欄位 | 說明 |
|------|------|
| Submission ID | 清單 ID（`S` 開頭） |
| Common Name | 俗名（依選擇語言）|
| Scientific Name | 學名 |
| Taxonomic Order | 分類順序 |
| Count | 數量 |
| State/Province | 州／省代碼 |
| County | 縣市代碼 |
| Location ID | 地點 ID（`L` 開頭） |
| Location | 地點名稱 |
| Latitude / Longitude | 座標 |
| Date | 日期（`YYYY-MM-DD`）|
| Time | 時間（`HH:MM`）|
| Protocol | 調查方法 |
| Duration (Min) | 調查時間（分鐘）|
| All Obs Reported | 是否完整記錄（0/1）|
| Distance Traveled (km) | 移動距離 |
| Area Covered (ha) | 調查面積 |
| Number of Observers | 觀察人數 |
| Breeding Code | 繁殖行為代碼 |
| Observation Details | 物種備註 |
| Checklist Comments | 清單備註 |

---

## 資料儲存位置

所有資料僅存於本機瀏覽器，不傳送至任何第三方伺服器。

| `localStorage` 鍵值 | 內容 |
|----------------------|------|
| `ebirdKey` | eBird API Token |
| `sp_info` | 鳥種名錄快取（JSON）|
| `ebird-groups` | 群組名片夾資料（JSON）|

---

## 開發

### 環境需求

- Node.js 18+
- Playwright（E2E 測試）

### 執行測試

```bash
npm install
npx playwright install chromium
npm test
```

測試涵蓋所有四個模組，共 23 個 E2E 測試案例（使用 mock 頁面，不需要真實 eBird 帳號）：

| 測試檔 | 涵蓋模組 | 測試數 |
|--------|----------|--------|
| `checklist.spec.js` | `/mychecklists` 清單下載 | 6 |
| `trip.spec.js` | `/mytripreports` 行程下載 | 4 |
| `group.spec.js` | 群組名片夾 CRUD | 8 |
| `track.spec.js` | GPS 軌跡 KML 輸出 | 4 |
| | **合計** | **23** |

### 依賴套件

| 套件 | 用途 | 載入方式 |
|------|------|----------|
| [SheetJS (xlsx)](https://sheetjs.com/) `0.18.5` | Excel 輸出 | `@require` CDN |
| [eBird API v2](https://documenter.getpostman.com/view/664302/S1ENwy59) | 清單、鳥種、熱點資料 | GM_xmlhttpRequest |

---

## 授權

MIT © [Kenny Chou](https://github.com/KennyChou)
