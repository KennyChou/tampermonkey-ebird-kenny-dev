# ebird-kenny

eBird 的 Tampermonkey 使用者腳本，提供清單批次下載、行程下載、群組名片夾與 GPS 軌跡匯出功能。

## 功能

| 功能 | 觸發頁面 |
|------|----------|
| 批次下載清單 (Excel) | `ebird.org/mychecklists` |
| 下載行程報告 (Excel) | `ebird.org/mytripreports` |
| 群組名片夾（快速分享） | `ebird.org/checklist/*` |
| 下載 GPS 軌跡 (KML) | `ebird.org/checklist/*` |

### 批次下載清單

在「我的清單」頁面，每筆清單前會出現勾選框，勾選後按「下載選取的清單」即可將觀察資料匯出為 `.xlsx`，欄位包含：物種俗名／學名、數量、地點、日期、觀察方式、繁殖代碼等。

### 下載行程報告

在「我的行程報告」頁面，每筆行程後會出現「下載」按鈕，點擊後自動抓取該行程所有清單並匯出為 `.xlsx`。

### 群組名片夾

在清單分享頁面（`/checklist/*`），側邊欄會多出「群組通訊錄」區塊：

- 可建立多個群組，每組存放一批 eBird 使用者 ID
- 點擊群組名稱即可一鍵填入分享欄位
- 資料儲存於瀏覽器 `localStorage`，不需額外帳號

### 下載 GPS 軌跡

在清單頁面，若該清單含 GPS 軌跡，「軌跡圖」旁會出現「下載軌跡」按鈕，下載格式為 `.kml`（可匯入 Google Earth、Garmin BaseCamp 等軟體）。

## 安裝

1. 安裝瀏覽器擴充功能 [Tampermonkey](https://www.tampermonkey.net/)
2. 建立新腳本，將 `ebird-kenny.user.js` 全文貼入並儲存

## 初始設定

首次使用需完成兩個步驟（在「我的清單」或「我的行程報告」頁面操作即可）：

1. **輸入 eBird API Key**
   - 至 [ebird.org/api/keygen](https://ebird.org/api/keygen) 申請
   - 貼入欄位後按「儲存」

2. **下載鳥名錄**
   - 選擇顯示語言（中文俗名 / 英文俗名）
   - 按「下載鳥名錄」（約數秒，資料會快取至 `localStorage`）

完成後即可正常使用所有清單／行程下載功能。如需切換語言，清除 `localStorage` 中的 `sp_info` 後重新下載即可。

## 資料存放位置

| 鍵值 | 內容 |
|------|------|
| `localStorage.ebirdKey` | eBird API Key |
| `localStorage.sp_info` | 鳥種名錄快取（JSON） |
| `localStorage.ebird-groups` | 群組名片夾資料（JSON） |

## 依賴

- [SheetJS (xlsx) 0.18.5](https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js)（透過 `@require` 自動載入）
- eBird API v2（需 API Key）

## 作者

Kenny Chou
