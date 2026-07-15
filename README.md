# Shipment Tracking Automation

一套為物流 Planner 團隊打造的貨物追蹤自動化工具。每日自動整合多個資料來源（船公司網站/API、報關行報表、倉庫進倉報表等），更新統一的 PO 追蹤總表，讓 Planner 能快速掌握各批貨物的即時狀態、預估進倉時間，並計算進倉優先順序。

> 本專案原為公司內部物流追蹤系統，已取得公司授權作為作品公開展示。所有帳號密碼、API 金鑰、內部網路路徑、真實人名等機敏資訊皆已移除或替換為環境變數/通用範例，不影響程式邏輯的完整性。

## 專案目的

- 提升 Planner 對貨況的可視性，減少手動彙整報表的時間
- 依據庫存水位與貨況自動判斷各批貨的進倉緊急程度

## 架構總覽

程式由一個主控腳本（`Main.py`）依序呼叫七大功能模組，另有五個船公司追蹤子模組與兩個共用工具模組：

| # | 模組 | 說明 |
|---|------|------|
| 1 | `sales_data_sync.py` | 同步最新的 PO 銷售資料到主表 |
| 2 | `cargoo.py` | 以瀏覽器自動化（Playwright）登入 Cargoo 平台，下載並匯入 Shipment 報表 |
| 3 | `carrier_tracking_main.py` | 呼叫五家船公司模組取得即時船運狀態，作為 Cargoo 資料的輔助/交叉驗證 |
| 4 | `customs.py` | 整合多家報關行報表，取得報關進度與 ETA |
| 5 | `warehouse.py` | 整合倉庫進櫃計畫表，推算隔日進倉排程 |
| 6 | `compare_last_file.py` | 比對前後兩版報表，記錄 ETA 變動歷程 |
| 7 | `sort_query_sap.py` | 更新 Excel Power Query 與 SAP 連結資料，供進倉優先順序計算使用 |

### 船公司追蹤子模組 (`carrier_tracking/`)

透過官方 API（如 Maersk）、非官方 API 或網頁爬蟲（Evergreen、萬海、陽明、ONE）取得各家船公司的即時船運狀態，涵蓋約六到七成的訂單。

### 共用工具 (`support_module/`)

- `support_function.py`：檔案讀寫、Excel 工作表複製等共用函式
- `email_automation.py`：透過 Outlook COM 介面自動抓取信箱附件（如進櫃通知單）

## 技術棧

- Python 3.12
- Playwright（瀏覽器自動化）
- pandas / openpyxl / xlwings（Excel 資料處理）
- requests（船公司 API 串接）
- pywin32（Outlook 自動化，僅限 Windows）

## 安裝與設定

```bash
pip install -r requirements.txt
playwright install chromium
```

程式需要以下環境變數（原始程式碼中的實際帳密與金鑰皆已移除，需自行申請/設定）：

```bash
CARGOO_EMAIL=your_email@example.com
CARGOO_PASSWORD=your_password
MAERSK_CLIENT_ID=xxx
MAERSK_CLIENT_SECRET=xxx
MAERSK_CLIENT_ID_2=xxx        # 若有第二組 API 權限
MAERSK_CLIENT_SECRET_2=xxx
```

各模組開頭的路徑常數（如 `YOUR_NETWORK_DRIVE\ShipmentTracking\...`）為原內部網路磁碟機路徑的通用替代，實際使用時請改為自己的資料夾路徑。

## 使用方式

1. 開啟 `Main.py`，確認來源資料夾與暫存資料夾路徑
2. 依需求選擇「創建新檔」或「使用現有檔」（將不需要的段落註解掉）
3. 選擇要執行的模組（七大模組皆可依需求註解跳過）
4. 執行 `Main.py`，程式會依序更新 Excel 主表各區塊

完整流程約需 16 分鐘，其中船運追蹤模組（`carrier_tracking_main`）佔比最高，可視需求選擇是否執行。

## 免責聲明

本專案為技術作品展示，程式碼中提及的第三方服務（Cargoo、Maersk、報關行等）名稱僅為說明資料整合對象，不代表任何商業合作關係。所有原始機敏資訊已於公開前完整移除。
