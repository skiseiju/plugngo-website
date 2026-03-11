# PlugnGO Beta 部署說明

## 檔案說明
- `Code.gs` → Google Apps Script（後端 API + Email 寄送）
- `beta/index.html` → 前端申請頁面，放進你的 GitHub repo

---

## 步驟一：建立 Google Sheets

1. 開新的 Google Sheets
2. 把網址列中的 ID 複製起來
   - 例如：`https://docs.google.com/spreadsheets/d/【這段就是ID】/edit`
3. 不需要手動建工作表，Apps Script 會自動建立

---

## 步驟二：部署 Apps Script

1. 開啟 [script.google.com](https://script.google.com)
2. 新增專案，把 `Code.gs` 的內容貼進去
3. 把第一行的 `YOUR_GOOGLE_SHEET_ID` 替換成你的 Sheets ID
4. 先跑一次 `setupRecordsSheet()`（建立申請紀錄工作表）
5. 再跑一次 `generateCodes()`（產生 300 組序號）
6. 部署 → 新增部署 → 類型選「網頁應用程式」
   - 執行身份：**我**
   - 存取權：**所有人**
7. 授權後，複製 Web App URL

---

## 步驟三：更新前端頁面

打開 `beta/index.html`，把這行：
```js
const API_URL = 'YOUR_APPS_SCRIPT_WEB_APP_URL';
```
替換成剛才複製的 Web App URL。

---

## 步驟四：Push 到 GitHub

把 `beta/` 資料夾放進你的 repo：
```
plugngo/
└── beta/
    └── index.html
```

Push 後 Cloudflare Pages 自動部署，
網址就會是：`plugngo.skiseiju.com/beta`

---

## 注意事項

- Apps Script 免費版每天可寄 **100 封** Email
  - 如果一天超過 100 人申請，當天後面的人會收不到信
  - 建議觀察流量，必要時升級 Google Workspace
- 序號發完後，頁面會自動顯示「名額已滿」訊息
- 所有申請紀錄都在 Sheets 的「申請紀錄」工作表

---

## Google Sheets 結構（自動建立）

**序號庫**
| 序號 | 狀態 |
|------|------|
| PNG-BETA-A3F7K2 | （空白=未使用）|
| PNG-BETA-B8M2NP | USED |

**申請紀錄**
| 時間戳記 | 姓名 | Email | 拍攝類型 | 序號 |
|----------|------|-------|----------|------|
