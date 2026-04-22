# 派工系統

以 Google Sheets 為資料庫、Google Apps Script 為後端 API、Cloudflare Pages 為前端的內部派工回報網站。  
設計以手機使用為主。

---

## 功能概覽

### 員工端
- 工號 + 密碼登入（含組別）
- 填寫派工記錄：選**船號** → 勾選當日**工單號碼**（可多選）→ 填施工日期
- 支援臨時工單：手動輸入船號與工單號碼
- 查看個人歷史記錄（依年/月篩選）

### 主管端
| 分頁 | 功能 |
|------|------|
| 工單管理 | 手動新增工單、從 Google Sheet 批次匯入（分頁名稱=船號，內容=工單號碼）、刪除工單 |
| 派工記錄 | 查詢所有員工的派工記錄，可依日期範圍、組別、工號、船號、工單號碼篩選，並選擇排序方式 |
| 統計分析 | 依工單、依船號、依組別、依員工的派工筆數統計，以及每日派工筆數時間軸（橫向可滑動） |

---

## 資料架構（Google Sheet 分頁）

| 分頁 | 欄位 | 說明 |
|------|------|------|
| `Users` | 工號、姓名、密碼、角色、組別 | 員工與主管帳號 |
| `Sessions` | token、userId、expiry | 登入 Session |
| `WorkOrders` | 工單號碼、船號、建立時間、備註 | 工單主檔 |
| `Reports` | 提交時間、施工日期、工號、員工姓名、組別、工單號碼、船號 | 所有派工記錄（集中儲存） |

> 工單號碼欄可儲存多個工單（逗號分隔），對應員工多選的情境。

---

## 部署流程

### 第一步：建立 Master Google Sheet

1. 前往 [Google Sheets](https://sheets.google.com) 新增一個空白試算表
2. 記下試算表的 **ID**（網址列中 `/d/` 和 `/edit` 之間的字串）：
   ```
   https://docs.google.com/spreadsheets/d/【這裡就是 ID】/edit
   ```
3. 分頁結構稍後由 `setupSheets()` 自動建立，**不需手動新增分頁**

---

### 第二步：部署 Google Apps Script

1. 在試算表中，點選上方選單 **擴充功能 → Apps Script**
2. 刪除編輯器中預設內容，將 `gas/code.gs` 全部內容貼上
3. 點選左側齒輪（**專案設定**） → **指令碼屬性** → 新增以下兩筆：

   | 屬性名稱 | 值 |
   |---------|---|
   | `MASTER_SHEET_ID` | 第一步記下的試算表 ID |
   | `MANAGER_INVITE_CODE` | 自訂邀請碼（主管註冊時使用，例如 `company2026`） |

4. **初始化分頁**：在編輯器上方函式選單選擇 `setupSheets`，按執行，授權後系統會自動建立 `Users`、`Sessions`、`WorkOrders`、`Reports` 四個分頁並設好標題列
5. 回到編輯器，點選「**部署 → 新增部署**」
6. 類型選「**網頁應用程式**」，設定如下：
   - 執行身分：**我**
   - 存取權限：**所有人**
7. 點選「**部署**」，授權後複製產生的 **Web App URL**

> ⚠️ 日後若修改 `code.gs`，必須重新部署（**部署 → 管理部署 → 編輯 → 版本選「新版本」**），URL 不會改變。

---

### 第三步：設定前端

開啟 `js/api.js`，將 `YOUR_GOOGLE_APPS_SCRIPT_URL` 替換為上一步取得的 Web App URL：

```js
const API_URL = 'https://script.google.com/macros/s/你的ID/exec';
```

---

### 第四步：建立 GitHub Private Repository

```bash
git init
git add .
git commit -m "init"
git branch -M main
git remote add origin https://github.com/你的帳號/你的repo名稱.git
git push -u origin main
```

---

### 第五步：部署到 Cloudflare Pages

1. 前往 [Cloudflare Dashboard](https://dash.cloudflare.com)，登入帳號
2. 左側選單 **Workers & Pages** → **Create** → **Pages**
3. 點選「**Connect to Git**」，選擇剛才建立的 private repo
4. Build settings 全部**留空**
5. 點選「**Save and Deploy**」，完成後取得 `https://xxx.pages.dev` 網址

> 之後每次 `git push` 到 `main`，Cloudflare 會自動重新部署。

---

## 首次使用

### 1. 建立主管帳號

1. 開啟網站，點選首頁「**主管帳號申請**」
2. 填入工號、姓名、密碼，以及 `MANAGER_INVITE_CODE`
3. 登入後進入管理者儀表板

### 2. 匯入船號與工單

準備一份 Google Sheet，格式如下：

- **每個分頁名稱 = 船號**（例如：`海勝一號`、`興盈二號`）
- **A 欄 = 工單號碼**（第 1 列為標題，第 2 列起為資料）

```
分頁「海勝一號」        分頁「興盈二號」
A1: 工單號碼           A1: 工單號碼
A2: WO-001            A2: WO-010
A3: WO-002            A3: WO-011
A4: WO-003            A4: WO-012
```

將此 Google Sheet **分享（可檢視）** 給部署 GAS 的 Google 帳號，然後：

1. 登入主管後台 → **工單管理** 分頁
2. 貼上 Google Sheet 網址，點選「**從 Google Sheet 匯入**」
3. 匯入成功後工單清單即會更新

> 若只想更新特定船號，在「指定分頁」欄填入該船的分頁名稱。

### 3. 員工自助註冊

1. 開啟網站，點選首頁「**員工註冊**」
2. 填入工號、姓名、**組別**、密碼
3. 登入後即可開始填寫派工記錄

---

## 日常使用

### 員工操作流程

1. 登入 → 「**填寫報工**」
2. 選擇**施工日期**
3. 從下拉選單選擇**船號** → 下方自動顯示該船的工單勾選清單
4. 勾選當天施工的**工單號碼**（可多選）
5. 若為臨時派工，選擇「手動輸入」填入船號與工單
6. 按「**送出報工**」

切換至「**我的記錄**」可依年/月查看歷史派工記錄。

### 主管操作

| 功能 | 路徑 |
|------|------|
| 新增單一工單 | 工單管理 → 填入工單號碼 + 船號 → 新增 |
| 批次匯入工單 | 工單管理 → 貼上 Google Sheet 網址 → 匯入 |
| 刪除工單 | 工單管理 → 工單清單 → 刪除 |
| 查詢派工記錄 | 派工記錄 → 設定篩選條件 → 查詢 |
| 統計分析 | 統計分析 → 設定日期範圍/組別 → 統計 |

---

## 注意事項

- GAS Web App URL 已寫入 `js/api.js`，請確保 repository 設為 **Private**
- 用於匯入的 Google Sheet 必須分享給 GAS 腳本的執行帳號（可檢視權限即可）
- Session token 有效期為 **1 年**，登出後失效
- 修改 `code.gs` 後必須在 Apps Script **重新部署新版本**才會生效
- `setupSheets()` 只會建立不存在的分頁，不會覆蓋已有資料，可安全重複執行
