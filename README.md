# LINE 群組一鍵報名系統 — 設定說明

## 這套系統做什麼


- **發布者**：在網頁上填好活動內容 → 按「建立並發送到群組」→ LINE 跳出選擇畫面 → 勾選群組 → 圖卡就以你自己的帳號發出去。
- **成員**：在群組按圖卡上的「我要報名」→ 自動帶入 LINE 身分 → 寫進 Google 試算表，完成報名。
  - 第一次會請他填「姓名＋手機」比對名單並綁定 LINE 帳號，之後每次都是一按就報名。

## 檔案

| 檔案 | 放哪裡 |
|---|---|
| `Code.gs` | Google 試算表的 Apps Script |
| `index.html` | 免費網頁空間（建議 GitHub Pages） |

試算表會有四個工作表：`名單`、`活動`、`報名紀錄`、`發布者`。

---

## 步驟一：Google 試算表 + Apps Script

1. 新建一份 Google 試算表。
2. 上方選單「擴充功能 → Apps Script」，把 `Code.gs` 全部內容貼上、存檔。
3. 上方函式選單選 `setup` → 按「執行」→ 依指示授權。回到試算表會看到四個工作表。
4. 在 `名單` 填入成員的 **姓名** 與 **手機**（後兩欄留空，系統會自動填）。
5. 右上「部署 → 新增部署作業」→ 類型選「網頁應用程式」：
   - 執行身分：**我**
   - 存取權：**所有人**
6. 部署後複製「網頁應用程式網址」（`https://script.google.com/macros/s/.../exec`），等下要用。

## 步驟二：放上 index.html（GitHub Pages）

1. 註冊 / 登入 GitHub，新建一個公開 repository（例如 `line-signup`）。
2. 上傳 `index.html`。
3. Settings → Pages → Branch 選 `main`、資料夾 `/ (root)` → Save。
4. 約一分鐘後得到網址：`https://你的帳號.github.io/line-signup/`

## 步驟三：LINE Developers 設定

1. 到 <https://developers.line.biz/console/> 用 LINE 帳號登入。
2. 建立一個 **Provider**（名稱隨意，例如你們團體名稱）。
3. 在 Provider 下建立 **LINE Login channel**：
   - Region：Taiwan；App types：勾 **Web app**
4. 建好後在 **Basic settings** 記下 **Channel ID**。
5. 切到 **LIFF** 分頁 → Add：
   - Size：**Full**
   - Endpoint URL：步驟二的 GitHub Pages 網址
   - Scopes：勾 **openid**、**profile**
   - **Share target picker：ON**（一定要開，否則無法發到群組）
   - 其他保持預設 → Add，記下 **LIFF ID**（像 `2001234567-AbCdEfGh`）
6. ⚠️ 把 channel 狀態從 **Developing** 改成 **Published**（頁面上方的按鈕）。沒改的話只有你自己能登入，成員會報名失敗。

## 步驟四：把三個值填回去

1. `Code.gs` 第 7 行：填入 **Channel ID** → 存檔 →「部署 → 管理部署作業 → 編輯（鉛筆）→ 版本選『新版本』→ 部署」（網址不會變）。
2. `index.html` 最上方兩行：
   ```js
   const LIFF_ID = '你的 LIFF ID';
   const GAS_URL = '步驟一的網頁應用程式網址';
   ```
   在 GitHub 上編輯存檔，等一分鐘生效。

## 步驟五：開通發布權限

1. 在手機 LINE 傳 `https://liff.line.me/你的LIFF ID` 給自己（例如 Keep 筆記），點開。
2. 第一次會顯示「你還沒有發布權限」和一串代碼 → 複製。
3. 貼到試算表 `發布者` 工作表第一欄（第二欄寫名字方便辨認）。
4. 重新開啟連結，就會看到發布平台。要讓其他人也能發布，請他們照做並把代碼給你加進去。

> 電腦也能用：直接用瀏覽器開 `https://liff.line.me/你的LIFF ID`，登入 LINE 後一樣可以建立活動並選擇群組發送。

---

## 日常使用

- **發布**：開發布平台 → 填內容 → 「建立並發送到群組」→ 選群組（可多選）→ 傳送。
- **再發一次 / 截止報名**：平台下方「已發布的活動」列表。
- **看報名名單**：試算表 `報名紀錄`，可用篩選看單一活動。

## 常見問題

- **圖片怎麼放？** 圖片網址必須是 `https://` 開頭、可公開存取的圖片連結（例如上傳到 imgur 後的圖片網址）。Google 雲端硬碟的分享連結通常無法顯示。
- **成員說「名單中找不到」**：確認 `名單` 的姓名與他輸入的完全一致（含空格），手機號碼正確。
- **有人換手機 / 換 LINE 帳號**：到 `名單` 把他那列的 `LINE userId` 清空，讓他重新綁定。
- **為什麼要選群組，不能直接指定？** 用個人帳號發送，LINE 只允許透過它的選擇畫面發送，這是 LINE 的限制。要完全不經選擇畫面自動發到固定群組，需要改用 LINE 官方帳號（機器人）加入群組。
- **改了 Code.gs 沒生效？** 每次修改都要「管理部署作業 → 新版本」重新部署。

## 個資提醒

系統會記錄成員的 LINE 名稱、LINE 內部識別碼、姓名與手機。建議在群組公告說明用途，並只開放必要的人存取試算表。

