# 亞翔學習平台 | L&K Engineering Learning Platform

企業內部教學平台原型 — 包含前台（課程瀏覽、學習紀錄、測驗、AI 助手）與後台管理（課程、分類、使用者、學習分析、測驗紀錄）。

## 📋 系統需求

開始之前，請先安裝 **Node.js**：
- 前往 https://nodejs.org
- 下載並安裝 **LTS** 版本（推薦長期支援版）
- 安裝完成後，打開命令提示字元（Win+R 輸入 cmd）輸入 `node -v` 確認安裝成功

## 🚀 第一次使用

### 步驟 1：解壓縮專案
將 zip 解壓縮到任意位置（例如 `D:\lk-learning`）

### 步驟 2：開啟命令提示字元並切換到專案目錄

Windows 範例：
```
cd D:\lk-learning
```

### 步驟 3：安裝套件（只需要做一次）
```
npm install
```
這會下載專案需要的所有套件，第一次大約需要 1-3 分鐘。

### 步驟 4：啟動開發伺服器
```
npm run dev
```

執行成功會看到類似這樣的訊息：
```
  VITE v5.x.x  ready in xxx ms

  ➜  Local:   http://localhost:5173/
  ➜  Network: http://192.168.x.x:5173/
```

打開瀏覽器，輸入 **http://localhost:5173** 就能看到平台了！

## 🔑 測試帳號

| 角色 | 帳號 | 密碼 |
|------|------|------|
| 管理員 | admin@lkeng.com | LK@dmin2026 |
| 同仁（首登需改密碼）| chen@lkeng.com | E00002 |
| 同仁（首登需改密碼）| lin@lkeng.com | E00003 |
| 同仁（首登需改密碼）| wang@lkeng.com | E00004 |

## 🌐 分享給其他人看（同一個公司內網）

執行 `npm run dev` 後，網路上會出現 `Network: http://192.168.x.x:5173/` 的網址。
**同一個 Wi-Fi/區域網路內的同事**，用他們的瀏覽器打開那個網址也可以連到你的電腦看 demo。

注意：你的電腦必須保持開機，且 `npm run dev` 不能關閉。

## 📦 打包成靜態檔案（要上線時用）

當你要把這個放到公司的 Web Server 時，需要先打包：

```
npm run build
```

執行完會產生一個 **`dist/`** 資料夾。把 dist 裡面的所有檔案上傳到你們的 Web Server（IIS、Apache、Nginx 都可以），網頁就能用了。

預覽打包結果（在上線前測試）：
```
npm run preview
```

## ⚠️ 重要：目前的限制

**這是純前端原型**，所有資料（使用者、課程、學習紀錄、測驗成績）都存在瀏覽器記憶體裡：

- ❌ 重新整理頁面 = 資料消失，回到初始狀態
- ❌ 不同電腦 = 各自獨立的資料，無法共享
- ❌ 沒有真正的帳號驗證機制

**要正式上線，必須搭配後端系統**：
- 一台 Web Server + 資料庫（MySQL / PostgreSQL）
- API 服務（Node.js + Express、或 Python + FastAPI）
- 身分驗證（建議整合公司 AD / SSO）

這個原型適合：
- ✅ 給主管 demo「未來成品大概長這樣」
- ✅ 跟 MIS 部門討論需求
- ✅ 跟外部廠商溝通時當作需求文件
- ✅ 確認 UI / UX 流程是否合適

## 🎬 YouTube 影片功能

YouTube 嵌入功能是**真實可運作的**，使用方式：

1. 用管理員登入 → 後台管理 → 課程管理
2. 編輯任一門課程，每個章節都可貼上 YouTube 網址
3. 建議將 YouTube 影片設為「不公開（Unlisted）」
4. 支援的網址格式：
   - `https://youtu.be/xxxxx`
   - `https://www.youtube.com/watch?v=xxxxx`
   - `https://www.youtube.com/embed/xxxxx`

## 📤 批次匯入使用者

1. 後台 → 使用者管理 → 點「📥 下載範本」取得 Excel 範本
2. 填寫資料：員工編號、姓名、Email、部門、角色（admin 或 user）、密碼（留空則預設為員工編號）
3. 回到頁面點「📤 批次匯入」上傳填好的 Excel
4. 系統會自動跳過已存在的 Email/員工編號，並顯示匯入結果

## 🛠 常用指令

| 指令 | 用途 |
|------|------|
| `npm install` | 安裝套件（第一次使用） |
| `npm run dev` | 啟動開發伺服器（看效果用） |
| `npm run build` | 打包成可上線的靜態檔案 |
| `npm run preview` | 預覽打包後的結果 |

## 📝 修改程式

主要的程式碼都在 `src/learning-platform.jsx` 一個檔案裡，可以用任何文字編輯器打開修改（建議用 **VS Code**，免費下載：https://code.visualstudio.com）。

修改後存檔，瀏覽器會自動重新整理顯示新內容（`npm run dev` 不用重啟）。

## ❓ 常見問題

**Q: 執行 npm install 時出現錯誤**
A: 檢查 Node.js 是否安裝正確（`node -v` 應顯示版本號），並確認電腦可以連線到網路。如果在公司內網有 Proxy，需要先設定：
```
npm config set proxy http://proxy.lkeng.com:8080
npm config set https-proxy http://proxy.lkeng.com:8080
```

**Q: 開啟 http://localhost:5173 沒反應**
A: 確認命令提示字元視窗仍在執行（不能關掉），並檢查防火牆是否阻擋。

**Q: 我關掉視窗就再也打不開了**
A: 重新進到專案目錄執行 `npm run dev` 就好，不用再 `npm install`。

---

**版本**：1.0.0 原型版  
**最後更新**：2026-05-15
