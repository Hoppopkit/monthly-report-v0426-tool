# 月結報告 V0426 流動版

在手機／平板瀏覽器增刪月底少數值勤後，下載可交的 MR Form Excel。  
**不會**連資料庫；請先用桌面程式匯出完整月結報告，再上傳此工具修改。

## 功能

- 上傳電腦已匯出的 `MR Form V0426 ….xlsx`
- **新增** 1～2 筆值勤（含隊員）
- **刪除** 勾選的值勤（一次最多 2 筆）
- 依 **Duty Code（D01→D10）→ 日期** 排序寫回 DutyList
- 同步 NameList（新人需填職級＋英文名；刪除時僅移除不再出現的編號）
- 只改 `DutyList` / `NameList`，其餘工作表（含 MR2 公式）保留

## 在裝置上使用

### 方式 A：GitHub Pages（推薦）

**正式網址：** https://hoppopkit.github.io/monthly-report-v0426-tool/

1. 用 Safari／Chrome 開啟上述網址（可「加入主畫面」）
2. 上傳 xlsx → 增刪 → 下載；檔案在 **檔案 → 下載項目**

> 需能連線載入 ExcelJS／JSZip CDN（`cdn.jsdelivr.net`）。

**維護者：** 倉庫 https://github.com/Hoppopkit/monthly-report-v0426-tool — Pages 來源為 `main` 分支根目錄。

#### 首次部署到 GitHub Pages（做法二：獨立 repo）

本資料夾已是獨立 git 倉庫。在終端機執行：

```bash
# 若尚未登入 GitHub（只需做一次）
gh auth login

cd /Users/macbookair/Documents/Python/monthly-report-v0426-tool
./publish_github_pages.sh
```

腳本會：建立 `Hoppopkit/monthly-report-v0426-tool` 公開倉庫 → push → 啟用 Pages。  
約 1–3 分鐘後開啟：https://hoppopkit.github.io/monthly-report-v0426-tool/

**若沒有 `gh`：** `brew install gh` 後再執行上述步驟。

**手動方式（不用腳本）：**

1. 在 GitHub 新建 public 倉庫 `monthly-report-v0426-tool`（不要勾選 README）
2. `git remote add origin https://github.com/Hoppopkit/monthly-report-v0426-tool.git`
3. `git push -u origin main`
4. 倉庫 **Settings → Pages** → Branch `main`、Folder `/ (root)` → Save

### 方式 B：本機簡易伺服器

```bash
cd monthly-report-v0426-tool
python3 -m http.server 8080
```

同一 Wi‑Fi 下手機開啟 `http://電腦IP:8080`。

### 方式 C：直接開 index.html

部分流動裝置對 `file://` 限制較多（上傳／CDN／下載），**不建議**；請用 A 或 B。

## 建議流程

1. 電腦：Duty 管理 → **匯出月結報告0426** → 存成完整 MR Form  
2. 傳到手機（AirDrop／iCloud／郵件）  
3. 開啟本工具 → 上傳該檔  
4. 月底多出的 1～2 筆：填日期、名稱、地點、Code、隊員 → **加入此值勤**  
5. 若要拿掉誤加項目：勾選 → **刪除**  
6. **下載修改後的 MR Form** → 用 Excel／Numbers 確認 MR2 頁面正常後再呈交  

## 欄位規則（與桌面一致）

| 項目 | 規則 |
|------|------|
| DutyList 列 | 約第 2–278 列 |
| NameList | 約第 4–65 列 |
| 同一值勤多列 | Date / Nature / Location **每列都填** |
| Code / Duty Case | **僅該值勤首列**；Duty Case 固定 `0` |
| Hours / DH | 4 / 0 / 0 |
| 排序 | Code 數字 01→10，同 Code 再依日期早→晚 |

## 注意

- 只支援 `.xlsx`（不支援舊 `.xls`）
- 下載後請用電腦 Excel 打開，確認 MR2 公式／頁數仍正確（第一次使用時建議對照未修改前的檔）
- 資料僅在瀏覽器記憶體處理，不會上傳到伺服器（GitHub Pages 只提供靜態網頁）
- 一次刪除超過 2 筆會被擋下，避免誤刪整月資料

## 檔案說明

| 檔案 | 說明 |
|------|------|
| `index.html` | 介面 |
| `app.js` | 流程與 UI |
| `excel-io.js` | 讀寫 xlsx |
| `dutylist.js` / `namelist.js` | 值勤／人員邏輯 |
| `constants.js` | 與桌面匯出對齊的常數 |

## 版本

見 `VERSION`。
