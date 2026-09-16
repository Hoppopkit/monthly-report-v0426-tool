# 變更日誌

## 1.0.3 (2026-09-16)

### 修正

- **操作指引工作表損壞**：ExcelJS 重寫會寫入無效 DPI，導致 Excel 要求修復；下載前用**上傳原檔**的「操作指引」工作表蓋回
- 仍保留 DutyList／NameList／MR2 由工具寫入的結果，以及已清理的命名範圍

## 1.0.2 (2026-09-16)

### 修正

- **檔案損壞**：注入命名範圍時過濾含 `[1]` 的外部活頁簿參照（ExcelJS 不會保留 externalLinks，貼回去會讓 Excel 判定 corrupt）
- **workbook.xml 順序**：`definedNames` 改插在 `calcPr` 之前，符合 OOXML／Excel 預期

## 1.0.1 (2026-09-16)

### 修正

- **刪除值勤清單**：修正 iOS Safari 文字被擠成窄欄／空白的排版問題
- **命名範圍**：下載時用 JSZip 寫回原檔 `definedNames`（含 `DutyList!$B:$B` 等），避免 MR2 公式失效、報表「變樣」
- **日期**：改以 Excel 序列日寫入，避免出現 04:00 時區偏移

## 1.0.0 (2026-09-16)

### 新增

- 流動版月結報告 V0426：上傳電腦匯出的 MR Form，增刪月底少數值勤
- 新增／刪除值勤（刪除一次最多 2 筆）
- DutyList 依 Duty Code（D01→D10）→ 日期排序寫回
- NameList 同步（新人補職級／英文名；刪除時移除未再使用的編號）
- GitHub Pages 部署腳本 `publish_github_pages.sh`

### 部署

- 正式網址：https://hoppopkit.github.io/monthly-report-v0426-tool/
- 倉庫：https://github.com/Hoppopkit/monthly-report-v0426-tool
