# v3.9.6.1 更新

請先閱讀「部署與驗證說明.txt」，其中列出本版更新與實測範圍。以下保留原版說明。

# 電子閱讀護照 v3.9.4

## 本版新增：管理者批次審核
- 系統管理者可在「心得審核」勾選多筆待審核心得。
- 支援「全選目前待審核」、「批次通過」、「批次退回」。
- 批次退回時必須輸入統一退回原因，學生可依原因修改後重新送審。
- 僅 pending 待審核資料可批次處理；已通過或已退回資料會自動略過。
- 單次最多 200 篇，避免大量操作造成 Apps Script 執行逾時。
- 班級教師仍保留逐篇審核；批次功能限定系統管理者。

# 電子閱讀護照 v3.8.9

首頁橫幅完整顯示版。

- 桌機、平板、手機皆以原始比例完整顯示橫幅
- 移除固定高度與 cover 裁切
- 保留 v3.8.9 全部閱讀護照、學生任務、教師審核與管理功能
- GitHub Pages：將 ZIP 內容直接放在 Repository 根目錄


## v3.8.9 安全修正
- 星等固定限制在 0～5，避免 `.repeat()` 收到負時間戳造成 `Invalid count value`。
- 後端若發現日期物件誤落在 rating/pages/viewCount 等數值欄，會安全轉為 0。
- `ReadingRecords` 仍以第一列欄位名稱映射，不依賴固定欄號。
- 新增 `validateSpreadsheetSchema()`：可在 Apps Script 手動執行，檢查缺欄、重複欄與異常 rating。
- 更新後請執行 `setupSpreadsheet()`，再執行 `validateSpreadsheetSchema()`，最後建立新版本重新部署 Web App。

## 更新步驟（v3.8.9）
1. GitHub Pages：將 ZIP 根目錄全部檔案覆蓋 Repository 根目錄。
2. Apps Script：用 `apps-script/Code.gs` 完整覆蓋舊版。
3. 執行 `setupSpreadsheet()`。
4. 執行 `validateSpreadsheetSchema()`；回傳 `OK` 最理想，若提示異常 rating，請檢查 ReadingRecords 該欄資料。
5. Apps Script「管理部署作業 → 編輯 → 新版本 → 部署」。
6. 瀏覽器以 Ctrl+F5 強制重新整理。


## v3.8.9 修正
- 管理後台心得圖片改由 Apps Script 讀取 Drive 檔案並顯示縮圖，不受公開分享政策影響。
- 審核視窗「通過並計入 1 篇」按鈕恢復清楚的綠色樣式。
- 前台統計卡顯示「審核通過篇數」與「發表者人數」。

- v3.8.9：修正心得圖片縮圖、審核通過按鈕、前台審核通過篇數與發表者人數統計。


## v3.8.9 修正
- 校園閱讀之星、年級閱讀量、每月閱讀趨勢、書籍分類分布：改採所有已審核通過紀錄的匿名彙總。
- 最新心得與熱門文章仍只顯示學生允許公開的心得內容。
- 排行榜預設依審核通過心得篇數排序。
- Chart.js 載入失敗時顯示友善備援，不會讓整段消失。


## v3.8.9 學生身分一致性修正

- 學生登入後會取得後端簽章 identityToken。
- 繳交指定任務或自由心得時，後端同時核對 studentId、studentName、className 與 identityToken。
- 同一 studentId 若在 Students 出現兩筆以上，系統會拒絕登入/送出，避免 A 生心得被寫到 B 生名下。
- 可在 Apps Script 手動執行 `validateStudentRoster()` 檢查重複 studentId。


## v3.9.0 班級＋姓名唯一身分版
- 學生登入與身分驗證改以「班級＋姓名」為唯一依據。
- 原始 Students.studentId 不必全校唯一，可保留 101、102 等班內編號。
- 系統自動產生 SC- 開頭的內部學生 key，供新心得紀錄使用。
- 舊心得以「班級＋姓名」自動相容，不需手動搬移。
- validateStudentRoster() 現在只檢查同班同名；若同班確有同名學生，需在名冊姓名加可辨識註記。


## v3.9.1 班級格式相容修正
- 學生登入班級改為一年1班至六年8班下拉選單。
- 後端自動將 607、六年7班、六年七班、六年級7班、6年7班等格式統一為 607。
- 姓名會移除前後、全形及不可見空白後再比對。


## v3.9.2 API 部署診斷
- POST 同時在 query string 與 JSON body 傳送 action，提高 Apps Script Web App 相容性。
- 可直接開啟 `API_URL?action=health`，正常應回傳 JSON 且版本為 3.9.2。
- 若 health 網址仍回傳 Google HTML，代表部署權限或 config.js 的部署網址尚未更新，並非管理金鑰錯誤。


## v3.9.5 語音閱讀心得
- 自由閱讀心得支援文字、圖片、語音三種形式。
- 學生可直接使用瀏覽器麥克風錄音，或上傳 MP3/M4A/WAV/WebM。
- 語音限制 10MB、最長 5 分鐘。
- 語音存放 Google Drive「電子閱讀護照-語音心得」。
- 教師/管理者審核視窗可直接播放語音。
- 退回後可保留原語音或重新錄製/上傳，沿用原 recordId。
- 更新 Code.gs 後請執行 setupSpreadsheet() 一次，讓 ReadingRecords 自動補上語音欄位。
