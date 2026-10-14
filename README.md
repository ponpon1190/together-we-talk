# 一起聽見 Together We Hear

The Hope Church Phase 1 互動發聲網站——森林舞台、動物排隊登台發言、問卷收集與 9:16 分享圖卡生成。

部署在 GitHub Pages，純靜態網頁（`index.html` + `assets/`）。

## 後端資料串接

問卷送出與發聲牆資料走 Google Apps Script：部署好後，把網址貼進 `index.html` 裡的 `GAS_ENDPOINT` 變數即可。Apps Script 原始碼在主專案的 `02_WEBSITE/together_we_talk/backend/Code.gs`。
