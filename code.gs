/**
 * ============================================================
 * 🧩 追管系統互動式導覽 - 後端主控 (Code.gs)
 * 功能：
 * 1️⃣ 控制前端頁面 include
 * 2️⃣ 提供導覽資料 (getGuideData)
 * 3️⃣ 建立入口 doGet()
 * ============================================================
 */

/** ✅ 載入 HTML 子頁 **/
/** ✅ 支援巢狀 include()（連 guidePage 裡面的 include 也會被處理） */
function include(filename) {
  // 把 HTML 檔案當成 Template 再 evaluate 一次
  const content = HtmlService.createTemplateFromFile(filename).evaluate().getContent();
  return content;
}

/** ✅ 網頁入口 **/
function doGet(e) {
  const output = HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('追管系統互動式導覽')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return output;
}

// 後端讀取導覽資料
function getGuideData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ============================================================
  // ① 讀取 GuideData（你原本的部分）
  // ============================================================
  const sheet = ss.getSheetByName('GuideData');
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return {};

  const headers = rows[0];
  const data = {};
  const headerIndex = {};
  headers.forEach((h, i) => headerIndex[h.trim()] = i);

  let mainPage = "";

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];

    const pageUrl = r[headerIndex["畫面圖片網址"]];
    if (!pageUrl) continue;

    const pagePath = r[headerIndex["畫面圖片路徑"]] || "";
    const x = Number(r[headerIndex["熱區X"]] || 0);
    const y = Number(r[headerIndex["熱區Y"]] || 0);
    const w = Number(r[headerIndex["寬度"]] || 0);
    const h = Number(r[headerIndex["高度"]] || 0);
    const title = r[headerIndex["功能名稱"]] || "";

    const text = HtmlService.createHtmlOutput(r[headerIndex["顯示文字"]] || "").getContent();

    const type = r[headerIndex["類型"]] || "";
    const targetPath = r[headerIndex["目標畫面路徑"]] || "";
    const targetUrl  = r[headerIndex["目標畫面網址"]] || "";
    const options = r[headerIndex["選項清單"]] || "";
    const frameScroll = r[headerIndex["Frame滾動方向"]] || "both";

    const optionDesc = (r[headerIndex["各選項對應說明"]] || "").toString();

    const isMain = String(r[headerIndex["是否主畫面"]] || "").trim() === "✅";
    if (isMain && !mainPage) mainPage = pageUrl;

    if (!data[pageUrl]) data[pageUrl] = [];

    const tags = (r[headerIndex["標籤"]] || "")
      .toString()
      .split(',')
      .map(t => t.trim())
      .filter(Boolean);

    data[pageUrl].push({
      pagePath,
      x, y, w, h,
      title, text, type,
      frameScroll,
      targetPath,
      targetUrl,
      options, optionDesc,
      tags
    });
  }


  // ============================================================
  // ② 新增：讀取「設定」工作表
  // ============================================================
  const settingSheet = ss.getSheetByName("設定");
  let settings = {};

  if (settingSheet) {
    const settingValues = settingSheet.getDataRange().getValues();

    settingValues.forEach(row => {
      const key = row[0];   // A 欄
      if (!key) return;

      // 從 B 開始往右收集（直到空值停止）
      const arr = row.slice(1); 

      settings[key] = arr;
    });
  }


  // ============================================================
  // ③ 原結構 + 新增設定 settings
  // ============================================================
  return {
    pages: data,
    mainPage,
    settings   // ⬅ 你要的功能：通關問題 / 通關答案 / 開發者密語都從這裡取
  };
}

