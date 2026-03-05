/**
 * ============================================================
 * 2. Main.gs
 * 主程式入口與 UI 選單
 * ============================================================
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ 海龜 AI 系統')
    .addItem('1. 全面同步 (在庫 + 出倉)', 'syncAllData')
    .addItem('2. 更新行情與海龜診斷', 'updateStockPriceOnly')
    .addSeparator()
    .addItem('同步股利資料庫', 'updateDividendDatabaseTask')
    .addToUi();
}

/**
 * 整合功能：一次同步在庫與出倉
 */
function syncAllData() {
  updateActiveFromInput();
  updateClosedFromInput();
  updateStockPriceOnly(); // 先更新行情以確保成本計算正確
  updateInvestmentPieChart(); // 更新投資分配圖
  SpreadsheetApp.getActiveSpreadsheet().toast("✅ 買賣資料與分配圖已全面同步！");
}

// --- 自動化觸發器設定 ---

/**
 * 執行此函式一次，即可設定每 5 分鐘自動同步
 */
function startAutoSync() {
  // 1. 先刪除舊的觸發器，避免重複建立
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'autoUpdateWrapper') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  // 2. 建立新的觸發器 -> 呼叫包裝函式
  ScriptApp.newTrigger('autoUpdateWrapper')
      .timeBased()
      .everyMinutes(5) // 設定頻率
      .create();
      
  SpreadsheetApp.getUi().alert("已啟動自動同步！\n每 5 分鐘會自動更新一次大盤資料。");
}

/**
 * 執行此函式可停止自動同步
 */
function stopAutoSync() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'autoUpdateWrapper') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  SpreadsheetApp.getUi().alert("已停止自動同步。");
}

/**
 * 給觸發器用的包裝函式
 */
function autoUpdateWrapper() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // 或指定 .getSheetByName("股票")
  updateMarketIndex(sheet);
}

/**
 * 監聽編輯事件
 * 當在「買賣」Sheet 勾選「賣出 (B欄)」時，自動填入「賣出日期 (H欄)」為今天
 */
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const col = range.getColumn();
  const row = range.getRow();
  
  // 檢查是否為「買賣」Sheet (使用 Config 中的常數 SHEET_INPUT)
  if (sheet.getName() !== SHEET_INPUT) return;
  
  // 檢查是否為第 2 欄 (B欄: 賣出勾選框) 且不在標題列
  // 假設標題在第 1 列，所以 row > 1
  if (col === 2 && row > 1) {
    const isChecked = range.getValue() === true;
    const dateCell = sheet.getRange(row, 8); // H 欄是第 8 欄
    
    if (isChecked) {
      // 如果勾選，且日期欄位原本是空的，才自動填入 (避免覆蓋舊紀錄)
      if (dateCell.getValue() === "") {
         const today = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd");
         dateCell.setValue(today);
      }
    } else {
      // (選擇性功能) 如果取消勾選，是否要清除日期？
      // 通常建議保留或讓使用者自己刪，以免誤刪重要紀錄。這裡暫不自動清除。
    }
  }
}

/**
 * 新增股票 (對應原本的 addNewTradeRow 按鈕)
 * 在「買賣」Sheet 的第 2 列 (標題下方) 插入一列空白列
 */
function addNewTradeRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_INPUT);
  if (!sheet) return;
  
  // 取得目前最後一列的位置
  const lastRow = sheet.getLastRow();
  const newRow = lastRow + 1;
  
  // 在 B 欄 (第 2 欄) 的新列插入檢核框
  sheet.getRange(newRow, 2).insertCheckboxes();
  
  // 自動填入 F 欄 (第 6 欄) 為今天日期 (買入日期)
  const today = Utilities.formatDate(new Date(), "GMT+8", "yyyy/MM/dd");
  sheet.getRange(newRow, 6).setValue(today);
  
  // 自動填入 A 欄 (第 1 欄) 為 "在庫中"
  sheet.getRange(newRow, 1).setValue("在庫中");

  // 提示
  ss.toast("已在末尾新增一筆空白交易紀錄");
}
