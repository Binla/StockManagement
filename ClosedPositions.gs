/**
 * ============================================================
 * 4. ClosedPositions.gs
 * 管理「出倉」Sheet 與已實現損益
 * ============================================================
 */

function updateClosedFromInput() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName(SHEET_INPUT);
  const closedSheet = ss.getSheetByName(SHEET_CLOSED);
  if (!inputSheet || !closedSheet) return;

  const lastInputRow = inputSheet.getLastRow();
  if (lastInputRow < 2) return;
  
  const inputData = inputSheet.getRange(2, 1, lastInputRow - 1, 9).getValues();
  let outputToClosed = [];

  for (let i = 0; i < inputData.length; i++) {
    const row = inputData[i];
    // 只抓已賣出或打勾的
    if (row[0] === "已賣出" || row[1] === true) { 
      const [status, isSell, name, id, qty, bDate, bPrice, sDate, sPrice] = row;
      
      const bTotal = Math.floor(bPrice * qty * (1 + FEE_RATE));
      const sTotal = sPrice * qty;
      const netIn  = sTotal - Math.floor(sTotal * (FEE_RATE + TAX_RATE));
      
      // 持有期股利精確對照 (呼叫 Utils.gs)
      const div = calculateDividendFromDB(id, bDate, sDate, qty);
      const profit = netIn - bTotal + div;
      const roi = bTotal > 0 ? profit / bTotal : 0;
      
      outputToClosed.push([
        name, "'" + id, qty, bDate, sDate, 
        Math.floor((new Date(sDate)-new Date(bDate))/86400000), 
        bPrice, sPrice, bTotal, netIn, div, profit, roi
      ]);
      
      // 如果是在庫中打勾，則更新 A欄
      if (row[1] === true && status === "在庫中") {
        inputSheet.getRange(i + 2, 1).setValue("已賣出");
      }
    }
  }

  // 強制刷新
  const currentClosedLastRow = closedSheet.getLastRow();
  if (currentClosedLastRow >= 3) {
    closedSheet.getRange(3, 1, currentClosedLastRow - 2, 13).clear(); 
  }

  if (outputToClosed.length > 0) {
    outputToClosed.sort((a, b) => new Date(a[4]) - new Date(b[4]));
    closedSheet.getRange(3, 1, outputToClosed.length, 13).setValues(outputToClosed).setHorizontalAlignment("center");
    closedSheet.getRange(3, 13, outputToClosed.length, 1).setNumberFormat("0.00%");
    
    finalizeClosedSummary(closedSheet, outputToClosed.length + 2);
    ss.toast("✅ 出倉資料已強制刷新並重新對齊股利");
  }
}

/**
 * [輔助] 出倉顏色校正與總計
 */
function finalizeClosedSummary(sheet, lastRow) {
  if (lastRow < 3) return;
  const range = sheet.getRange(3, 12, lastRow - 2, 2);
  const vals = range.getValues();
  for (let i = 0; i < vals.length; i++) {
    const color = (Number(vals[i][0]) || 0) >= 0 ? "#ff0000" : "#1e8e3e";
    sheet.getRange(i + 3, 12, 1, 2).setFontColor(color).setFontWeight("bold");
  }
  // 計算 I1~L1
  const totals = sheet.getRange(3, 9, lastRow - 2, 4).getValues();
  let sums = [0, 0, 0, 0];
  totals.forEach(r => { sums[0]+=r[0]; sums[1]+=r[1]; sums[2]+=r[2]; sums[3]+=r[3]; });
  sheet.getRange("I1:L1").setValues([[sums[0], sums[1], sums[2], sums[3]]]).setNumberFormat("#,##0");
  const roi = sums[0] > 0 ? sums[3] / sums[0] : 0;
  const tColor = sums[3] >= 0 ? "#ff0000" : "#1e8e3e";
  sheet.getRange("M1").setValue(roi).setNumberFormat("0.00%").setFontColor(tColor).setFontWeight("bold");
  sheet.getRange("L1").setFontColor(tColor).setFontWeight("bold");
}
