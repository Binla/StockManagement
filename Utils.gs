/**
 * ============================================================
 * 6. Utils.gs
 * 通用工具函式 (格式化、股利計算)
 * ============================================================
 */

function formatAllColumns(sheet, lastRow) {
  const numRows = lastRow - 2;

  // 1. 金額區
  const integerRanges = ["Q3:W", "Y3:Z"];
  integerRanges.forEach(range => {
    sheet.getRange(range + lastRow).setNumberFormat("#,##0");
  });

  // 2. 行情區
  const decimalRanges = ["H3:I", "K3:M", "O3:O"];
  decimalRanges.forEach(range => {
    sheet.getRange(range + lastRow).setNumberFormat("0.00");
  });

  // 3. 百分比區
  sheet.getRange("J3:J" + lastRow).setNumberFormat("0.00%");
  sheet.getRange("X3:X" + lastRow).setNumberFormat("0.00%");

  // 4. 損益變色
  const profitValues = sheet.getRange("W3:W" + lastRow).getValues();
  for (let i = 0; i < profitValues.length; i++) {
    const val = profitValues[i][0];
    const color = val >= 0 ? "#ff0000" : "#1e8e3e";
    sheet.getRange(i + 3, 23).setFontColor(color).setFontWeight("bold");
    sheet.getRange(i + 3, 24).setFontColor(color).setFontWeight("bold");
    sheet.getRange(i + 3, 26).setFontColor(color).setFontWeight("bold");
  }
}

function calculateDividendFromDB(fullId, startDate, endDate, qty) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const divSheet = ss.getSheetByName(SHEET_DIVIDEND); // 使用 Config
  if (!divSheet || divSheet.getLastRow() < 2) return 0;

  const divData = divSheet.getDataRange().getValues();
  let totalDivPerShare = 0;
  const tid = fullId.toString().replace("'", "").trim();
  const startTs = new Date(startDate).getTime();
  const endTs = new Date(endDate).getTime();

  for (let i = 1; i < divData.length; i++) {
    const dId = divData[i][1].toString().replace("'", "").trim();
    const dDateTs = new Date(divData[i][2]).getTime();
    const dAmt = Number(divData[i][3]) || 0;
    
    if (dId === tid && dDateTs >= startTs && dDateTs <= endTs) {
      totalDivPerShare += dAmt;
    }
  }
  return Math.round(totalDivPerShare * qty);
}
