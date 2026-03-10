/**
 * ============================================================
 * 7. Chart.gs
 * 管理統計圖表 (如：投資分配圖)
 * ============================================================
 */

function updateInvestmentPieChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName(SHEET_ACTIVE);
  let chartSheet = ss.getSheetByName(SHEET_CHART);

  if (!invSheet) return;

  // 1. 取得或建立「投資分配圖」Sheet
  if (!chartSheet) {
    chartSheet = ss.insertSheet(SHEET_CHART);
  } else {
    chartSheet.clear(); // 清空舊資料
    // [FIX] 顯式移除所有舊圖表，避免重複
    const charts = chartSheet.getCharts();
    for (let i = 0; i < charts.length; i++) {
        chartSheet.removeChart(charts[i]);
    }
  }

  const lastRow = invSheet.getLastRow();
  if (lastRow < 3) {
    ss.toast("在庫目前無持股，無法產生分配圖。");
    return;
  }

  // 2. 從「在庫」抓取持股名稱、代號與總成本 (Column R)
  // Column A: 名稱, Column B: 代號, Column R: 總成本 (第 18 欄)
  const data = invSheet.getRange(3, 1, lastRow - 2, 18).getValues();
  
  // 3. 彙整資料 (同一支股票加總)
  const distribution = {};
  for (let i = 0; i < data.length; i++) {
    const name = data[i][0].toString().trim();
    const id = data[i][1].toString().trim();
    const shares = Number(data[i][2]) || 0; // Column C is index 2
    const cost = Number(data[i][17]) || 0; // Column R is index 17
    
    if (!id || cost <= 0) continue;
    
    const label = `${name} (${id})`;
    if (!distribution[label]) {
      distribution[label] = { cost: 0, shares: 0 };
    }
    distribution[label].cost += cost;
    distribution[label].shares += shares;
  }

  // D欄改為佔比 (%), E欄改為總投資金額, F欄為圖表標籤 (隱藏)
  const outputData = [["股票名稱", "張數", "投資金額", "佔比", "總投資金額", "圖表標籤"]];
  let totalCost = 0;
  for (let label in distribution) {
    totalCost += distribution[label].cost;
  }

  let rowIndex = 1;
  for (let label in distribution) {
    const cost = distribution[label].cost;
    const shares = distribution[label].shares;
    const sheets = Math.floor(shares / 1000); // 無小數點
    const percent = totalCost > 0 ? cost / totalCost : 0; // 佔比
    
    // 總投資金額只放在第一行資料
    const totalCostStr = (rowIndex === 1) ? totalCost : "";
    outputData.push([label, sheets, cost, percent, totalCostStr, `${label} (${sheets} 張)`]);
    
    rowIndex++;
  }

  if (outputData.length <= 1) {
    ss.toast("無有效的投入成本數據。");
    return;
  }

  // 4. 將摘要表寫入「投資分配圖」Sheet
  chartSheet.getRange(1, 1, outputData.length, 6).setValues(outputData);
  
  // 樣式與數字格式
  chartSheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#d9ead3");
  chartSheet.getRange(2, 2, outputData.length - 1, 1).setNumberFormat("#,##0");   // 張數
  chartSheet.getRange(2, 3, outputData.length - 1, 1).setNumberFormat("#,##0");   // 投資金額
  chartSheet.getRange(2, 4, outputData.length - 1, 1).setNumberFormat("0.00%");   // 佔比
  chartSheet.getRange(2, 5, outputData.length - 1, 1).setNumberFormat("#,##0");   // 總投資金額
  
  // 自動調整欄寬與隱藏輔助圖表標籤
  // 為了確保「總投資金額」和「佔比」夠寬，直接給予設定
  chartSheet.autoResizeColumns(1, 5); 
  chartSheet.setColumnWidth(4, 80); // 佔比寬一點點
  chartSheet.setColumnWidth(5, 120); // 總金額寬一點點
  chartSheet.hideColumns(6);

  // 動態根據項目多寡決定圖表高度/寬度
  // 增加基礎長寬，確保標籤與數值完整顯示不會被切斷
  let chartHeight = Math.max(700, outputData.length * 40 + 200);
  let chartWidth = Math.max(1000, outputData.length * 30 + 500);

  // 5. 建立/更新 Pie Chart 
  const chart = chartSheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    // 第一個 range: 圖表標籤 (Col 6) -> EX: "台積電 (5 張)"
    .addRange(chartSheet.getRange(1, 6, outputData.length, 1)) 
    // 第二個 range: 投資金額 (Col 3) -> 用金額作為圖餅大小
    .addRange(chartSheet.getRange(1, 3, outputData.length, 1))
    .setPosition(2, 7, 0, 0)
    .setOption('title', '在庫持股投資分配圖 (依金額分佈, 顯示張數)')
    .setOption('is3D', true)
    .setOption('pieSliceText', 'value-and-percentage')
    .setOption('width', chartWidth)
    .setOption('height', chartHeight)
    .setOption('legend', {textStyle: {fontSize: 14}})
    .build();

  chartSheet.insertChart(chart);

  ss.toast("✅ 投資數量分配圖已更新。");
}
