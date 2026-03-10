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

  const chartData = [["股票名稱", "張數", "投資金額"]];
  let totalCost = 0;
  for (let label in distribution) {
    totalCost += distribution[label].cost;
  }

  for (let label in distribution) {
    const cost = distribution[label].cost;
    const shares = distribution[label].shares;
    const sheets = Math.floor(shares / 1000); // 無小數點
    chartData.push([label, sheets, cost]);
  }

  if (chartData.length <= 1) {
    ss.toast("無有效的投入成本數據。");
    return;
  }

  // 4. 將摘要表寫入「投資分配圖」Sheet
  chartSheet.getRange(1, 1, chartData.length, 3).setValues(chartData);
  
  // D1 寫入總投資金額
  chartSheet.getRange(1, 4).setValue("總投資金額");
  chartSheet.getRange(2, 4).setValue(totalCost);
  
  // 樣式與數字格式
  chartSheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#d9ead3");
  chartSheet.getRange(2, 2, chartData.length - 1, 1).setNumberFormat("#,##0");   // 張數
  chartSheet.getRange(2, 3, chartData.length - 1, 1).setNumberFormat("#,##0");   // 投資金額
  chartSheet.getRange(2, 4).setNumberFormat("#,##0");                            // 總投資金額
  chartSheet.autoResizeColumns(1, 4);

  // 5. 建立/更新 Pie Chart 
  const chart = chartSheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    // 第一個 range: 股票名稱 (Col 1)
    .addRange(chartSheet.getRange(1, 1, chartData.length, 1)) 
    // 第二個 range: 改為 張數 (Col 2)，在圓餅圖上顯示
    .addRange(chartSheet.getRange(1, 2, chartData.length, 1))
    .setPosition(2, 6, 0, 0)
    .setOption('title', '在庫持股投資數量圖 (依張數)')
    .setOption('is3D', true)
    .setOption('pieSliceText', 'value-and-percentage')
    .setOption('width', 800)
    .setOption('height', 600)
    .build();

  chartSheet.insertChart(chart);

  ss.toast("✅ 投資數量分配圖已更新。");
}
