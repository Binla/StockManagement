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

  // A:股票名稱, B:張數, C:投資金額, D:佔比, E:空, F:空, G:總投資金額, H:圖表標籤 (隱藏)
  const outputData = [["股票名稱", "張數", "投資金額", "佔比", "", "", "總投資金額", "圖表標籤"]];
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
    
    // 總投資金額只放在第一行資料 (對應 G 欄)
    const totalCostStr = (rowIndex === 1) ? totalCost : "";
    
    // 將張數與佔比一起包入標籤文字中，因為要顯示在圖表內
    const percentStr = (percent * 100).toFixed(1) + "%";
    const chartLabel = `${label}\n${sheets} 張\n${percentStr}`;
    outputData.push([label, sheets, cost, percent, "", "", totalCostStr, chartLabel]);
    
    rowIndex++;
  }

  if (outputData.length <= 1) {
    ss.toast("無有效的投入成本數據。");
    return;
  }

  // 4. 將摘要表寫入「投資分配圖」Sheet
  chartSheet.getRange(1, 1, outputData.length, 8).setValues(outputData);
  
  // 文字全部置中
  chartSheet.getRange(1, 1, outputData.length, 8).setHorizontalAlignment("center");
  
  // 樣式：表頭粗體與背景 (A~D 欄與 G 欄)
  chartSheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#d9ead3");
  chartSheet.getRange(1, 7, 1, 1).setFontWeight("bold").setBackground("#d9ead3");
  
  // 數字格式
  chartSheet.getRange(2, 2, outputData.length - 1, 1).setNumberFormat("#,##0");   // 張數
  chartSheet.getRange(2, 3, outputData.length - 1, 1).setNumberFormat("#,##0");   // 投資金額
  chartSheet.getRange(2, 4, outputData.length - 1, 1).setNumberFormat("0.00%");   // 佔比
  chartSheet.getRange(2, 7, outputData.length - 1, 1).setNumberFormat("#,##0");   // 總投資金額
  
  // 自動調整欄寬
  chartSheet.autoResizeColumns(1, 8); 
  
  // 若字串太長導致 autoResize 沒算好，手動預留足夠寬度
  chartSheet.setColumnWidth(1, 150); // 確保股票名稱完整
  chartSheet.setColumnWidth(2, 80);  // 張數寬度
  chartSheet.setColumnWidth(3, 120); // 投資金額寬度
  chartSheet.setColumnWidth(4, 90);  // 佔比寬度
  chartSheet.setColumnWidth(7, 130); // 總投資金額寬度
  
  // 隱藏 H 欄 (第 8 欄) 圖表標籤
  chartSheet.hideColumns(8);

  // 動態根據項目多寡決定圖表高度/寬度
  let chartHeight = Math.max(800, outputData.length * 40 + 200);
  let chartWidth = Math.max(1200, outputData.length * 40 + 500);

  // 5. 建立/更新 Pie Chart 
  const chart = chartSheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    // 第一個 range: 圖表標籤 (Col 8 -> H 欄)
    .addRange(chartSheet.getRange(1, 8, outputData.length, 1)) 
    // 第二個 range: 投資金額 (Col 3 -> C 欄)
    .addRange(chartSheet.getRange(1, 3, outputData.length, 1))
    // 位置: 放在 I 欄 (第 9 欄) 之後，避免擋到 G 欄
    .setPosition(2, 9, 0, 0)
    .setOption('title', '在庫持股投資分配圖 (依金額分佈)')
    // 重點 1：設為 label，把隱藏的 H 欄文字塞進圓餅圖區塊裡面
    .setOption('pieSliceText', 'label')  
    // 重點 2：縮小圖表本體以預留空間給文字 (字體不擠壓)
    .setOption('chartArea', {left: 50, top: 50, width: '70%', height: '80%'})
    // 圖例放在右側，讓出左邊空間，避免擁擠
    .setOption('legend', {position: 'right', textStyle: {fontSize: 14}})
    .setOption('width', chartWidth)
    .setOption('height', chartHeight)
    .build();

  chartSheet.insertChart(chart);

  ss.toast("✅ 投資數量分配圖已更新。");
}
