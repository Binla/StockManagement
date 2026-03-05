/**
 * ============================================================
 * 3. Inventory.gs
 * 管理「在庫」Sheet 的功能
 * ============================================================
 */

function updateStockPriceOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ACTIVE);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;

  ss.toast("正在抓取行情與執行海龜診斷...", "執行中");
  const data = sheet.getRange(3, 1, lastRow - 2, 7).getValues();
  
  // 快取物件：避免同一支股票重複抓取
  const stockCache = {};

  // [Step 1] 預先統計每支股票的目前總持倉
  const stockHoldings = {}; 
  const uniqueIds = [];
  for (let i = 0; i < data.length; i++) {
     const id = data[i][1].toString().trim().replace("'", "").toUpperCase(); // Normalize to UpperCase
     if (!id) continue;
     const shares = Number(data[i][2]) || 0;
     stockHoldings[id] = (stockHoldings[id] || 0) + shares;
     if (!uniqueIds.includes(id)) uniqueIds.push(id);
  }
  
  // [NEW Step 1.5] 批次抓取 Snapshot 資料 (現價, 昨收, EPS, PE)
  const snapshotBatch = getYahooBatchQuotes(uniqueIds);

  // [Step 2] 準備批次寫入的容器
  const marketValuesHIA = [];      // H, I, J (現價, 漲跌, 漲跌%)
  const turtleValuesKMO = [];      // K, L, M, N, O (海龜指標)
  const profitFormulasPZ = [];     // P~Z (損益公式)
  const fundamentalValuesAAAC = []; // AA, AB, AC (EPS, PE, 診斷)
  const technicalValuesADAE = [];  // AD, AE (RSI, 建議倉位)
  
  const rowColors = []; // 儲存需要特別變色的列資訊

  // [Step 3] 主迴圈：計算數據
  for (let i = 0; i < data.length; i++) {
    const row = i + 3;
    const fullId = data[i][1].toString().trim().replace("'", "");
    const buyPrice = Number(data[i][6]) || 0;
    const qty = Number(data[i][2]) || 0;

    if (!fullId) {
      marketValuesHIA.push([null, null, null]);
      turtleValuesKMO.push([null, null, null, null, null]);
      profitFormulasPZ.push(new Array(11).fill(null));
      fundamentalValuesAAAC.push([null, null, null]);
      technicalValuesADAE.push([null, null]);
      continue;
    }

    // A. 優先從 Snapshot 取得基礎行情 (快取)
    const upperId = fullId.toUpperCase();
    const snapshot = snapshotBatch[upperId] || {};
    const cp = snapshot.price || 0;
    
    // B. 抓取 海龜指標與 RSI (需要歷史資料)
    if (!stockCache.hasOwnProperty(upperId)) {
        stockCache[upperId] = getYahooCompleteData(fullId);
    }
    const yahoo = stockCache[upperId];
    
    // 確保有資料
    if (cp > 0 || (yahoo && yahoo.price > 0)) {
      const currentPrice = cp || (yahoo ? yahoo.price : 0);
      
      // [Robust Fix] 判斷資料來源與處理平盤 (0)
      let changeVal = 0;
      let changePct = 0;
      let dataSource = "NONE";

      if (snapshot.change !== undefined) {
        changeVal = snapshot.change;
        changePct = snapshot.changePercent;
        dataSource = "BATCH_SNAPSHOT";
      } else if (yahoo && yahoo.change !== undefined) {
        changeVal = yahoo.change;
        changePct = yahoo.changePercent;
        dataSource = "FALLBACK_CHART";
      }
      
      const h20 = yahoo ? yahoo.high : currentPrice;
      const n = yahoo ? yahoo.nValue : (currentPrice * 0.015);
      const rsiVal = yahoo ? yahoo.rsi : 50;
      
      console.log(`[${dataSource}] ${upperId}: Price=${currentPrice}, Change=${changeVal}, ATR(N)=${n.toFixed(2)}`);

      // H~J: 行情
      marketValuesHIA.push([currentPrice, changeVal, changePct]);

      // K~O: 海龜指標
      // ... (rest of indicator logic unchanged) ...
      const mStop = yahoo ? yahoo.low10 : (h20 - 2 * n);
      const turtleSignal = (currentPrice < mStop) ? "EXIT" : (currentPrice >= h20 ? (rsiVal > RSI_OVERBOUGHT ? "WAIT" : "ADD") : "HOLD");
      
      let turtleStatus = "【持有】";
      let turtleColor = "#000000";
      let turtleBg = null;
      let mStopBg = null;
      let mStopColor = "#000000";

      if (turtleSignal === "EXIT") {
        turtleStatus = "【清倉】";
        turtleColor = "#1e8e3e";
        turtleBg = "#d9ead3";
        mStopBg = "#d9ead3";
        mStopColor = "#1e8e3e";
      } else if (turtleSignal === "ADD") {
        turtleStatus = "【加碼】";
        turtleColor = "#cc0000";
        turtleBg = "#f4cccc";
      } else if (turtleSignal === "WAIT") {
        turtleStatus = "⚠️ 暫緩 (RSI過熱)";
        turtleColor = "#e69138";
        turtleBg = "#fce5cd";
      }
      
      turtleValuesKMO.push([
        Math.round((buyPrice - (2 * n)) * 100) / 100, 
        h20, 
        mStop, 
        turtleStatus, 
        h20
      ]);

      // P~Z: 損益與手續費 (全公式)
      profitFormulasPZ.push([
        `=G${row}*C${row}`,           // P
        `=P${row}*${FEE_RATE}`,       // Q
        `=P${row}+Q${row}`,           // R
        `=H${row}*C${row}`,           // S
        `=S${row}*${FEE_RATE}`,       // T
        `=S${row}*${TAX_RATE}`,       // U
        `=S${row}-T${row}-U${row}`,   // V
        `=V${row}-R${row}`,           // W
        `=IFERROR(W${row}/R${row}, 0)`,// X
        `=IFERROR(SUMIFS('${SHEET_DIVIDEND}'!D:D, '${SHEET_DIVIDEND}'!B:B, B${row}, '${SHEET_DIVIDEND}'!C:C, ">="&D${row}) * C${row}, 0)`, // Y
        `=W${row}+Y${row}`            // Z
      ]);

      // AA~AC: 基本面 (優先用 Snapshot, 其次用 Fallback)
      let eps = (snapshot.eps !== undefined && snapshot.eps !== 0) ? snapshot.eps : (yahoo ? yahoo.eps : 0);
      let pe = (snapshot.pe !== undefined && snapshot.pe !== 0) ? snapshot.pe : (yahoo ? yahoo.pe : 0);
      
      // 如果 Snapshot 沒有基本面，又是台股，嘗試用 Open Data 補
      if (eps === 0 && pe === 0 && !fullId.startsWith("00")) {
        const fundamental = getYahooQuote(fullId, currentPrice);
        eps = fundamental.eps;
        pe = fundamental.pe;
      }
      
      fundamentalValuesAAAC.push([
        eps,
        pe,
        `=IFS(LEFT(B${row},2)="00","🔹 ETF",AND(AA${row}=0,AB${row}=0,IFERROR(SEARCH(".TWO",B${row}),0)>0),"⚠️ 上櫃缺資料",AND(AA${row}=0,AB${row}=0),"❓ 無數據",AA${row}<0,"⚠️ 虧損中",AB${row}>25,"⚠️ 昂貴",TRUE,"✅ 體質佳")`
      ]);

      // AD~AE: RSI 與建議倉位
      let unitLots = 0;
      if (n > 0) {
        const riskPerTrade = TOTAL_CAPITAL * 0.01;
        const currentHeld = stockHoldings[fullId] || 0;
        const remainingCap = Math.max(0, TOTAL_CAPITAL - (currentHeld * cp));
        // 將精度從 0.1 (百股) 提高到 0.01 (十股)
        unitLots = Math.round((Math.min(riskPerTrade/n, remainingCap/cp) / 1000) * 100) / 100;
      }
      technicalValuesADAE.push([rsiVal, unitLots]);

      // 收集變色邏輯所需的資訊
      const estProfit = (currentPrice * qty * (1 - FEE_RATE - TAX_RATE)) - (buyPrice * qty * (1 + FEE_RATE));
      const trendColor = (changeVal > 0) ? "#ff0000" : (changeVal < 0 ? "#1e8e3e" : "#000000");
      rowColors.push({
        row: row, 
        profitColor: (estProfit >= 0 ? "#ff0000" : "#1e8e3e"), 
        trendColor: trendColor,
        turtleColor: turtleColor,
        turtleBg: turtleBg,
        mStopColor: mStopColor,
        mStopBg: mStopBg,
        rsi: rsiVal
      });

    } else {
      marketValuesHIA.push([null, null, null]);
      turtleValuesKMO.push([null, null, null, null, null]);
      profitFormulasPZ.push(new Array(11).fill(null));
      fundamentalValuesAAAC.push([null, null, null]);
      technicalValuesADAE.push([null, null]);
    }
  }

  // [Step 4] 批次寫入 Sheet (顯著提升速度)
  const numRows = marketValuesHIA.length;
  sheet.getRange(3, 8, numRows, 3).setValues(marketValuesHIA);
  sheet.getRange(3, 11, numRows, 5).setValues(turtleValuesKMO);
  sheet.getRange(3, 16, numRows, 11).setValues(profitFormulasPZ);
  sheet.getRange(3, 27, numRows, 3).setValues(fundamentalValuesAAAC);
  sheet.getRange(3, 30, numRows, 2).setValues(technicalValuesADAE);

  // [Step 5] 批次設定樣式
  rowColors.forEach(item => {
    sheet.getRange(item.row, 8).setFontWeight("bold"); // 現價加粗
    sheet.getRange(item.row, 9, 1, 2).setFontColor(item.trendColor); // I, J 漲跌變色
    
    // M 欄 (停利) 與 N 欄 (診斷) 變色與背景
    sheet.getRange(item.row, 13).setFontWeight("bold").setFontColor(item.mStopColor).setBackground(item.mStopBg);
    sheet.getRange(item.row, 14).setFontColor(item.turtleColor).setBackground(item.turtleBg);

    sheet.getRange(item.row, 23, 1, 2).setFontColor(item.profitColor).setFontWeight("bold"); // W, X 損益變色
    // Z 欄 (總損益) 改用條件式格式化，故此處移除單獨設定
    
    // RSI 變色
    const rsiCell = sheet.getRange(item.row, 30);
    if (item.rsi > RSI_OVERBOUGHT) rsiCell.setFontColor("#e69138").setFontWeight("bold");
    else if (item.rsi < 30) rsiCell.setFontColor("#1e8e3e").setFontWeight("bold");
    else rsiCell.setFontColor("#000000").setFontWeight("normal");
  });

  updateMarketIndex(sheet);
  applyStockGroupingBorders(sheet);
  applyConditionalFormatting(sheet);

  // 更新總結公式 (R1~Z1)
  const totalFormulas = [[
    `=SUM(R3:R${lastRow})`, // R1: 總支出
    `=SUM(S3:S${lastRow})`, // S1: 總市值
    `=SUM(T3:T${lastRow})`, // T1: 總手續費
    `=SUM(U3:U${lastRow})`, // U1: 總稅
    `=SUM(V3:V${lastRow})`, // V1: 總淨收
    `=SUM(W3:W${lastRow})`, // W1: 總損益
    `=IFERROR(W1/R1, 0)`,   // X1: 總報酬率 (總損益/總支出)
    `=SUM(Y3:Y${lastRow})`, // Y1: 總股利
    `=SUM(Z3:Z${lastRow})`  // Z1: 含息總損益
  ]];
  sheet.getRange(1, 18, 1, 9).setFormulas(totalFormulas);

  ss.toast("✅ 行情與海龜診斷更新完成！");
}

/**
 * 輔助函式：自動設定 AB, AC 欄的條件式格式化 (顏色)
 */
function applyConditionalFormatting(sheet) {
  // 移除這欄原本相關的規則，避免重複堆疊
  // 只清除 AC 欄 (第29欄) 與 AB 欄 (第28欄) 的規則，不影響全表
  const rangeAC = sheet.getRange("AC3:AC");
  const rangeAB = sheet.getRange("AB3:AB");
  
  const rangeZ = sheet.getRange("Z3:Z");
  
  const rules = sheet.getConditionalFormatRules();
  // 過濾掉受管理欄位的規則，重新建立
  const filteredRules = rules.filter(rule => {
    const ranges = rule.getRanges();
    return !ranges.some(r => ["AC3:AC", "AB3:AB", "Z3:Z"].includes(r.getA1Notation()));
  });
  
  const newRules = [];
  
  // 1. BLUE for ETF
  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("ETF")
    .setBackground("#cfe2f3") 
    .setFontColor("#1155cc")  
    .setRanges([rangeAC])
    .build());
    
  // 2. RED for 虧損
  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("虧損")
    .setBackground("#f4cccc") // 紅底
    .setFontColor("#cc0000")  // 紅字
    .setRanges([rangeAC])
    .build());

  // 3. GREEN for 體質佳
  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("體質佳")
    .setBackground("#d9ead3") // 綠底
    .setFontColor("#1e8e3e")  // 綠字
    .setRanges([rangeAC])
    .build());
    
  // 4. ORANGE for 昂貴
  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("昂貴")
    .setFontColor("#e69138")  // 橘字
    .setRanges([rangeAC])
    .build());
    
  // 5. GREY for 無數據
  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains("無數據")
    .setBackground("#efefef") // 灰底
    .setFontColor("#999999")  // 灰字
    .setRanges([rangeAC])
    .build());

  // [NEW] 針對 AB 欄 (PE 本益比) 加入條件式格式化
  const rulesPE = [];
  
  // 1. PE > 25 (昂貴) -> 橘色
  rulesPE.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(25)
    .setFontColor("#e69138")
    .setBold(true)
    .setRanges([rangeAB])
    .build());
    
  // 2. 0 < PE < 15 (便宜) -> 綠色
  rulesPE.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=AND(AB3>0, AB3<15)") 
    .setFontColor("#1e8e3e")
    .setBold(true)
    .setRanges([rangeAB])
    .build());
    
  // 3. PE = 0 (無數據) -> 灰色
  rulesPE.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(0)
    .setFontColor("#999999")
    .setRanges([rangeAB])
    .build());
    
  // [NEW] 針對 Z 欄 (總損益) 加入條件式格式化
  const rulesZ = [];
  rulesZ.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setFontColor("#ff0000") // 正值紅字
    .setBold(true)
    .setRanges([rangeZ])
    .build());
    
  rulesZ.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setFontColor("#1e8e3e") // 負值綠字
    .setBold(true)
    .setRanges([rangeZ])
    .build());

  const allRules = filteredRules.concat(newRules).concat(rulesPE).concat(rulesZ);
  sheet.setConditionalFormatRules(allRules);
}

/**
 * [純搬移] 從買賣同步至在庫
 */
function updateActiveFromInput() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tradeSheet = ss.getSheetByName(SHEET_INPUT);
  const invSheet = ss.getSheetByName(SHEET_ACTIVE);
  
  if (!tradeSheet || !invSheet) return;

  ss.toast("正在同步持股紀錄...", "同步中");

  // 1. 清空「在庫」所有動態資料 (包含新擴充的 AA~AE 欄位)
  const invLastRow = invSheet.getLastRow();
  if (invLastRow >= 3) {
    invSheet.getRange(3, 1, invLastRow - 2, 31).clearContent().setBackground(null).setFontColor(null);
  }

  // 2. 從「買賣」抓取
  const tradeLastRow = tradeSheet.getLastRow();
  if (tradeLastRow < 2) return;
  
  const tradeData = tradeSheet.getRange(2, 2, tradeLastRow - 1, 6).getValues(); 
  let activeTrades = [];

  for (let i = 0; i < tradeData.length; i++) {
    const isSold = tradeData[i][0]; 
    const stockId = tradeData[i][2]; 
    
    if (isSold === false && stockId !== "") {
      activeTrades.push([
        tradeData[i][1], // A
        tradeData[i][2], // B
        tradeData[i][3], // C
        tradeData[i][4], // D
        "",              // E
        "",              // F
        tradeData[i][5]  // G
      ]);
    }
  }

  if (activeTrades.length === 0) {
    ss.toast("沒有未賣出的持股。");
    return;
  }

  // 3. 排序與寫入
  activeTrades.sort((a, b) => (a[1] < b[1] ? -1 : 1));
  invSheet.getRange(3, 1, activeTrades.length, 7).setValues(activeTrades);

  // 4. 公式
  for (let j = 0; j < activeTrades.length; j++) {
    const row = j + 3;
    invSheet.getRange(row, 5).setFormula(`=TODAY()`);
    invSheet.getRange(row, 6).setFormula(`=IFERROR(DATEDIF(D${row}, E${row}, "d"), 0)`);
  }

  // 5. 套用股票分組外框
  applyStockGroupingBorders(invSheet);

  ss.toast("✅ 在庫已整理。請點擊『更新行情』以補齊數據。");
}

/**
 * [NEW] 幫同一支股票加上大框框
 */
function applyStockGroupingBorders(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;

  const lastCol = sheet.getLastColumn();
  const range = sheet.getRange(3, 1, lastRow - 2, lastCol);
  
  // 1. 先清除所有舊外框
  range.setBorder(false, false, false, false, false, false);

  const data = sheet.getRange(3, 2, lastRow - 2, 1).getValues(); // 只讀 B 欄 (代碼)
  
  let startRow = 3;
  let currentId = data[0][0].toString().trim();

  for (let i = 1; i < data.length; i++) {
    const nextId = data[i][0].toString().trim();
    const currentRow = i + 3;

    if (nextId !== currentId) {
      drawBlockBorder(sheet, startRow, currentRow - 1, lastCol);
      startRow = currentRow;
      currentId = nextId;
    }
  }
  
  drawBlockBorder(sheet, startRow, lastRow, lastCol);
}

/**
 * 輔助函式：畫外框
 */
function drawBlockBorder(sheet, startRow, endRow, lastCol) {
  const range = sheet.getRange(startRow, 1, (endRow - startRow) + 1, lastCol);
  range.setBorder(true, true, true, true, null, null, "#444444", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}
