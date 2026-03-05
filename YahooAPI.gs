/**
 * ============================================================
 * 5. YahooAPI.gs
 * 負責所有對外的 Yahoo Finance API 呼叫
 * ============================================================
 */

function testTWSE() {
  try {
    const res = UrlFetchApp.fetch("https://openapi.twse.com.tw/v1/exchangeReport/BWIBBU_ALL", { muteHttpExceptions: true });
    const text = res.getContentText();
    console.log(`[Diagnostic TWSE] Text length: ${text.length}`);
    console.log(`[Diagnostic TWSE] Is Array? ${text.trim().startsWith('[')}`);
    console.log(`[Diagnostic TWSE] 2330 Snippet: ${text.substring(text.indexOf('2330') - 30, text.indexOf('2330') + 150)}`);
  } catch (e) {
    console.warn(`[Diagnostic TWSE] Error: ${e.message}`);
  }
}

function getYahooCompleteData(id) {
  try {
    // 抓取 60 天數據 (使用 query2 增加穩定性)
    const url = `https://query2.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(id)}?interval=1d&range=60d`;
    const options = {
      muteHttpExceptions: true,
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
      }
    };
    const res = JSON.parse(UrlFetchApp.fetch(url, options).getContentText());
    if (!res.chart || !res.chart.result) {
      console.warn(`getYahooCompleteData: No result for ${id}`);
      return null;
    }
    
    const result = res.chart.result[0];
    const indicators = result.indicators.quote[0];
    const highs = indicators.high || [];
    const lows = indicators.low || [];
    const closes = indicators.close || [];

    const validCloses = closes.filter(c => c !== null);
    if (validCloses.length === 0) return null;
    
    // 現價 (cp)
    const cp = result.meta.regularMarketPrice || validCloses[validCloses.length - 1];
    
    // 昨收 (pc): 倒數第二筆，若只有一筆則用現價
    const pc = validCloses.length > 1 ? validCloses[validCloses.length - 2] : cp; 

    // 海龜指標：過去 20 日最高價
    const last20Highs = highs.slice(-21, -1).filter(h => h !== null);
    const high20 = last20Highs.length > 0 ? Math.max(...last20Highs) : cp;

    // 海龜指標：過去 10 日 最低價 (做為停利/停損點)
    const last10Lows = lows.slice(-11, -1).filter(l => l !== null);
    const low10 = last10Lows.length > 0 ? Math.min(...last10Lows) : (cp * 0.95);
    
    // ATR (N 值)
    let sumTR = 0;
    const nPeriod = 20;
    const startIdx = Math.max(1, highs.length - nPeriod);
    let trCount = 0;
    for (let i = startIdx; i < highs.length; i++) {
      if (highs[i] == null || closes[i-1] == null || lows[i] == null) continue;
      let tr = Math.max(
        highs[i] - lows[i], 
        Math.abs(highs[i] - closes[i-1]), 
        Math.abs(lows[i] - closes[i-1])
      );
      sumTR += tr;
      trCount++;
    }

    // RSI 計算 (14日)
    let rsi = 50; // 預設中值
    if (validCloses.length > RSI_PERIOD) {
       let gains = 0;
       let losses = 0;
       const rsiData = validCloses.slice(-(RSI_PERIOD + 1)); 
       
       for (let i = 1; i < rsiData.length; i++) {
         const diff = rsiData[i] - rsiData[i-1];
         if (diff > 0) gains += diff;
         else losses -= diff;
       }
       
       const avgGain = gains / RSI_PERIOD;
       const avgLoss = losses / RSI_PERIOD;
       
       if (avgLoss === 0) {
         rsi = 100;
       } else {
         const rs = avgGain / avgLoss;
         rsi = 100 - (100 / (1 + rs));
       }
    }

    return {
      price: cp,
      prevClose: pc,
      high: high20,
      low10: low10,
      nValue: trCount > 0 ? sumTR / trCount : (cp * 0.015),
      rsi: rsi
    };
  } catch (e) { 
    console.error(`getYahooCompleteData Error (${id}): ${e.message}`);
    return null; 
  }
}

/**
 * [NEW] 批次取得 Yahoo Finance 報價與基本面 (EPS, PE)
 * 支援傳入陣列，例如 ["2330.TW", "8069.TWO"]
 */
function getYahooBatchQuotes(ids) {
  if (!ids || ids.length === 0) return {};
  
  try {
    const symbols = ids.join(",");
    // 使用 query2 並加入 User-Agent 模擬瀏覽器，提升穩定度 (query1 較容易擋 GAS)
    const url = `https://query2.finance.yahoo.com/v7/finance/quote?symbols=${encodeURIComponent(symbols)}`;
    console.log("Yahoo Batch URL (query2): " + url);
    
    const options = {
      muteHttpExceptions: true,
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
      }
    };
    
    const responseText = UrlFetchApp.fetch(url, options).getContentText();
    const res = JSON.parse(responseText);
    
    if (!res.quoteResponse || !res.quoteResponse.result || res.quoteResponse.result.length === 0) {
      console.log("No quote response or empty result. Symbols: " + symbols);
      console.log("Raw Response Snippet: " + responseText.substring(0, 500));
      return {};
    }
    
    const results = res.quoteResponse.result;
    const batchData = {};
    
    results.forEach(quote => {
      const sym = (quote.symbol || "").toUpperCase();
      console.log(`Debug: Symbol=${sym}, EPS=${quote.epsTrailingTwelveMonths}, PE=${quote.trailingPE}`);
      batchData[sym] = {
        eps: quote.epsTrailingTwelveMonths || 0,
        pe: quote.trailingPE || 0,
        yield: (quote.trailingAnnualDividendYield || 0) * 100 // 轉成百分比
      };
    });
    
    return batchData;
  } catch (e) {
    console.error("Batch Quot Fetch Error: " + e.message);
    return {};
  }
}

/**
 * 取得股票基本面資料 (EPS, PE, Yield)
 */
/**
 * 取得股票基本面資料 (EPS, PE, Yield)
 * 備註：由於 Yahoo Finance API 的 quote 端點可能會擋 GAS IP (401 Unauthorized)，
 * 這裡改用直接抓取網頁 HTML 解析的方式作為備案。
 */
// 取得台灣證交所與櫃買中心的官方每日 PE 資料 (快取 6 小時)
function getTwsePeData() {
  const cache = CacheService.getScriptCache();
  if (!cache) return {}; // 如果不在 Apps Script 環境中
  
  const cached = cache.get("TWSE_PE_DATA_V2");
  if (cached) return JSON.parse(cached);

  const dict = {};
  
  const fetchOptions = {
    muteHttpExceptions: true,
    headers: {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
  };

  // 上市 (TWSE)
  try {
    const res = UrlFetchApp.fetch("https://openapi.twse.com.tw/v1/exchangeReport/BWIBBU_ALL", fetchOptions);
    const json = JSON.parse(res.getContentText());
    if (Array.isArray(json)) {
       json.forEach(item => { 
         const code = item.Code || item.stockCode || item.SecuritiesCompanyCode;
         const pe = parseFloat(item.PeRatio || item.PERatio || item.PEratio) || 0;
         if (code) dict[code.toString()] = pe; 
       });
    }
  } catch(e) { console.warn("TWSE Cache Err:", e.message); }

  // 上櫃 (TPEx)
  try {
    const tpexOptions = {
       muteHttpExceptions: true
       // 不帶特別的 User-Agent，讓它預設以 Google Apps Script 身分連線，避免觸發 Cloudflare
    };
    const res2 = UrlFetchApp.fetch("https://www.tpex.org.tw/openapi/v1/t187ap14_L", tpexOptions);
    const text2 = res2.getContentText();
    if (text2.trim().startsWith("[")) {
      const json2 = JSON.parse(text2);
      if (Array.isArray(json2)) {
         json2.forEach(item => { 
           const code = item.Code || item.stockCode || item.SecuritiesCompanyCode;
           const pe = parseFloat(item.PeRatio || item.PERatio || item.PEratio) || 0;
           if (code) dict[code.toString()] = pe; 
         });
      }
    }
  } catch(e) { console.warn("TPEx Cache Err:", e.message); }

  if (Object.keys(dict).length > 0) {
      cache.put("TWSE_PE_DATA_V2", JSON.stringify(dict), 21600); // 快取 6 小時 (Google Apps Script 上限)
  }
  return dict;
}

/**
 * 取得基本分析資料 (EPS、本益比、殖利率)
 * @param {string} id 股票代號 (如 2330.TW, AAPL)
 * @param {number} currentPrice 選擇性：用於官方 API 反推 EPS
 * @return {object} {eps: 數值, pe: 數值, yield: 數值}
 */
function getYahooQuote(id, currentPrice = 0) {
  if (!id) return { eps: 0, pe: 0, yield: 0 };
  
  let pureId = id.toString();
  const isTaiwan = pureId.includes(".TW") || pureId.includes(".TWO") || pureId.includes(".tw");
  
  if (isTaiwan) {
    pureId = pureId.replace(".TW", "").replace(".TWO", "").replace(".tw", "");
    
    // 台股唯一解霸：政府官方 Open API 資料集 (無痛無阻擋)
    const peDict = getTwsePeData();
    const pe = peDict[pureId] || 0;
    
    if (pe !== 0) {
      let cp = currentPrice;
      
      // 如果沒有傳入股價，自己去抓 (確保 testSingleStock 也能算出 EPS)
      if (cp <= 0) {
         try {
           const url = `https://query2.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(id)}?interval=1d&range=1d`;
           const res = JSON.parse(UrlFetchApp.fetch(url, {
             muteHttpExceptions: true,
             headers: { "User-Agent": "Mozilla/5.0" }
           }).getContentText());
           if (res.chart && res.chart.result) {
              cp = res.chart.result[0].meta.regularMarketPrice || 0;
           }
         } catch(e) {}
      }

      // 透過 PE 與當前股價，用數學反推 EPS (這是最準的 Trailing EPS)
      let eps = 0;
      if (cp > 0) {
         eps = cp / pe;
         eps = Math.round(eps * 100) / 100;
      }
      console.log(`[Diagnostic] TW OpenData Success for ${id}: PE=${pe}, Calc EPS=${eps}`);
      return { pe: pe, eps: eps, yield: 0 };
    } else {
      console.log(`[Diagnostic] Not found in TW OpenData for ${id}.`);
    }
  }

  // 備用方案 2：全球版 Yahoo Finance HTML 硬爬蟲 (如果奇摩股市也失敗)
  
  // 美股或其他 fallback (使用原本的 Yahoo query1 作為底線)
  try {
    const symbol = encodeURIComponent(id);
    const url = `https://query1.finance.yahoo.com/v7/finance/quote?symbols=${symbol}`;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const res = JSON.parse(response.getContentText());
    
    if (res.quoteResponse && res.quoteResponse.result && res.quoteResponse.result.length > 0) {
      const q = res.quoteResponse.result[0];
      return {
        eps: q.epsTrailingTwelveMonths || q.epsForward || q.epsCurrentYear || 0,
        pe: q.trailingPE || q.forwardPE || 0,
        yield: (q.trailingAnnualDividendYield || q.dividendYield || 0) * 100
      };
    }
  } catch (e) { }
  
  return { eps: 0, pe: 0, yield: 0 };
}

function updateMarketIndex(sheet) {
  if (!sheet) {
    try {
      sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    } catch (e) {
      console.error("無法取得 Sheet");
      return;
    }
  }

  try {
    const timestamp = new Date().getTime();
    
    // ---------------------------------------------------------
    // 1. 取得即時報價 (1m)
    // ---------------------------------------------------------
    const urlLive = `https://query1.finance.yahoo.com/v8/finance/chart/%5ETWII?interval=1m&range=1d&_=${timestamp}`;
    const resLive = JSON.parse(UrlFetchApp.fetch(urlLive, { muteHttpExceptions: true }).getContentText());
    
    let currentPrice = 0;
    
    if (resLive.chart && resLive.chart.result) {
      const meta = resLive.chart.result[0].meta;
      currentPrice = meta.regularMarketPrice;
      const prevClose = meta.chartPreviousClose || meta.previousClose;
      const change = currentPrice - prevClose;
      const changePercent = (change / prevClose);
      const marketTime = new Date(meta.regularMarketTime * 1000);
      const updateTime = Utilities.formatDate(marketTime, "GMT+8", "yyyy/MM/dd HH:mm:ss");

      // 顯示基本行情
      sheet.getRange("B1").setValue("大盤");
      sheet.getRange("C1").setValue(currentPrice).setNumberFormat("#,##0.00").setFontWeight("bold");
      sheet.getRange("D1").setValue(change).setNumberFormat("#,##0.00");
      sheet.getRange("E1").setValue(changePercent).setNumberFormat("0.00%");
      sheet.getRange("F1").setValue("前收：");
      sheet.getRange("G1").setValue(prevClose).setNumberFormat("#,##0.00");
      
      let color = (change > 0) ? "#ff0000" : (change < 0 ? "#008000" : "#000000");
      sheet.getRange("C1:E1").setFontColor(color);
      sheet.getRange("M1").setValue("更新: " + updateTime);
    }

    // ---------------------------------------------------------
    // 2. 取得歷史資料計算 60MA (季線)
    // ---------------------------------------------------------
    // 抓取 90 天日線，確保有足夠的交易日 (假日無資料)
    const urlHist = `https://query1.finance.yahoo.com/v8/finance/chart/%5ETWII?interval=1d&range=3mo&_=${timestamp}`;
    const resHist = JSON.parse(UrlFetchApp.fetch(urlHist, { muteHttpExceptions: true }).getContentText());

    if (resHist.chart && resHist.chart.result) {
      const closes = resHist.chart.result[0].indicators.quote[0].close;
      // 過濾掉 null 值
      const validCloses = closes.filter(c => c !== null);
      
      // 取最後 60 筆 (如果不足 60 筆則取全部)
      const maPeriod = 60;
      const dataForMa = validCloses.slice(-maPeriod);
      
      if (dataForMa.length > 0) {
        const sum = dataForMa.reduce((a, b) => a + b, 0);
        const ma60 = sum / dataForMa.length;
        
        // 顯示 60MA
        sheet.getRange("H1").setValue("季線(60MA):");
        sheet.getRange("I1").setValue(ma60).setNumberFormat("#,##0.00");
        
        // 判斷多空 (紅底多頭，綠底空頭)
        const rangeStatus = sheet.getRange("J1");
        if (currentPrice >= ma60) {
          rangeStatus.setValue("📈 多頭趨勢").setBackground("#f4cccc").setFontColor("#cc0000").setFontWeight("bold"); // 紅底紅字
        } else {
          rangeStatus.setValue("📉 空頭警戒").setBackground("#d9ead3").setFontColor("#1e8e3e").setFontWeight("bold"); // 綠底綠字
        }
      }
    }
    
    console.log(`大盤更新成功: ${currentPrice}`);

  } catch (e) {
    console.error("大盤更新失敗: " + e.message);
  }
}

/**
 * 診斷工具：測試單一股票抓取
 * 請在 Apps Script 編輯器中選擇此函式並點擊「執行」
 */
function testSingleStock() {
  const testId = "2330.TW"; // 您可以更換成其他代碼測試，例如 "0050.TW"
  console.log(`--- 開始測試 ${testId} ---`);
  
  const data = getYahooCompleteData(testId);
  if (data) {
    console.log("✅ 行情抓取成功:");
    console.log(`   現價: ${data.price}`);
    console.log(`   昨收: ${data.prevClose}`);
    console.log(`   海龜 N值: ${data.nValue}`);
    console.log(`   RSI: ${data.rsi}`);
  } else {
    console.error("❌ 行情抓取失敗，請確認網路或代碼是否正確。");
  }
  
  const fundamental = getYahooQuote(testId);
  if (fundamental) {
    console.log("✅ 基本面抓取成功:");
    console.log(`   EPS: ${fundamental.eps}`);
    console.log(`   PE: ${fundamental.pe}`);
  } else {
    console.error("❌ 基本面抓取失敗。");
  }
  console.log("--- 測試結束 ---");
}
