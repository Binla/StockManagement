/**
 * ============================================================
 * 1. Config.gs
 * 全域設定檔
 * 這裡定義的變數，在其他所有 .gs 檔案中都可以直接使用
 * ============================================================
 */

const SHEET_ACTIVE   = "在庫";
const SHEET_CLOSED   = "出倉";
const SHEET_INPUT    = "買賣";
const SHEET_DIVIDEND = "股利"; 
const SHEET_CHART    = "投資分配圖"; 
const FEE_RATE       = 0.000397; 
const TAX_RATE       = 0.003;    

// [NEW] 資金管理與技術指標設定
const TOTAL_CAPITAL  = 3000000; // 總資金 (預設 300 萬)
const RSI_PERIOD     = 14;      // RSI 週期
const RSI_OVERBOUGHT = 80;      // RSI 過熱門檻 (Mean Reversion)
