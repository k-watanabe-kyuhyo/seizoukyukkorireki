/**
 * 【Code.gs】製造計画・入出庫・年間計画統合システム サーバーサイド
 * 貼り付け先: GASの「コード.gs」
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('製造計画・入出庫・年間計画統合システム')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * 年月セルの値を安全にパースするヘルパー関数
 * （Dateオブジェクト、文字列 "2026-01", "2026/01" などすべてに対応）
 */
function parseYm(val) {
  if (!val) return null;
  if (val instanceof Date) {
    return {
      year: val.getFullYear(),
      month: val.getMonth() + 1
    };
  }
  const str = String(val).trim().replace(/^['"]/, '');
  const m = str.match(/(\d{4})[-/年](\d{1,2})/);
  if (m) {
    return {
      year: Number(m[1]),
      month: Number(m[2])
    };
  }
  return null;
}

/**
 * 初期データ一括取得API (商品マスタ、当月計画、出庫履歴、確定実在庫、年間計画)
 */
function apiGetInitialData(year, month) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    year = Number(year) || 2026;
    month = Number(month) || 10;

    let baseYear = 2026;
    let baseMonth = 10;
    const configSheet = ss.getSheetByName('システム設定');
    if (configSheet && configSheet.getLastRow() > 1) {
      const cfgValues = configSheet.getDataRange().getValues();
      cfgValues.forEach(row => {
        if (row[0] === '運用対象年' && row[1]) baseYear = Number(row[1]);
        if (row[0] === '運用対象月' && row[1]) baseMonth = Number(row[1]);
      });
    }

    // 商品マスタの読み込み（万が一シートが空でも画面が壊れないフォールバック付き）
    const masterSheet = ss.getSheetByName('商品マスタ');
    const products = [];
    if (masterSheet && masterSheet.getLastRow() > 1) {
      const masterValues = masterSheet.getDataRange().getValues();
      for (let i = 1; i < masterValues.length; i++) {
        const row = masterValues[i];
        if (row[0] !== '' && row[0] !== undefined) {
          products.push({
            id: Number(row[0]),
            code: Number(row[0]),
            name: String(row[1]),
            spec: String(row[2] || ''),
            target: Number(row[3]) || 0,
            masterInitStock: Number(row[4]) || 0,
            initStock: Number(row[4]) || 0,
            safeStock: Number(row[5]) || 0,
            memo: String(row[6] || ''),
            packQty: Number(row[7]) || 1
          });
        }
      }
    }

    // デフォルト14品目のフォールバック
    if (products.length === 0) {
      const defaultProducts = [
        { id: 1,  code: 112, name: 'おいしい氷3kg',         spec: '3kg×4',   packQty: 4,  target: 7200, masterInitStock: 6208, initStock: 6208, safeStock: 500 },
        { id: 2,  code: 116, name: 'おいしい氷アイスカップ', spec: 'カップ',  packQty: 12, target: 60,   masterInitStock: 50,   initStock: 50,   safeStock: 20 },
        { id: 3,  code: 23,  name: '博多純氷カチ割り1kg',   spec: '1kg×10',  packQty: 10, target: 1400, masterInitStock: 706,  initStock: 706,  safeStock: 200 },
        { id: 4,  code: 31,  name: 'おいしい氷1kg',         spec: '1kg×10',  packQty: 12, target: 1200, masterInitStock: 2015, initStock: 2015, safeStock: 300 },
        { id: 5,  code: 163, name: 'ドンキホーテ極氷1.1kg', spec: '1.1kg×8', packQty: 10, target: 1400, masterInitStock: 2336, initStock: 2336, safeStock: 200 },
        { id: 6,  code: 37,  name: 'ヒロキュー千年氷1kg',   spec: '1kg×10',  packQty: 10, target: 500,  masterInitStock: 200,  initStock: 200,  safeStock: 100 },
        { id: 7,  code: 161, name: 'CGC1kg',                spec: '1kg×10',  packQty: 12, target: 1200, masterInitStock: 11,   initStock: 11,   safeStock: 300 },
        { id: 8,  code: 52,  name: 'おいしい氷2kgチャック', spec: '2kg×5',   packQty: 6,  target: 1600, masterInitStock: 1761, initStock: 1761, safeStock: 250 },
        { id: 9,  code: 74,  name: 'おいしい氷1.1kg',       spec: '1.1kg×8', packQty: 12, target: 3600, masterInitStock: 5975, initStock: 5975, safeStock: 600 },
        { id: 10, code: 159, name: 'ゆめイチ1kg',           spec: '1kg×10',  packQty: 12, target: 2400, masterInitStock: 1636, initStock: 1636, safeStock: 400 },
        { id: 11, code: 160, name: 'ゆめイチ2.1kg',         spec: '2.1kg×4', packQty: 6,  target: 4000, masterInitStock: 5829, initStock: 5829, safeStock: 600 },
        { id: 12, code: 169, name: 'スティック40mm',        spec: '本',      packQty: 1,  target: 200,  masterInitStock: 69,   initStock: 69,   safeStock: 30 },
        { id: 13, code: 170, name: 'スティック60mm',        spec: '本',      packQty: 1,  target: 50,   masterInitStock: 30,   initStock: 30,   safeStock: 10 },
        { id: 14, code: 171, name: 'スティック110mm',       spec: '本',      packQty: 1,  target: 50,   masterInitStock: 40,   initStock: 40,   safeStock: 10 }
      ];
      defaultProducts.forEach(p => products.push(p));
    }

    const daysInMonth = new Date(year, month, 0).getDate();
    const planSheet = ss.getSheetByName('日別製造計画');
    const planValues = (planSheet && planSheet.getLastRow() > 1) ? planSheet.getDataRange().getValues() : [];

    const outSheet = ss.getSheetByName('出庫実績トランザクション');
    const outValues = (outSheet && outSheet.getLastRow() > 1) ? outSheet.getDataRange().getValues() : [];

    const ymKey = Utilities.formatString('%04d-%02d', year, month);
    let isRealStockConfirmed = false;
    const realStockMap = {};

    const stockSheet = ss.getSheetByName('月別月初実在庫');
    if (stockSheet && stockSheet.getLastRow() > 1) {
      const stockValues = stockSheet.getDataRange().getValues();
      for (let i = 1; i < stockValues.length; i++) {
        const rowYm = String(stockValues[i][0]).trim();
        if (rowYm === ymKey) {
          const pId = Number(stockValues[i][1]);
          const sVal = Number(stockValues[i][3]) || 0;
          realStockMap[pId] = sVal;
          isRealStockConfirmed = true;
        }
      }
    }

    if (isRealStockConfirmed) {
      products.forEach(p => {
        if (realStockMap[p.id] !== undefined) p.initStock = realStockMap[p.id];
      });
    }

    const matrix = {};
    products.forEach(p => {
      matrix[p.id] = {};
      for (let d = 1; d <= daysInMonth; d++) {
        matrix[p.id][d] = { make: 0, out: 0 };
      }
    });

    for (let i = 1; i < planValues.length; i++) {
      const dateVal = new Date(planValues[i][0]);
      if (!isNaN(dateVal.getTime()) && dateVal.getFullYear() === year && (dateVal.getMonth() + 1) === month) {
        const pId = Number(planValues[i][1]);
        const day = dateVal.getDate();
        const qty = Number(planValues[i][3]) || 0;
        if (matrix[pId] && matrix[pId][day]) matrix[pId][day].make = qty;
      }
    }

    for (let i = 1; i < outValues.length; i++) {
      const dateVal = new Date(outValues[i][0]);
      if (!isNaN(dateVal.getTime()) && dateVal.getFullYear() === year && (dateVal.getMonth() + 1) === month) {
        const pId = Number(outValues[i][1]);
        const day = dateVal.getDate();
        const qty = Number(outValues[i][3]) || 0;
        if (matrix[pId] && matrix[pId][day]) matrix[pId][day].out += qty;
      }
    }

    // --- 年間計画データの読み込み（全年度対応・日付型自動認識） ---
    const annualSheet = ss.getSheetByName('年間計画');
    const annualMatrixByYear = {};
    const annualInitStocksByYear = {}; // 年間計画専用の前月末在庫マップ

    if (annualSheet && annualSheet.getLastRow() > 1) {
      const annualValues = annualSheet.getDataRange().getValues();
      for (let i = 1; i < annualValues.length; i++) {
        const parsed = parseYm(annualValues[i][0]);
        if (parsed) {
          const y = parsed.year;
          const m = parsed.month;
          const pId = Number(annualValues[i][1]);
          const sales = Number(annualValues[i][3]) || 0;
          const make = Number(annualValues[i][4]) || 0;
          let customStock = null;
          if (annualValues[i][8] !== undefined && annualValues[i][8] !== '') {
            customStock = Number(annualValues[i][8]);
          }

          // 前月末在庫（列7または列8）の読み取り
          let initStk = undefined;
          if (annualValues[i][6] !== undefined && typeof annualValues[i][6] === 'number') {
            initStk = Number(annualValues[i][6]);
          } else if (annualValues[i][7] !== undefined && typeof annualValues[i][7] === 'number') {
            initStk = Number(annualValues[i][7]);
          } else if (!isNaN(Number(annualValues[i][6])) && annualValues[i][6] !== '') {
            if (!(annualValues[i][6] instanceof Date)) {
              initStk = Number(annualValues[i][6]);
            }
          }

          if (!annualMatrixByYear[y]) annualMatrixByYear[y] = {};
          if (!annualInitStocksByYear[y]) annualInitStocksByYear[y] = {};

          if (!annualMatrixByYear[y][pId]) {
            annualMatrixByYear[y][pId] = {};
            for (let k = 1; k <= 12; k++) annualMatrixByYear[y][pId][k] = { sales: 0, make: 0, customStock: null };
          }
          if (m >= 1 && m <= 12) {
            annualMatrixByYear[y][pId][m] = { sales: sales, make: make, customStock: customStock };
          }
          if (initStk !== undefined && (m === 1 || annualInitStocksByYear[y][pId] === undefined)) {
            annualInitStocksByYear[y][pId] = initStk;
          }
        }
      }
    }

    return {
      success: true,
      year: year,
      month: month,
      isRealStockConfirmed: isRealStockConfirmed,
      realStockMap: realStockMap,
      daysInMonth: daysInMonth,
      products: products,
      matrix: matrix,
      annualMatrix: annualMatrixByYear[year] || {},
      annualMatrixByYear: annualMatrixByYear,
      annualInitStocksByYear: annualInitStocksByYear
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/** 月別実在庫の一括保存 */
function apiSaveMonthlyStock(records, year, month) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let stockSheet = ss.getSheetByName('月別月初実在庫') || ss.insertSheet('月別月初実在庫');
    year = Number(year);
    month = Number(month);
    const ymKey = Utilities.formatString('%04d-%02d', year, month);

    const data = stockSheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]).trim() === ymKey) stockSheet.deleteRow(i + 1);
    }

    const rowsToAdd = [];
    const now = new Date();
    records.forEach(r => rowsToAdd.push([ymKey, r.productId, r.productName, r.stock, now]));
    if (rowsToAdd.length > 0) {
      stockSheet.getRange(stockSheet.getLastRow() + 1, 1, rowsToAdd.length, 5).setValues(rowsToAdd);
    }
    return { success: true, count: rowsToAdd.length };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/** 製造計画の一括保存 */
function apiSaveDailyPlans(records, year, month) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let planSheet = ss.getSheetByName('日別製造計画') || ss.insertSheet('日別製造計画');
    year = Number(year);
    month = Number(month);

    const data = planSheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      const d = new Date(data[i][0]);
      if (!isNaN(d.getTime()) && d.getFullYear() === year && (d.getMonth() + 1) === month) {
        planSheet.deleteRow(i + 1);
      }
    }

    const rowsToAdd = [];
    const now = new Date();
    records.forEach(r => {
      if (r.make > 0) {
        const dStr = Utilities.formatDate(new Date(year, month - 1, r.day), 'JST', 'yyyy-MM-dd');
        rowsToAdd.push([dStr, r.productId, r.productName, r.make, now]);
      }
    });

    if (rowsToAdd.length > 0) {
      planSheet.getRange(planSheet.getLastRow() + 1, 1, rowsToAdd.length, 5).setValues(rowsToAdd);
    }
    return { success: true, count: rowsToAdd.length };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/** 出庫実績の一括保存 */
function apiSaveDailyOuts(records, year, month) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let outSheet = ss.getSheetByName('出庫実績トランザクション') || ss.insertSheet('出庫実績トランザクション');
    year = Number(year);
    month = Number(month);

    const data = outSheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      const d = new Date(data[i][0]);
      if (!isNaN(d.getTime()) && d.getFullYear() === year && (d.getMonth() + 1) === month) {
        outSheet.deleteRow(i + 1);
      }
    }

    const rowsToAdd = [];
    const now = new Date();
    records.forEach(r => {
      if (r.out > 0) {
        const dStr = Utilities.formatDate(new Date(year, month - 1, r.day), 'JST', 'yyyy-MM-dd');
        rowsToAdd.push([dStr, r.productId, r.productName, r.out, '出庫CSV/予定', now]);
      }
    });

    if (rowsToAdd.length > 0) {
      outSheet.getRange(outSheet.getLastRow() + 1, 1, rowsToAdd.length, 6).setValues(rowsToAdd);
    }
    return { success: true, count: rowsToAdd.length };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/** 年間計画の一括保存 */
function apiSaveAnnualPlan(records, year) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let annualSheet = ss.getSheetByName('年間計画') || ss.insertSheet('年間計画');
    year = Number(year);

    // ヘッダーが未設定の場合は初期化
    if (annualSheet.getLastRow() === 0) {
      annualSheet.appendRow(['年月', '商品ID', '商品名', '販売数', '計画製造数', '月末在庫', '前月末在庫', '更新日時', '手入力実在庫']);
      annualSheet.getRange(1, 1, 1, 9)
        .setBackground('#1e1b4b')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
    }

    const data = annualSheet.getDataRange().getValues();
    // 該当年の古い行を確実に判定して削除（上書き保存）
    for (let i = data.length - 1; i >= 1; i--) {
      const parsed = parseYm(data[i][0]);
      if (parsed && parsed.year === year) {
        annualSheet.deleteRow(i + 1);
      }
    }

    const rowsToAdd = [];
    const now = new Date();
    records.forEach(r => {
      rowsToAdd.push([
        r.ymKey,
        r.productId,
        r.productName,
        r.sales,
        r.make,
        r.stock,
        (r.annualInitStock !== undefined ? r.annualInitStock : 0),
        now,
        (r.customStock !== undefined && r.customStock !== null && r.customStock !== '' ? r.customStock : '')
      ]);
    });
    
    if (rowsToAdd.length > 0) {
      const startRow = annualSheet.getLastRow() + 1;
      const targetRange = annualSheet.getRange(startRow, 1, rowsToAdd.length, 9);
      targetRange.setNumberFormat('@'); // テキスト書式を強制して自動変換を防止
      targetRange.setValues(rowsToAdd);
    }
    return { success: true, count: rowsToAdd.length };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}