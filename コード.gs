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

    // デフォルトコードのフォールバック辞書（万が一スプシのI列が空でも安全に稼働）
    const defaultCodes = {
      1: 112, 2: 116, 3: 23, 4: 31, 5: 163, 6: 37, 7: 161,
      8: 52, 9: 74, 10: 159, 11: 160, 12: 169, 13: 170, 14: 171
    };

    // 商品マスタの読み込み（I列: 商品コードを反映）
    const masterSheet = ss.getSheetByName('商品マスタ');
    const products = [];
    if (masterSheet && masterSheet.getLastRow() > 1) {
      const masterValues = masterSheet.getDataRange().getValues();
      for (let i = 1; i < masterValues.length; i++) {
        const row = masterValues[i];
        if (row[0] !== '' && row[0] !== undefined) {
          const pId = Number(row[0]);
          // I列（インデックス8）の商品コードを取得。空ならフォールバック辞書から取得
          let pCode = Number(row[8]);
          if (!pCode || isNaN(pCode)) {
            pCode = defaultCodes[pId] || pId;
          }

          products.push({
            id: pId,
            code: pCode, // ★ スプレッドシートI列の商品コードを正しく連携
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

    // --- 年間計画データの読み込み ---
    const annualSheet = ss.getSheetByName('年間計画');
    const annualMatrixByYear = {};
    const annualInitStocksByYear = {};

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

          let initStk = undefined;
          if (annualValues[i][6] !== undefined && typeof annualValues[i][6] === 'number') {
            initStk = Number(annualValues[i][6]);
          } else if (annualValues[i][7] !== undefined && typeof annualValues[i][7] === 'number') {
            initStk = Number(annualValues[i][7]);
          }

          if (!annualMatrixByYear[y]) annualMatrixByYear[y] = {};
          if (!annualMatrixByYear[y][pId]) annualMatrixByYear[y][pId] = {};
          annualMatrixByYear[y][pId][m] = { sales: sales, make: make, customStock: customStock };

          if (initStk !== undefined) {
            if (!annualInitStocksByYear[y]) annualInitStocksByYear[y] = {};
            annualInitStocksByYear[y][pId] = initStk;
          }
        }
      }
    }

    return {
      success: true,
      year: year,
      month: month,
      daysInMonth: daysInMonth,
      products: products,
      matrix: matrix,
      realStockMap: realStockMap,
      annualMatrixByYear: annualMatrixByYear,
      annualInitStocksByYear: annualInitStocksByYear
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * 日別計画保存API
 */
function apiSaveDailyPlans(records, year, month) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let planSheet = ss.getSheetByName('日別製造計画');
    if (!planSheet) {
      planSheet = ss.insertSheet('日別製造計画');
      planSheet.appendRow(['年月日(YYYY-MM-DD)', '商品ID', '商品名', '計画製造数(cs)', '最終更新日時']);
    }

    year = Number(year) || 2026;
    month = Number(month) || 10;
    const nowStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    const existing = planSheet.getDataRange().getValues();
    const rowsToKeep = [];

    if (existing.length > 0) {
      rowsToKeep.push(existing[0]);
      for (let i = 1; i < existing.length; i++) {
        const d = new Date(existing[i][0]);
        if (isNaN(d.getTime()) || d.getFullYear() !== year || (d.getMonth() + 1) !== month) {
          rowsToKeep.push(existing[i]);
        }
      }
    }

    const newRows = [];
    if (records && records.length > 0) {
      records.forEach(r => {
        const dateStr = Utilities.formatString('%04d-%02d-%02d', year, month, Number(r.day));
        newRows.push([dateStr, Number(r.productId), String(r.productName), Number(r.make), nowStr]);
      });
    }

    planSheet.clearContents();
    const allRows = rowsToKeep.concat(newRows);
    if (allRows.length > 0) {
      planSheet.getRange(1, 1, allRows.length, allRows[0].length).setValues(allRows);
    }
    return { success: true, count: newRows.length };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * 出庫実績保存API
 */
function apiSaveDailyOuts(records, year, month) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let outSheet = ss.getSheetByName('出庫実績トランザクション');
    if (!outSheet) {
      outSheet = ss.insertSheet('出庫実績トランザクション');
      outSheet.appendRow(['出庫年月日(YYYY-MM-DD)', '商品ID', '商品名', '出庫数(cs)', '区分(実績/予定)', '登録日時']);
    }

    year = Number(year) || 2026;
    month = Number(month) || 10;
    const nowStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    const existing = outSheet.getDataRange().getValues();
    const rowsToKeep = [];

    if (existing.length > 0) {
      rowsToKeep.push(existing[0]);
      for (let i = 1; i < existing.length; i++) {
        const d = new Date(existing[i][0]);
        if (isNaN(d.getTime()) || d.getFullYear() !== year || (d.getMonth() + 1) !== month) {
          rowsToKeep.push(existing[i]);
        }
      }
    }

    const newRows = [];
    if (records && records.length > 0) {
      records.forEach(r => {
        const dateStr = Utilities.formatString('%04d-%02d-%02d', year, month, Number(r.day));
        newRows.push([dateStr, Number(r.productId), String(r.productName), Number(r.out), '出庫CSV取込', nowStr]);
      });
    }

    outSheet.clearContents();
    const allRows = rowsToKeep.concat(newRows);
    if (allRows.length > 0) {
      outSheet.getRange(1, 1, allRows.length, allRows[0].length).setValues(allRows);
    }
    return { success: true, count: newRows.length };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * 月初実在庫保存API
 */
function apiSaveMonthlyStock(records, year, month) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let stockSheet = ss.getSheetByName('月別月初実在庫');
    if (!stockSheet) {
      stockSheet = ss.insertSheet('月別月初実在庫');
      stockSheet.appendRow(['対象年月(YYYY-MM)', '商品ID', '商品名', '確定実在庫(cs)', '最終更新日時']);
    }

    year = Number(year) || 2026;
    month = Number(month) || 10;
    const ymKey = Utilities.formatString('%04d-%02d', year, month);
    const nowStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    const existing = stockSheet.getDataRange().getValues();
    const rowsToKeep = [];

    if (existing.length > 0) {
      rowsToKeep.push(existing[0]);
      for (let i = 1; i < existing.length; i++) {
        if (String(existing[i][0]).trim() !== ymKey) {
          rowsToKeep.push(existing[i]);
        }
      }
    }

    const newRows = [];
    if (records && records.length > 0) {
      records.forEach(r => {
        newRows.push([ymKey, Number(r.productId), String(r.productName), Number(r.stock), nowStr]);
      });
    }

    stockSheet.clearContents();
    const allRows = rowsToKeep.concat(newRows);
    if (allRows.length > 0) {
      stockSheet.getRange(1, 1, allRows.length, allRows[0].length).setValues(allRows);
    }
    return { success: true, count: newRows.length };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * 年間計画保存API
 */
function apiSaveAnnualPlan(records, year) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let annualSheet = ss.getSheetByName('年間計画');
    if (!annualSheet) {
      annualSheet = ss.insertSheet('年間計画');
      annualSheet.appendRow(['年月', '商品ID', '商品名', '販売計画(cs)', '計画製造数(cs)', '月末在庫(cs)', '前月末在庫', '備考', '手入力在庫']);
    }

    year = Number(year) || 2026;
    const existing = annualSheet.getDataRange().getValues();
    const rowsToKeep = [];

    if (existing.length > 0) {
      rowsToKeep.push(existing[0]);
      for (let i = 1; i < existing.length; i++) {
        const parsed = parseYm(existing[i][0]);
        if (!parsed || parsed.year !== year) {
          rowsToKeep.push(existing[i]);
        }
      }
    }

    const newRows = [];
    if (records && records.length > 0) {
      records.forEach(r => {
        newRows.push([
          String(r.ymKey),
          Number(r.productId),
          String(r.productName),
          Number(r.sales) || 0,
          Number(r.make) || 0,
          Number(r.stock) || 0,
          (r.annualInitStock !== undefined ? Number(r.annualInitStock) : ''),
          '',
          (r.customStock !== null && r.customStock !== undefined ? Number(r.customStock) : '')
        ]);
      });
    }

    annualSheet.clearContents();
    const allRows = rowsToKeep.concat(newRows);
    if (allRows.length > 0) {
      annualSheet.getRange(1, 1, allRows.length, allRows[0].length).setValues(allRows);
    }
    return { success: true, count: newRows.length };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}