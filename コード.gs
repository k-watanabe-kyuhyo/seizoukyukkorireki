/**
 * 【ファイル1: Code.gs】
 * 製造計画・入出庫管理 Webアプリのサーバーサイド通信スクリプト
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('製造計画・入出庫管理システム')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * 初期データ取得API (商品マスタ、当月の製造計画、出庫履歴を一括返却)
 */
function apiGetInitialData(year, month) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    year = Number(year) || 2026;
    month = Number(month) || 10;

    // 基準年月（初期マスタ設定月）の取得
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

    // 1. 商品マスタの読み込み
    const masterSheet = ss.getSheetByName('商品マスタ');
    if (!masterSheet) {
      throw new Error('シート「商品マスタ」が見つかりません。先に Setup.gs の setupProductionSystem を実行してください。');
    }

    const masterValues = masterSheet.getDataRange().getValues();
    const products = [];
    for (let i = 1; i < masterValues.length; i++) {
      const row = masterValues[i];
      if (row[0] !== '' && row[0] !== undefined) {
        products.push({
          id: Number(row[0]),
          name: String(row[1]),
          spec: String(row[2] || ''),
          target: Number(row[3]) || 0,
          masterInitStock: Number(row[4]) || 0, // マスタ設定の基準在庫
          initStock: Number(row[4]) || 0,       // 当月月初在庫（後で自動計算）
          safeStock: Number(row[5]) || 0,
          memo: String(row[6] || ''),
          packQty: Number(row[7]) || 1          // H列: 入数
        });
      }
    }

    const daysInMonth = new Date(year, month, 0).getDate();

    // 製造計画シートと出庫実績シートを事前に確実に読み込み（スコープエラー防止）
    const planSheet = ss.getSheetByName('日別製造計画');
    const planValues = (planSheet && planSheet.getLastRow() > 1) ? planSheet.getDataRange().getValues() : [];

    const outSheet = ss.getSheetByName('出庫実績トランザクション');
    const outValues = (outSheet && outSheet.getLastRow() > 1) ? outSheet.getDataRange().getValues() : [];

    // 1-2. 月別月初実在庫シートの読み込み（当月がCSV確定実在庫か判定）
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
      // 当月は実在庫CSVで確定済み：確定実在庫を適用
      products.forEach(p => {
        if (realStockMap[p.id] !== undefined) {
          p.initStock = realStockMap[p.id];
        }
      });
    } else {
      // 2. 未確定の場合：自動引き継ぎ計算（基準日から対象月1日までの全製造・全出庫を集計）
      const targetStartDate = new Date(year, month - 1, 1);
      const baseStartDate = new Date(baseYear, baseMonth - 1, 1);

      const diffMap = {};
      products.forEach(p => { diffMap[p.id] = 0; });

      // 未来月への繰越の場合
      if (targetStartDate.getTime() > baseStartDate.getTime()) {
        for (let i = 1; i < planValues.length; i++) {
          const d = new Date(planValues[i][0]);
          if (!isNaN(d.getTime()) && d >= baseStartDate && d < targetStartDate) {
            const pId = Number(planValues[i][1]);
            if (diffMap[pId] !== undefined) diffMap[pId] += (Number(planValues[i][3]) || 0);
          }
        }
        for (let i = 1; i < outValues.length; i++) {
          const d = new Date(outValues[i][0]);
          if (!isNaN(d.getTime()) && d >= baseStartDate && d < targetStartDate) {
            const pId = Number(outValues[i][1]);
            if (diffMap[pId] !== undefined) diffMap[pId] -= (Number(outValues[i][3]) || 0);
          }
        }
        products.forEach(p => {
          p.initStock = p.masterInitStock + (diffMap[p.id] || 0);
        });
      } else if (targetStartDate.getTime() < baseStartDate.getTime()) {
        // 過去月へ遡る場合
        for (let i = 1; i < planValues.length; i++) {
          const d = new Date(planValues[i][0]);
          if (!isNaN(d.getTime()) && d >= targetStartDate && d < baseStartDate) {
            const pId = Number(planValues[i][1]);
            if (diffMap[pId] !== undefined) diffMap[pId] += (Number(planValues[i][3]) || 0);
          }
        }
        for (let i = 1; i < outValues.length; i++) {
          const d = new Date(outValues[i][0]);
          if (!isNaN(d.getTime()) && d >= targetStartDate && d < baseStartDate) {
            const pId = Number(outValues[i][1]);
            if (diffMap[pId] !== undefined) diffMap[pId] -= (Number(outValues[i][3]) || 0);
          }
        }
        products.forEach(p => {
          p.initStock = p.masterInitStock - (diffMap[p.id] || 0);
        });
      }
    }

    // 当月のマトリクス初期化
    const matrix = {};
    products.forEach(p => {
      matrix[p.id] = {};
      for (let d = 1; d <= daysInMonth; d++) {
        matrix[p.id][d] = { make: 0, out: 0 };
      }
    });

    // 当月の製造計画データをマトリクスへ展開
    for (let i = 1; i < planValues.length; i++) {
      const dateVal = new Date(planValues[i][0]);
      if (!isNaN(dateVal.getTime())) {
        if (dateVal.getFullYear() === year && (dateVal.getMonth() + 1) === month) {
          const pId = Number(planValues[i][1]);
          const day = dateVal.getDate();
          const qty = Number(planValues[i][3]) || 0;
          if (matrix[pId] && matrix[pId][day]) {
            matrix[pId][day].make = qty;
          }
        }
      }
    }

    // 当月の出庫実績データをマトリクスへ展開
    for (let i = 1; i < outValues.length; i++) {
      const dateVal = new Date(outValues[i][0]);
      if (!isNaN(dateVal.getTime())) {
        if (dateVal.getFullYear() === year && (dateVal.getMonth() + 1) === month) {
          const pId = Number(outValues[i][1]);
          const day = dateVal.getDate();
          const qty = Number(outValues[i][3]) || 0;
          if (matrix[pId] && matrix[pId][day]) {
            matrix[pId][day].out += qty;
          }
        }
      }
    }

    return {
      success: true,
      year: year,
      month: month,
      baseYear: baseYear,
      baseMonth: baseMonth,
      isAutoCarried: (year !== baseYear || month !== baseMonth),
      isRealStockConfirmed: isRealStockConfirmed,
      realStockMap: realStockMap,
      daysInMonth: daysInMonth,
      products: products,
      matrix: matrix
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * 月別月初実在庫の一括保存API（月別月初実在庫シートへ上書き保存）
 */
function apiSaveMonthlyStock(records, year, month) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let stockSheet = ss.getSheetByName('月別月初実在庫');
    if (!stockSheet) {
      stockSheet = ss.insertSheet('月別月初実在庫');
      stockSheet.appendRow(['対象年月(YYYY-MM)', '商品ID', '商品名', '確定実在庫(cs)', '登録日時']);
      stockSheet.getRange(1, 1, 1, 5)
        .setBackground('#0e7490')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
    }

    year = Number(year);
    month = Number(month);
    const ymKey = Utilities.formatString('%04d-%02d', year, month);

    // 既存の当月実在庫データを削除（上書き）
    const data = stockSheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      const rowYm = String(data[i][0]).trim();
      if (rowYm === ymKey) {
        stockSheet.deleteRow(i + 1);
      }
    }

    // 新規レコードを一括追記
    const rowsToAdd = [];
    const now = new Date();
    records.forEach(r => {
      rowsToAdd.push([ymKey, r.productId, r.productName, r.stock, now]);
    });

    if (rowsToAdd.length > 0) {
      stockSheet.getRange(stockSheet.getLastRow() + 1, 1, rowsToAdd.length, 5).setValues(rowsToAdd);
    }

    return { success: true, count: rowsToAdd.length };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * 製造計画一括保存API
 */
function apiSaveDailyPlans(records, year, month) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let planSheet = ss.getSheetByName('日別製造計画');
    if (!planSheet) {
      planSheet = ss.insertSheet('日別製造計画');
      planSheet.appendRow(['日付(YYYY-MM-DD)', '商品ID', '商品名', '計画製造数', '更新日時']);
    }

    year = Number(year);
    month = Number(month);

    // 既存の当月計画データを削除
    const data = planSheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      const d = new Date(data[i][0]);
      if (!isNaN(d.getTime()) && d.getFullYear() === year && (d.getMonth() + 1) === month) {
        planSheet.deleteRow(i + 1);
      }
    }

    // 新規計画レコードを一括追記
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

/**
 * 出庫実績・予定の一括保存API（当月分を「出庫実績トランザクション」シートへ上書き保存）
 */
function apiSaveDailyOuts(records, year, month) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let outSheet = ss.getSheetByName('出庫実績トランザクション');
    if (!outSheet) {
      outSheet = ss.insertSheet('出庫実績トランザクション');
      outSheet.appendRow(['出庫年月日(YYYY-MM-DD)', '商品ID', '商品名', '出庫数(cs)', '区分(実績/予定)', '登録日時']);
    }

    year = Number(year);
    month = Number(month);

    // 既存の当月出庫データを削除（再取り込み時の重複防止）
    const data = outSheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      const d = new Date(data[i][0]);
      if (!isNaN(d.getTime()) && d.getFullYear() === year && (d.getMonth() + 1) === month) {
        outSheet.deleteRow(i + 1);
      }
    }

    // 新規出庫レコードを一括追記
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