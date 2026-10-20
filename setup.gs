/**
 * 【ファイル2: Setup.gs】
 * 製造計画・入出庫管理 初期セットアップ用スクリプト
 * 
 * 使い方：
 * このコードを貼り付けた後、Apps Script上部の関数選択で「setupProductionSystem」を選び、
 * 「実行」ボタンを1回押してください。必要なシートと商品マスタが自動生成されます。
 */

function setupProductionSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. 商品マスタシートの作成
  let masterSheet = ss.getSheetByName('商品マスタ');
  if (!masterSheet) masterSheet = ss.insertSheet('商品マスタ');
  masterSheet.clear();

  const masterHeaders = ['商品ID', '商品名', '規格・単位', '月間販売目標(cs)', '月初在庫(cs)', '安全在庫(cs)', '備考', '入数'];
  masterSheet.appendRow(masterHeaders);
  masterSheet.getRange(1, 1, 1, masterHeaders.length)
    .setBackground('#1e293b')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  // Excel画像に記載されていた商品マスタと目標・月初在庫・H列入数データ
  const initialProducts = [
    [1, 'おいしい氷3kg', '3kg×4', 7200, 6208, 500, '主力商品', 4],
    [2, 'おいしい氷アイスカップ', 'カップ', 60, 50, 20, '', 12],
    [3, '博多純氷カチ割り1kg', '1kg×10', 1400, 706, 200, '', 10],
    [4, 'おいしい氷1kg', '1kg×10', 1200, 2015, 300, '', 12],
    [5, 'ドンキホーテ極氷1.1kg', '1.1kg×8', 1400, 2336, 200, '', 10],
    [6, 'ヒロキュー千年氷1kg', '1kg×10', 500, 200, 100, '', 10],
    [7, 'CGC1kg', '1kg×10', 1200, 11, 300, '', 12],
    [8, 'おいしい氷2kgチャック', '2kg×5', 1600, 1761, 250, 'チャック付き袋', 6],
    [9, 'おいしい氷1.1kg', '1.1kg×8', 3600, 5975, 600, '', 12],
    [10, 'ゆめイチ1kg', '1kg×10', 2400, 1636, 400, '', 12],
    [11, 'ゆめイチ2.1kg', '2.1kg×4', 4000, 5829, 600, '', 6],
    [12, 'スティック40mm', '本', 200, 69, 30, '', 1],
    [13, 'スティック60mm', '本', 50, 30, 10, '', 1],
    [14, 'スティック110mm', '本', 50, 40, 10, '', 1]
  ];
  masterSheet.getRange(2, 1, initialProducts.length, initialProducts[0].length).setValues(initialProducts);
  masterSheet.autoResizeColumns(1, masterHeaders.length);

  // 2. 日別製造計画シートの作成
  let planSheet = ss.getSheetByName('日別製造計画');
  if (!planSheet) planSheet = ss.insertSheet('日別製造計画');
  planSheet.clear();
  const planHeaders = ['年月日(YYYY-MM-DD)', '商品ID', '商品名', '計画製造数(cs)', '最終更新日時'];
  planSheet.appendRow(planHeaders);
  planSheet.getRange(1, 1, 1, planHeaders.length)
    .setBackground('#1d4ed8')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  planSheet.autoResizeColumns(1, planHeaders.length);

  // 3. 出庫実績トランザクションシートの作成
  let outSheet = ss.getSheetByName('出庫実績トランザクション');
  if (!outSheet) outSheet = ss.insertSheet('出庫実績トランザクション');
  outSheet.clear();
  const outHeaders = ['出庫年月日(YYYY-MM-DD)', '商品ID', '商品名', '出庫数(cs)', '区分(実績/予定)', '登録日時'];
  outSheet.appendRow(outHeaders);
  outSheet.getRange(1, 1, 1, outHeaders.length)
    .setBackground('#c2410c')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  outSheet.autoResizeColumns(1, outHeaders.length);

  // 4. システム設定シートの作成
  let configSheet = ss.getSheetByName('システム設定');
  if (!configSheet) configSheet = ss.insertSheet('システム設定');
  configSheet.clear();
  configSheet.appendRow(['設定項目', '設定値']);
  configSheet.getRange(1, 1, 1, 2)
    .setBackground('#475569')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  configSheet.appendRow(['運用対象年', 2026]);
  configSheet.appendRow(['運用対象月', 11]);
  configSheet.appendRow(['システム初期化日時', new Date()]);
  configSheet.autoResizeColumns(1, 2);

  // 完了ダイアログの表示
  SpreadsheetApp.getUi().alert(
    '【初期セットアップ完了】\n\n' +
    '以下の4シートを自動作成し、商品マスタを登録しました：\n' +
    '1. 商品マスタ（14品目登録済み）\n' +
    '2. 日別製造計画\n' +
    '3. 出庫実績トランザクション\n' +
    '4. システム設定\n\n' +
    '続けてWebアプリのデプロイを行ってください。'
  );
}