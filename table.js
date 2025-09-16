/**
 * テーブルスライドを処理し、行・列数を動的に調整してデータを入力
 * @param {object} slide スライドのオブジェクト
 * @param {array} headers ヘッダーデータの配列
 * @param {array} rows 行データの配列
 */
function handleTableSlide(slide, headers, rows) {
  // 処理の開始をログに出力
  Logger.log("テーブルスライドの処理を開始します");
  try {
    // スライド内の表を取得
    const tables = slide.getTables();
    // 最初の表を対象とする
    const table = tables[0];
    // 現在の表の行数・列数を取得
    const currentRows = table.getNumRows();
    const currentCols = table.getNumColumns();
    // 表の目標の行数・列数を計算
    const targetRows = rows.length + 1; // ヘッダー行 + データ行
    const targetCols = headers.length;
    // 列数の調整
    adjustTableColumns(table, currentCols, targetCols);
    // 行数の調整
    adjustTableRows(table, currentRows, targetRows);
    // 調整後の実際の行数・列数を取得
    const finalRows = table.getNumRows();
    const finalCols = table.getNumColumns();
    // ヘッダーデータを入力
    headers.forEach((headerData, col) => {
      const cell = table.getCell(0, col);
      const cellText = cell.getText();
      cellText.setText(String(headerData || ""));
    });
    // 行データを入力
    rows.forEach((rowData, row) => {
      // 列数分ループ
      for (let col = 0; col < headers.length; col++) {
        const cell = table.getCell(row + 1, col);
        const cellText = cell.getText();
        cellText.setText(String(rowData?.[col] || ""));
      }
    });
    // 処理の完了をログに出力
    Logger.log("テーブルスライドの処理が完了しました");
  } catch (e) {
    // 処理の失敗をログに出力
    Logger.log(`テーブルスライドの処理中にエラーが発生しました: ${e.message}`);
  }
}

/**
 * 表の列数を調整。削除はAPIで不可能なので追加のみ。削除が不要なようにテンプレートスライドの表の列数は1列のみ。
 * @param {object} table 表オブジェクト
 * @param {number} currentCols 現在の列数
 * @param {number} targetCols 目標の列数
 */
function adjustTableColumns(table, currentCols, targetCols) {
  // 処理の開始をログに出力
  Logger.log(`列の調整を開始します`);
  try {
    if (targetCols > currentCols) {
      Array.from({ length: targetCols - currentCols }).forEach((_, index) => {
        const insertColumnIndex = currentCols + index;
        table.insertColumn(insertColumnIndex);
      });
    }
    // 処理の完了をログに出力
    Logger.log(`列の調整が完了しました`);
  } catch (e) {
    // 処理の失敗をログに出力
    Logger.log(`列の調整中にエラーが発生しました: ${e.message}`);
  }
}

/**
 * 表の行数を調整。削除はAPIで不可能なので追加のみ。削除が不要なようにテンプレートスライドの表の行数は1行のみ。
 * @param {object} table 表オブジェクト
 * @param {number} currentRows 現在の行数
 * @param {number} targetRows 目標の行数
 */
function adjustTableRows(table, currentRows, targetRows) {
  // 処理の開始をログに出力
  Logger.log(`行の調整を開始します`);
  try {
    if (targetRows > currentRows) {
      Array.from({ length: targetRows - currentRows }).forEach((_, index) => {
        const insertRowIndex = currentRows + index;
        table.insertRow(insertRowIndex);
      });
    }
    // 処理の完了をログに出力
    Logger.log(`行の調整が完了しました`);
  } catch (e) {
    // 処理の失敗をログに出力
    Logger.log(`行の調整中にエラーが発生しました: ${e.message}`);
  }
}
