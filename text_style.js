/**
 * スライドのテキストボックス内の太文字、重要語に対してスタイルを適用
 * @param {object} slide スライドのオブジェクト
 */
function handleTextStyle(slide) {
  // 処理の開始をログに出力
  Logger.log(`スライドのテキストスタイル適用を開始します`);
  // スライド内の全図形を取得
  const shapes = slide.getShapes();
  // テキストボックスのみを対象とした、処理を実施
  shapes.forEach((shape) => {
    if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
      processTextBox(shape);
    }
  });
  // 処理の完了をログに出力
  Logger.log(`スライドのテキストスタイル適用が完了しました`);
}

/**
 * テキストボックス単体の処理
 * @param {object} textBox テキストボックスのオブジェクト
 */
function processTextBox(textBox) {
  // テキストボックス内のテキストを取得
  const textRange = textBox.getText();
  const text = textRange.asString();
  // 太字(**で始まる) or 重要語([[で始まる)が含まれているかチェック
  if (!text.includes("**") && !text.includes("[[")) {
    return; // マーカーがない場合は処理をスキップ
  }
  // マーカー情報を取得
  const markers = findTextMarkers(text);
  // マーカーを除去してスタイルを適用する範囲を計算
  const { processedText, styleRanges } = removeMarkersAndCalculateStyleRanges(
    text,
    markers
  );
  // マーカー除去後のテキストを設定
  textRange.setText(processedText);
  // スタイルを適用
  applyTextStyle(textRange, styleRanges);
}

/**
 * テキスト内のマーカー（太字・重要語）を検索して位置情報を取得
 * @param {string} text 検索対象のテキスト
 * @returns {array} マーカー情報の配列
 */
function findTextMarkers(text) {
  // マーカー情報を格納する配列
  const markers = [];
  // マッチ結果を格納する変数
  let match;
  // 太字マーカー（**text**）を検索
  const boldPattern = /\*\*(.+?)\*\*/g;
  /**
   * テキスト内の太字マーカー毎にループ処理
   * 例: text = "これは**太字**のテスト"
   * boldPattern.exec(text) の返り値 match には以下が格納される:
   *  match[0]: マッチした全文字列（例: "**太字**"）
   *  match[1]: グループ化した部分（例: "太字"）
   *  match.index: マッチ開始位置（例: 3）
   */
  while ((match = boldPattern.exec(text)) !== null) {
    markers.push({
      start: match.index, // マーカーの開始位置（例: 3）
      end: match.index + match[0].length, // マーカーの終了位置（例: 3 + 6 = 9）
      content: match[1], // マーカー内のテキスト（例: "太字"）
      type: "bold",
    });
  }
  // 重要語マーカー（[[text]]）を検索
  const importantPattern = /\[\[(.+?)\]\]/g;
  /**
   * テキスト内の重要語マーカー毎にループ処理
   * 例: text = "これは[[重要語]]のテスト"
   * importantPattern.exec(text) の返り値 match には以下が格納される:
   *  match[0]: マッチした全文字列（例: "[[重要語]]"）
   *  match[1]: グループ化した部分（例: "重要語"）
   *  match.index: マッチ開始位置（例: 3）
   */
  while ((match = importantPattern.exec(text)) !== null) {
    markers.push({
      start: match.index,
      end: match.index + match[0].length,
      content: match[1],
      type: "important",
    });
  }
  // 位置順でソート（前から処理するため昇順）
  return markers.sort((a, b) => a.start - b.start);
}

/**
 * マーカーを除去してスタイル範囲を計算
 * @param {string} text 元のテキスト
 * @param {array} markers マーカー情報の配列
 * @returns {object} マーカー除去後のテキスト(processedText), スタイル範囲(styleRanges)
 */
function removeMarkersAndCalculateStyleRanges(text, markers) {
  // マーカーを除去したテキストを格納する変数
  let processedText = text;
  // スタイル範囲を格納する変数
  let styleRanges = [];
  // 累積オフセットを格納する変数
  let cumulativeOffset = 0;
  // マーカー毎にループ処理
  markers.forEach((marker) => {
    // オフセットを適用した実際の位置
    const actualStart = marker.start - cumulativeOffset;
    const actualEnd = marker.end - cumulativeOffset;
    // 処理中のマーカーを除去したテキストを作成
    processedText =
      // 処理中のマーカーの開始位置までのテキスト
      processedText.substring(0, actualStart) +
      // 処理中のマーカーの中身
      marker.content +
      // 処理中のマーカーの終了位置以降のテキスト
      processedText.substring(actualEnd);
    // スタイル範囲を記録（オフセット適用後の位置で）
    styleRanges.push({
      start: actualStart,
      end: actualStart + marker.content.length,
      type: marker.type,
    });
    // 除去された文字数を累積オフセットに加算
    cumulativeOffset += marker.end - marker.start - marker.content.length;
  });
  // 処理結果を返却
  return { processedText, styleRanges };
}

/**
 * テキストにスタイルを適用
 * @param {object} textRange テキストのオブジェクト
 * @param {array} styleRanges スタイルを適用するテキストの情報が格納された配列
 */
function applyTextStyle(textRange, styleRanges) {
  styleRanges.forEach((styleRange) => {
    // 範囲を取得
    const textSubRange = textRange.getRange(styleRange.start, styleRange.end);
    if (textSubRange) {
      const textStyle = textSubRange.getTextStyle();
      // 太字の場合
      if (styleRange.type === "bold") {
        textStyle.setBold(true);
      }
      // 重要語の場合
      else if (styleRange.type === "important") {
        textStyle.setBold(true);
        textStyle.setForegroundColor("#0E7BCF");
      }
    }
  });
}
