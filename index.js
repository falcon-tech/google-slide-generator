// 既存のスライドを削除するかどうか
const DELETE_ALREADY_SLIDES = true;
// デバッグモードかどうか
const DEBUG = false;
// テスト用のスライドデータ
const testSlideData = [
  {
    type: "title",
    to: "クライアント 御中",
    title: "メインタイトル",
    body: "本文",
    date: "2025.08.29",
    notes: "スピーカノート",
  },
  {
    type: "agenda",
    title: "アジェンダ",
    items: ["アイテム1", "アイテム2", "アイテム3"],
    notes: "スピーカノート",
  },
  {
    type: "section",
    title: "章タイトル",
    notes: "スピーカノート",
  },
  {
    type: "bullet",
    title: "箇条書き",
    header: "ヘッダー",
    items: ["**アイテム1**", "[[アイテム2]]", "アイテム3"],
    notes: "スピーカノート",
  },
  {
    type: "compare_two_boxes",
    title: "比較",
    description: "比較の説明",
    box1_header: "ボックス1ヘッダー",
    box1_items: [
      "ボックス1アイテム1",
      "ボックス1アイテム2",
      "ボックス1アイテム3",
    ],
    box2_header: "ボックス2ヘッダー",
    box2_items: [
      "**ボックス2**[[アイテム1]]",
      "**ボックス2アイテム2**",
      "[[ボックス2アイテム3]]",
    ],
    notes: "スピーカノート",
  },
  {
    type: "compare_three_boxes",
    title: "比較",
    description: "比較の説明",
    box1_header: "ボックス1ヘッダー",
    box1_items: [
      "ボックス1アイテム1",
      "ボックス1アイテム2",
      "ボックス1アイテム3",
    ],
    box2_header: "ボックス2ヘッダー",
    box2_items: [
      "**ボックス2**[[アイテム1]]",
      "**ボックス2アイテム2**",
      "[[ボックス2アイテム3]]",
    ],
    box3_header: "ボックス3ヘッダー",
    box3_items: [
      "**ボックス3**[[アイテム1]]",
      "**ボックス3アイテム2**",
      "[[ボックス3アイテム3]]",
    ],
    notes: "スピーカノート",
  },
  {
    type: "compare_four_boxes",
    title: "比較",
    description: "比較の説明",
    box1_header: "ボックス1ヘッダー",
    box1_items: [
      "ボックス1アイテム1",
      "ボックス1アイテム2",
      "ボックス1アイテム3",
    ],
    box2_header: "ボックス2ヘッダー",
    box2_items: [
      "**ボックス2**[[アイテム1]]",
      "**ボックス2アイテム2**",
      "[[ボックス2アイテム3]]",
    ],
    box3_header: "ボックス3ヘッダー",
    box3_items: [
      "**ボックス3**[[アイテム1]]",
      "**ボックス3アイテム2**",
      "[[ボックス3アイテム3]]",
    ],
    box4_header: "ボックス4ヘッダー",
    box4_items: [
      "**ボックス4**[[アイテム1]]",
      "**ボックス4アイテム2**",
      "[[ボックス4アイテム3]]",
    ],
    notes: "スピーカノート",
  },
  {
    type: "table",
    title: "テーブル",
    description: "テーブルの説明",
    headers: ["ヘッダー1", "ヘッダー2", "ヘッダー3", "ヘッダー4"],
    rows: [
      ["データ1-1", "データ2-1", "データ3-1", "データ4-1"],
      ["データ1-2", "データ2-2", "データ3-2", "データ4-2"],
      ["データ1-3", "データ2-3", "データ3-3", "データ4-3"],
    ],
    notes: "スピーカノート",
  },
  {
    type: "closing",
    notes: "スピーカノート",
  },
];

/**
 * アドオンをインストールした際に実行される関数
 */
function onInstall() {
  onOpen();
}

/**
 * プレゼンテーションを開いた際に実行される関数
 */
function onOpen() {
  const menu = SlidesApp.getUi();
  menu
    .createAddonMenu()
    // ユーザがメニューから「プレゼンテーション生成」を選択した際に、showDataInputDialog関数を実行
    .addItem("プレゼンテーション生成", "showDataInputDialog")
    .addToUi();
}

/**
 * JSONデータ入力用のダイアログを表示
 */
function showDataInputDialog() {
  const html = HtmlService.createTemplateFromFile("dialog")
    .evaluate()
    .setWidth(400)
    .setHeight(300);

  const dialog = SlidesApp.getUi();
  dialog.showModalDialog(html, "スライドデータを入力");
}

/**
 * ユーザーから入力されたJSONデータでプレゼンテーションを生成。クライアントサイドのgeneratePresentation関数から呼び出される
 * @param {string} jsonData JSON形式のスライドデータ
 */
function generatePresentation(jsonData) {
  // 処理の開始をログに出力
  Logger.log(`プレゼンテーション生成を開始します`);
  try {
    // ユーザーから入力されたJSONデータをパース。デバッグモードの場合はtestSlideDataを使用
    const slideData = DEBUG ? testSlideData : JSON.parse(jsonData);
    // プレゼンテーションを取得
    const presentation = SlidesApp.getActivePresentation();
    // 既存のスライドを削除
    if (DELETE_ALREADY_SLIDES) {
      Logger.log(`既存のスライドを削除します`);
      const slides = presentation.getSlides();
      slides.forEach((slide) => slide.remove());
    }
    // スライドデータの内容を元に、スライドを生成
    slideData.forEach((data) => createSlide(presentation, data));
    // 処理の完了をログに出力
    Logger.log(`プレゼンテーション生成が完了しました`);
  } catch (e) {
    // 処理の失敗をログに出力
    Logger.log(`プレゼンテーション生成中にエラーが発生しました: ${e.message}`);
    // エラーを返却(ダイアログに表示)
    throw e;
  }
}

/**
 * スクリプトプロパティからテンプレート設定(プレゼンテーションIDとスライドID)を読み込みする関数。グローバルスコープでスクリプトプロパティを読み込むと、アドオンの初回インストール以降の実行で、遅延か何かで権限エラーでスクリプトプロパティが読み込めないので、関数スコープで読み込むようにしている。
 * @returns {object} テンプレート設定のオブジェクト
 */
function getTemplateConfig() {
  // 処理の開始をログに出力
  Logger.log(`テンプレート設定を取得します`);
  const properties = PropertiesService.getScriptProperties();
  return {
    presentationId: properties.getProperty("TEMPLATE_PRESENTATION_ID"),
    slideId: {
      title: properties.getProperty("TEMPLATE_SLIDE_ID_TITLE"),
      agenda: properties.getProperty("TEMPLATE_SLIDE_ID_AGENDA"),
      section: properties.getProperty("TEMPLATE_SLIDE_ID_SECTION"),
      compare_two_boxes: properties.getProperty(
        "TEMPLATE_SLIDE_ID_COMPARE_TWO_BOXES"
      ),
      compare_three_boxes: properties.getProperty(
        "TEMPLATE_SLIDE_ID_COMPARE_THREE_BOXES"
      ),
      compare_four_boxes: properties.getProperty(
        "TEMPLATE_SLIDE_ID_COMPARE_FOUR_BOXES"
      ),
      bullet: properties.getProperty("TEMPLATE_SLIDE_ID_BULLET"),
      table: properties.getProperty("TEMPLATE_SLIDE_ID_TABLE"),
      closing: properties.getProperty("TEMPLATE_SLIDE_ID_CLOSING"),
    },
  };
}

/**
 * スライドを生成
 * @param {object} presentation プレゼンテーションのオブジェクト
 * @param {object} data スライドデータのオブジェクト
 */
function createSlide(presentation, data) {
  // 処理の開始をログに出力
  Logger.log(`スライド(${data.type})を生成します`);
  // テンプレート設定を取得
  const templateConfig = getTemplateConfig();
  // ソーススライドを取得
  const sourceSlide = getSlide(
    templateConfig.presentationId,
    templateConfig.slideId[data.type]
  );
  // ソーススライドが見つからない場合はエラーを投げる
  if (!sourceSlide) {
    throw new Error(
      `指定されたスライド「${
        templateConfig.slideId[data.type]
      }」が見つかりませんでした。`
    );
  }
  // ソーススライドを複製して、プレゼンテーションに追加
  const slide = presentation.appendSlide(
    sourceSlide,
    SlidesApp.SlideLinkingMode.UNLINKED
  );
  // テキストを置き換え
  switch (data.type) {
    case "title":
      slide.replaceAllText("{{to}}", data.to);
      slide.replaceAllText("{{title}}", data.title);
      slide.replaceAllText("{{body}}", data.body);
      slide.replaceAllText("{{date}}", data.date);
      break;
    case "agenda":
      slide.replaceAllText("{{title}}", data.title);
      slide.replaceAllText("{{items}}", data.items.join("\n"));
      break;
    case "section":
      slide.replaceAllText("{{title}}", data.title);
      break;
    case "compare_two_boxes":
      slide.replaceAllText("{{title}}", data.title);
      slide.replaceAllText("{{description}}", data.description);
      slide.replaceAllText("{{box1_header}}", data.box1_header);
      slide.replaceAllText("{{box1_items}}", data.box1_items.join("\n"));
      slide.replaceAllText("{{box2_header}}", data.box2_header);
      slide.replaceAllText("{{box2_items}}", data.box2_items.join("\n"));
      break;
    case "compare_three_boxes":
      slide.replaceAllText("{{title}}", data.title);
      slide.replaceAllText("{{description}}", data.description);
      slide.replaceAllText("{{box1_header}}", data.box1_header);
      slide.replaceAllText("{{box1_items}}", data.box1_items.join("\n"));
      slide.replaceAllText("{{box2_header}}", data.box2_header);
      slide.replaceAllText("{{box2_items}}", data.box2_items.join("\n"));
      slide.replaceAllText("{{box3_header}}", data.box3_header);
      slide.replaceAllText("{{box3_items}}", data.box3_items.join("\n"));
      break;
    case "compare_four_boxes":
      slide.replaceAllText("{{title}}", data.title);
      slide.replaceAllText("{{description}}", data.description);
      slide.replaceAllText("{{box1_header}}", data.box1_header);
      slide.replaceAllText("{{box1_items}}", data.box1_items.join("\n"));
      slide.replaceAllText("{{box2_header}}", data.box2_header);
      slide.replaceAllText("{{box2_items}}", data.box2_items.join("\n"));
      slide.replaceAllText("{{box3_header}}", data.box3_header);
      slide.replaceAllText("{{box3_items}}", data.box3_items.join("\n"));
      slide.replaceAllText("{{box4_header}}", data.box4_header);
      slide.replaceAllText("{{box4_items}}", data.box4_items.join("\n"));
      break;
    case "bullet":
      slide.replaceAllText("{{title}}", data.title);
      slide.replaceAllText("{{header}}", data.header);
      slide.replaceAllText("{{items}}", data.items.join("\n"));
      break;
    case "table":
      slide.replaceAllText("{{title}}", data.title);
      slide.replaceAllText("{{description}}", data.description);
      handleTableSlide(slide, data.headers, data.rows);
      break;
    case "closing":
      break;
  }
  // テキスト上の太字と重要語に対してスタイルを適用
  handleTextStyle(slide);
  // スピーカーノートを設定
  if (data.notes) {
    const notesShape = slide.getNotesPage().getSpeakerNotesShape();
    if (notesShape) {
      notesShape.getText().setText(data.notes);
    }
  }
  // 処理の完了をログに出力
  Logger.log(`スライド(${data.type})の生成が完了しました`);
}

/**
 * テンプレートプレゼンテーションIDとスライドのオブジェクトIDを指定して、スライドを取得
 * @param {string} templatePresentationId テンプレートのプレゼンテーションID
 * @param {string} slideId 取得したいスライドのオブジェクトID
 * @returns {object | undefined} スライドのオブジェクトを返却。スライドが見つからない場合はundefinedを返却
 */
function getSlide(templatePresentationId, slideId) {
  // 処理の開始をログに出力
  Logger.log(`テンプレートスライド(${slideId})を取得します`);
  // テンプレートプレゼンテーションを取得
  const templatePresentation = SlidesApp.openById(templatePresentationId);
  // テンプレートプレゼンテーションからスライド一覧を取得
  const slides = templatePresentation.getSlides();
  // 指定のオブジェクトIDでスライドを取得
  return slides.find((slide) => slide.getObjectId() === slideId);
}
