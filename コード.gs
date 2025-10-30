// ---------------------------------------------
// 1. WebアプリとしてHTMLを表示するための関数
// ---------------------------------------------
function doGet() {
  // index.htmlという名前のHTMLファイルをWebページとして表示
  return HtmlService.createHtmlOutputFromFile('index');
}

// ---------------------------------------------
// 2. HTML側から呼び出される「司令塔」関数 (修正版)
// ---------------------------------------------
function processUserInput(promptText) {

  // --- AIシミュレーションをバイパス ---
  // 受け取ったテキスト（promptText）が、
  // AIによって生成されたJSON（設計図）そのものであると仮定して、
  // 実行部隊（createFormFromJsonString）に直接渡します。
  
  // これにより、WebアプリのテキストエリアでJSONのテストが可能になります。
  
  try {
    // promptText がAIの生成したJSON文字列そのものとして扱う
    const result = createFormFromJsonString(promptText); 
    return result; // 成功結果（URLなど）をHTML側に返す
  } catch (e) {
    // JSON.parseのエラーなどもここでキャッチされます
    return JSON.stringify({ 
      status: 'error', 
      message: 'JSONの解析エラー: ' + e.message 
    });
  }
}


// ---------------------------------------------
// 3. JSON（設計図）を実行する「作業部隊」関数
// ---------------------------------------------
function createFormFromJsonString(jsonString) {
  try {
    const data = JSON.parse(jsonString);
    
    // フォームの基本設定
    const form = FormApp.create(data.title);
    form.setDescription(data.description);
    if (data.collectEmail === true) {
      form.setCollectEmail(true);
    }

    // 質問項目をループで作成
    data.items.forEach(itemData => {
      addItemToForm(form, itemData);
    });

    // 成功したらURLを返す (HTML側でパースできるようJSON文字列にする)
    return JSON.stringify({
      status: 'success',
      editUrl: form.getEditUrl(),
      publishedUrl: form.getPublishedUrl()
    });

  } catch (e) {
    // エラーハンドリング
    return JSON.stringify({
      status: 'error',
      message: e.message,
      stack: e.stack
    });
  }
}

// ---------------------------------------------
// 4. 質問項目を追加するヘルパー関数
// ---------------------------------------------
function addItemToForm(form, itemData) {
  let item;

  switch (itemData.type) {
    case 'section':
      item = form.addSectionHeaderItem();
      break;
    case 'text':
      item = form.addTextItem();
      break;
    case 'paragraphText':
      item = form.addParagraphTextItem();
      break;
    case 'multipleChoice':
      item = form.addMultipleChoiceItem();
      item.setChoices(itemData.choices.map(c => item.createChoice(c)));
      if (itemData.otherOption === true) item.showOtherOption(true);
      break;
    case 'checkbox':
      item = form.addCheckboxItem();
      item.setChoices(itemData.choices.map(c => item.createChoice(c)));
      if (itemData.otherOption === true) item.showOtherOption(true);
      if (itemData.validation && itemData.validation.requireSelectAtMost) {
        const validation = FormApp.createCheckboxValidation()
          .requireSelectAtMost(itemData.validation.requireSelectAtMost)
          .build();
        item.setValidation(validation);
      }
      break;
    case 'dropdown':
      item = form.addListItem();
      item.setChoices(itemData.choices.map(c => item.createChoice(c)));
      break;
    case 'scale':
      item = form.addScaleItem();
      item.setBounds(itemData.lowerBound, itemData.upperBound);
      if (itemData.lowerLabel) item.setLabels(itemData.lowerLabel, itemData.upperLabel);
      break;
    case 'date':
      item = form.addDateItem();
      if (itemData.includeTime === true) item.setIncludesTime(true);
      break;
    case 'time':
      item = form.addTimeItem();
      break;
    default:
      return;
  }

  // 共通プロパティの設定
  item.setTitle(itemData.title);
  if (itemData.helpText) {
    item.setHelpText(itemData.helpText);
  }
  // 'section' 以外は required を設定可能
  if (itemData.type !== 'section' && itemData.required === true) {
    item.setRequired(true);
  }
}
