/**
 * @OnlyCurrentDoc
 */

// =================================================================
// 1. 定数定義
// =================================================================

const PRESET_SHEET_NAME = 'GF式神';
const SETTINGS_SHEET_NAME = '設定';
const PRESET_NAME_HEADER = '名前';

const SETTINGS_KEY_TO_HEADER = {
  folderUrl: '保存先フォルダURL',
  collectEmail: 'メールアドレスを収集する',
  sendCopy: '回答のコピーを回答者に送信'
};

const HEADER_TO_SETTINGS_KEY = Object.fromEntries(
  Object.entries(SETTINGS_KEY_TO_HEADER).map(([key, value]) => [value, key])
);

// =================================================================
// 2. Webアプリ起動 (doGet)
// =================================================================

function doGet(e) {
  let initialPresets = {};
  let appSettings = {};
  
  try {
    initialPresets = getPresetsFromSheet();
    appSettings = getAppSettingsFromSheet();
    
    const htmlTemplate = HtmlService.createTemplateFromFile('index.html');
    
    htmlTemplate.presets = initialPresets;
    htmlTemplate.appSettings = appSettings;

    return htmlTemplate.evaluate()
      .setTitle('Gフォーム式神 v2')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);

  } catch (e) {
    Logger.log('doGet実行時エラー: ' + e.message);
    return HtmlService.createHtmlOutput(
      '<body style="background:#1d232a; color:white; font-family:sans-serif;">' +
      '<h1>Gフォーム式神 v2 - 起動エラー</h1>' +
      '<p style="color:#f87171;">初期設定の読み込みに失敗しました。</p>' +
      '<p><strong>エラー内容:</strong> ' + e.message + '</p>' +
      '<p style="color:#FFF;"><strong>原因:</strong> スプレッドシートの権限が承認されていないか、シート名が間違っています。</p>' +
      '<ol style="color:#DDD;">' +
        '<li>シート名が「' + SETTINGS_SHEET_NAME + '」および「' + PRESET_SHEET_NAME + '」になっているか確認してください。</li>' +
        '<li>「デプロイ」→「新しいデプロイ」を実行し、権限を再承認してください。</li>' +
      '</ol>' +
      '</body>'
    ).setTitle('Gフォーム式神 v2 - エラー');
  }
}

/**
 * 「設定」シートからキーバリューペアで設定を読み込む
 */
function getAppSettingsFromSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    if (!sheet) {
      throw new Error("「" + SETTINGS_SHEET_NAME + "」シートが見つかりません。");
    }
    
    const data = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues(); 
    const settings = {};
    
    for (let i = 0; i < data.length; i++) {
      const key = data[i][0]; 
      const value = data[i][1]; 
      if (key && typeof key === 'string' && key.trim() !== '' && 
          value && typeof value === 'string' && value.trim() !== '') {
        settings[key.trim()] = value.trim();
      }
    }
    
    if (!settings['式神GフォームGem']) {
       throw new Error("「" + SETTINGS_SHEET_NAME + "」シートに「式神GフォームGem」のURLが設定されていません。");
    }
    
    return settings;
  } catch (e) {
    Logger.log('設定シート読み込みエラー: ' + e.message);
    throw e; 
  }
}


// =================================================================
// 3. プリセット操作関数 (HTMLから呼び出される)
// =================================================================

function getPresetsFromSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(PRESET_SHEET_NAME);
    if (!sheet) {
       throw new Error("「" + PRESET_SHEET_NAME + "」シートが見つかりません。");
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return {}; 
    
    const headers = data[0];
    const presets = {};
    const nameColIndex = headers.indexOf(PRESET_NAME_HEADER);
    if (nameColIndex === -1) {
      throw new Error("「" + PRESET_SHEET_NAME + "」シートに「" + PRESET_NAME_HEADER + "」列が必要です。");
    }

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const presetName = row[nameColIndex];
      if (!presetName) continue;
      
      const settings = {};
      headers.forEach((header, colIndex) => {
        const key = HEADER_TO_SETTINGS_KEY[header];
        if (key) {
          let value = row[colIndex];
          if (['collectEmail', 'sendCopy'].includes(key)) {
            value = (value === true || String(value).toUpperCase() === 'TRUE');
          }
          settings[key] = value;
        }
      });
      presets[presetName] = settings;
    }
    return presets;
  } catch (e) {
    Logger.log('プリセット読み込みエラー: ' + e.message);
    throw e; 
  }
}

function savePresetToSheet(presetName, settings) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(PRESET_SHEET_NAME);
    if (!sheet) throw new Error("「" + PRESET_SHEET_NAME + "」シートが見つかりません。");
    
    const data = sheet.getDataRange().getValues();
    const headers = data.length > 0 ? data[0] : [];
    const nameColIndex = headers.indexOf(PRESET_NAME_HEADER);
    if (nameColIndex === -1) throw new Error("シートに「" + PRESET_NAME_HEADER + "」列が必要です。");

    let targetRowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][nameColIndex] === presetName) {
        targetRowIndex = i + 1; 
        break;
      }
    }

    const newRowData = headers.map(header => {
      if (header === PRESET_NAME_HEADER) return presetName;
      const key = HEADER_TO_SETTINGS_KEY[header];
      return key && settings.hasOwnProperty(key) ? settings[key] : '';
    });

    if (targetRowIndex !== -1) {
      sheet.getRange(targetRowIndex, 1, 1, newRowData.length).setValues([newRowData]);
    } else {
      sheet.appendRow(newRowData);
    }
    
    return { status: 'success', message: 'プリセットを保存しました。' };
  } catch (e) {
    Logger.log('プリセット保存エラー: ' + e.message);
    return { status: 'error', message: 'プリセットの保存に失敗しました: ' + e.message };
  }
}

function deletePresetFromSheet(presetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(PRESET_SHEET_NAME);
    if (!sheet) throw new Error("「" + PRESET_SHEET_NAME + "」シートが見つかりません。");
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const nameColIndex = headers.indexOf(PRESET_NAME_HEADER);
    if (nameColIndex === -1) throw new Error("シートに「" + PRESET_NAME_HEADER + "」列が必要です。");
    
    let targetRowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][nameColIndex] === presetName) {
        targetRowIndex = i + 1; 
        break;
      }
    }

    if (targetRowIndex !== -1) {
      sheet.deleteRow(targetRowIndex);
      return { status: 'success', message: 'プリセットを削除しました。' };
    } else {
      return { status: 'info', message: '削除対象のプリセットが見つかりませんでした。' };
    }
  } catch (e) {
    Logger.log('プリセット削除エラー: ' + e.message);
    return { status: 'error', message: 'プリセットの削除に失敗しました: ' + e.message };
  }
}


// =================================================================
// 4. フォーム構築関数
// =================================================================

function processUserInput(promptText, presetName) {
  try {
    const allPresets = getPresetsFromSheet();
    const settings = allPresets[presetName];
    if (!settings) {
      throw new Error('指定されたプリセット「' + presetName + '」が見つかりません。');
    }

    let jsonData;
    try {
      jsonData = JSON.parse(promptText);
    } catch (e) {
      throw new Error('JSONの形式が正しくありません: ' + e.message);
    }

    const form = FormApp.create(jsonData.title || '無題のフォーム');
    
    if (jsonData.description) {
      form.setDescription(jsonData.description);
    }
    
    form.setCollectEmail(settings.collectEmail === true);
    form.setAllowResponseEdits(false); 
    
    if (settings.sendCopy === true) {
      form.setConfirmationMessage('回答のコピーが送信されました。');
      form.setLimitOneResponsePerUser(true); 
    }

    if (settings.folderUrl) {
      const folderId = getFolderIdFromUrl(settings.folderUrl);
      if (folderId) {
        try {
          const formFile = DriveApp.getFileById(form.getId());
          const folder = DriveApp.getFolderById(folderId);
          formFile.moveTo(folder);
        } catch (folderError) {
          Logger.log('フォルダ移動に失敗: ' + folderError.message);
        }
      }
    }

    if (Array.isArray(jsonData.items)) {
      jsonData.items.forEach(item => {
        addFormItem(form, item);
      });
    }

    return JSON.stringify({
      status: 'success',
      publishedUrl: form.getPublishedUrl(),
      editUrl: form.getEditUrl()
    });

  } catch (e) {
    Logger.log('processUserInputエラー: ' + e.message); 
    return JSON.stringify({ status: 'error', message: e.message });
  }
}

function addFormItem(form, item) {
  if (!item || !item.type) return;
  let formItem;
  switch (item.type) {
    case 'text': formItem = form.addTextItem(); break;
    case 'paragraphText': formItem = form.addParagraphTextItem(); break;
    case 'multipleChoice':
      formItem = form.addMultipleChoiceItem();
      if (Array.isArray(item.choices)) formItem.setChoiceValues(item.choices);
      break;
    case 'checkbox':
      formItem = form.addCheckboxItem();
      if (Array.isArray(item.choices)) formItem.setChoiceValues(item.choices);
      break;
    case 'list':
      formItem = form.addListItem();
      if (Array.isArray(item.choices)) formItem.setChoiceValues(item.choices);
      break;
    case 'date': formItem = form.addDateItem(); break;
    case 'time': formItem = form.addTimeItem(); break;
    case 'scale':
      formItem = form.addScaleItem();
      formItem.setBounds(item.lowerBound || 1, item.upperBound || 5);
      if(item.leftLabel) formItem.setLabels(item.leftLabel, item.rightLabel || '');
      break;
    default: return; 
  }
  if (item.title) formItem.setTitle(item.title);
  if (item.helpText) formItem.setHelpText(item.helpText);
  if (item.required) formItem.setRequired(item.required === true);
}

function getFolderIdFromUrl(url) {
  if (!url) return null;
  const match = url.match(/\/folders\/([a-zA-Z0_9-]+)/);
  return match ? match[1] : null;
}

// =================================================================
// 5. ▼▼▼ HTMLインクルード関数 (必須) ▼▼▼
// =================================================================

/**
 * 別のHTMLファイル(この場合はJSライブラリ)をインクルードするための関数
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}
