/* Code.gs */

// 定数設定
const FOLDER_NAMES = ['Recordings', 'ShadowingResults', 'AudioFiles', 'InboxSubmissions', 'Processed', 'PassageBooks'];
const MASTER_SHEET_NAME = 'ResultMaster';
const WHITELIST_SHEET_NAME = 'whitelist';

/**
 * Webアプリのエントリーポイント
 */
function doGet() {
  const template = HtmlService.createTemplateFromFile('index');
  template.userEmail = Session.getActiveUser().getEmail();
  return template.evaluate()
    .setTitle('Shadowing Training App')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 初回起動時の環境構築
 * 親フォルダ内に必要なフォルダ群とMasterスプレッドシートを作成・取得
 * 追加仕様: PassageBooksが空の場合、サンプルブックを生成
 */
function setupEnvironment() {
  const scriptId = ScriptApp.getScriptId();
  const scriptFile = DriveApp.getFileById(scriptId);
  const parentFolder = scriptFile.getParents().next(); // スクリプトのある親フォルダ

  const env = {};

  // 1. フォルダの確認・作成
  FOLDER_NAMES.forEach(name => {
    const folders = parentFolder.getFoldersByName(name);
    if (folders.hasNext()) {
      env[name] = folders.next().getId();
    } else {
      const newFolder = parentFolder.createFolder(name);
      env[name] = newFolder.getId();
    }
  });

  // --- 【仕様追加】PassageBooksフォルダのチェックとサンプル生成 ---
  const pbFolder = DriveApp.getFolderById(env['PassageBooks']);
  if (!pbFolder.getFilesByType(MimeType.GOOGLE_SHEETS).hasNext()) {
    createSampleBook(pbFolder);
  }
  // -------------------------------------------------------------

  // 2. Masterスプレッドシートの確認・作成
  const files = parentFolder.getFilesByName('ShadowingApp_Master');
  let ss;
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create('ShadowingApp_Master');
    DriveApp.getFileById(ss.getId()).moveTo(parentFolder);
  }

  // ResultMasterシート設定
  let resultSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!resultSheet) {
    resultSheet = ss.insertSheet(MASTER_SHEET_NAME);
    resultSheet.appendRow(['Timestamp', 'Book', 'Unit', 'TaskID', 'UserID', 'ShadowingScore', 'ReadingScore', 'Speed', 'JSON_File', 'Audio_File']); // ヘッダー
  } else {
    // 既存シートがある場合、ヘッダー行を確認して更新（簡易的実装）
    const header = resultSheet.getRange(1, 1, 1, 10).getValues()[0];
    if (header[5] === 'Score') {
      // 古いヘッダーの場合は警告ログを出すか、ユーザーに手動対応を促す（破壊的変更を避けるためここでは変更しないが、新規行は新形式で追加される）
      console.warn("ResultMaster has old header format. New columns will be appended.");
    }
  }

  // Whitelistシート設定
  let whiteSheet = ss.getSheetByName(WHITELIST_SHEET_NAME);
  if (!whiteSheet) {
    whiteSheet = ss.insertSheet(WHITELIST_SHEET_NAME);
    whiteSheet.appendRow(['Email', 'Name', '4DigitID']); // ヘッダー
    // 実行者自身の情報をデモ用に追加
    whiteSheet.appendRow([Session.getActiveUser().getEmail(), 'Demo User', '0000']);
  }

  return { folderIds: env, masterSsId: ss.getId() };
}

/**
 * サンプル教材ブックの作成（PassageBooksが空の時に実行）
 */
function createSampleBook(folder) {
  const bookName = "Sample_Grade1_English";
  const sheetName = "Unit1_Greetings";

  const ss = SpreadsheetApp.create(bookName);
  const file = DriveApp.getFileById(ss.getId());
  file.moveTo(folder);

  const sheet = ss.getSheets()[0];
  sheet.setName(sheetName);

  const headers = ['ID', 'Task Title', 'Reference Text (Script)', 'Display Text (Hidden)', 'Audio Filename'];
  const sampleTasks = [
    [1, 'Hello', 'Hello, nice to meet you.', '', 'sample_hello.opus'],
    [2, 'Morning', 'Good morning. How are you doing?', '', 'sample_morning.opus'],
    [3, 'Weather', 'It is sunny today, isn\'t it?', '', 'sample_weather.opus']
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, sampleTasks.length, sampleTasks[0].length).setValues(sampleTasks);

  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f3f3');
  sheet.setColumnWidth(3, 400);
  sheet.setColumnWidth(5, 200);
}

/**
 * ユーザー情報の取得（Whitelist照合）
 */
function getUserInfo() {
  const email = Session.getActiveUser().getEmail();
  const env = setupEnvironment(); // ここで初回セットアップも兼ねる
  const ss = SpreadsheetApp.openById(env.masterSsId);
  const sheet = ss.getSheetByName(WHITELIST_SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      return {
        name: data[i][1],
        id4: data[i][2],
        isAuthorized: true
      };
    }
  }

  return { name: 'Guest', id4: '0000', isAuthorized: false };
}

/**
 * PassageBooksフォルダ内の教材構造を取得（キャッシュ対応）
 */
function getMaterialsStructure() {
  const props = PropertiesService.getScriptProperties();
  const cached = props.getProperty('MATERIALS_CACHE');

  if (cached) {
    return JSON.parse(cached);
  }

  return updateMaterialsCache();
}

/**
 * 教材構造のキャッシュ強制更新（管理者・任意更新用）
 */
function updateMaterialsCache() {
  const env = setupEnvironment();
  const folder = DriveApp.getFolderById(env.folderIds['PassageBooks']);
  const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);

  const structure = {};

  while (files.hasNext()) {
    const file = files.next();
    const bookName = file.getName();
    const ss = SpreadsheetApp.open(file);
    structure[bookName] = {};

    const sheets = ss.getSheets();
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      const rows = sheet.getDataRange().getValues();
      const tasks = [];

      // 列構成: 通し番号(0), 見出し(1), 参照文(2), 表示文(3), 音声ファイル名(4)
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0]) { // IDがあれば有効とする（音声ファイルがない場合はReadingOnly）
          tasks.push({
            id: rows[i][0],
            title: rows[i][1],
            refText: rows[i][2],
            displayText: rows[i][3], // 表示文も取得
            audioName: rows[i][4] || "" // 空文字許容
          });
        }
      }
      if (tasks.length > 0) {
        structure[bookName][sheetName] = tasks;
      }
    });
  }

  PropertiesService.getScriptProperties().setProperty('MATERIALS_CACHE', JSON.stringify(structure));
  return structure;
}

/**
 * 音声ファイル名からBase64データを取得
 */
function getAudioData(filename) {
  const env = setupEnvironment();
  const folder = DriveApp.getFolderById(env.folderIds['AudioFiles']);
  const files = folder.getFilesByName(filename);

  if (files.hasNext()) {
    const file = files.next();
    const blob = file.getBlob();
    return {
      name: filename,
      mime: blob.getContentType(),
      base64: Utilities.base64Encode(blob.getBytes())
    };
  } else {
    // ファイルがない場合のエラーハンドリング
    throw new Error('Audio file not found: ' + filename);
  }
}

/**
 * 結果の保存（JSONと音声）
 */
function saveSubmission(payload) {
  // 1. Get User Email safely
  const email = Session.getActiveUser().getEmail();

  // 2. Get Shared Folder
  const props = PropertiesService.getScriptProperties();
  const folderId = props.getProperty('SHARED_FOLDER_ID');
  if (!folderId) throw new Error("Configuration Error: SHARED_FOLDER_ID not set.");

  let folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    throw new Error("Cannot access storage folder. Please contact admin.");
  }

  // 3. Inject Email into JSON
  const data = JSON.parse(payload.jsonStr);
  data.verifiedEmail = email; // FORCE OVERWRITE or ADD
  const secureJsonStr = JSON.stringify(data);

  // 4. Save Files (User must have write permission)
  const jsonName = payload.filenameBase + '.json';
  folder.createFile(jsonName, secureJsonStr, MimeType.PLAIN_TEXT);

  const audioName = payload.filenameBase + '.webm';
  const audioBlob = Utilities.newBlob(
    Utilities.base64Decode(payload.audioBase64),
    'audio/webm;codecs=opus',
    audioName
  );
  folder.createFile(audioBlob);

  return "Success";
}

/**
 * トリガー実行用：InboxSubmissionsの処理
 */
function processInboxQueue() {
  const env = setupEnvironment();
  const inbox = DriveApp.getFolderById(env.folderIds['InboxSubmissions']);
  const processed = DriveApp.getFolderById(env.folderIds['Processed']);
  const files = inbox.getFiles();
  const ss = SpreadsheetApp.openById(env.masterSsId);
  const sheet = ss.getSheetByName(MASTER_SHEET_NAME);

  while (files.hasNext()) {
    const file = files.next();
    try {
      const content = file.getBlob().getDataAsString();
      const data = JSON.parse(content);

      sheet.appendRow([
        new Date(),
        data.book,
        data.unit,
        data.taskId,
        data.userId,
        data.shadowing_score, // New
        data.reading_score,   // New
        data.playback_speed,  // New
        file.getName(),
        data.filenameBase + '.webm'
      ]);

      file.moveTo(processed);
    } catch (e) {
      console.error("Error processing file " + file.getName() + ": " + e.message);
    }
  }
}