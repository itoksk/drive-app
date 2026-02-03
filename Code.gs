/**
 * Google Driveの特定のフォルダ内のファイルとフォルダ情報を取得し、各フォルダ名のシートに出力します
 * 新しいファイルやフォルダが追加された場合には、メールで通知されます
 *
 * 【改善版】共有ドライブ対応、URL処理改善、エラーハンドリング強化、シート名の重複防止
 *
 * 使い方:
 * 1. スプレッドシートを開くとカスタムメニュー「ドライブ監視」が表示されます
 * 2. 「初期セットアップ」を実行するとmainシートが自動作成されます
 * 3. mainシートのB列に監視したいフォルダのURLまたはIDを入力します
 * 4. 「今すぐチェック実行」でテスト、「自動実行を設定」で定期実行を設定します
 *
 * 注意点:
 * - Google App Scriptのプロジェクト名は任意です
 * - フォルダの階層が変わると影響してしまうため、可能であればフォルダ構成は変更しないでください
 * - フォルダに新しいファイルやフォルダが追加されずに更新された場合、メール送信は行われません
 * - 共有ドライブのフォルダを監視する場合は、Drive APIが有効になっている必要があります
 */

// ============================================
// カスタムメニューと初期セットアップ
// ============================================

/**
 * スプレッドシートを開いたときにカスタムメニューを追加
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('ドライブ監視')
    .addItem('初期セットアップ', 'setupSheet')
    .addSeparator()
    .addItem('今すぐチェック実行', 'checkNewFilesWithNotification')
    .addSeparator()
    .addSubMenu(ui.createMenu('自動実行設定')
      .addItem('1時間ごとに自動実行', 'setupHourlyTrigger')
      .addItem('6時間ごとに自動実行', 'setupSixHourlyTrigger')
      .addItem('毎日自動実行', 'setupDailyTrigger')
      .addItem('自動実行を停止', 'removeTriggersWithNotification'))
    .addSeparator()
    .addItem('サンプル行を追加', 'addSampleRow')
    .addItem('使い方を表示', 'showHelp')
    .addToUi();
}

/**
 * 初期セットアップ - mainシートとヘッダーを自動作成
 */
function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName('main');

  // mainシートがなければ作成
  if (!mainSheet) {
    // 最初のシートをmainにリネームするか、新規作成
    var sheets = ss.getSheets();
    if (sheets.length === 1 && sheets[0].getLastRow() === 0) {
      // 空のシートが1つだけならリネーム
      mainSheet = sheets[0];
      mainSheet.setName('main');
    } else {
      // 新規作成
      mainSheet = ss.insertSheet('main', 0);
    }
  }

  // ヘッダー行を設定
  var headers = ['有効化', 'フォルダURL/ID', 'フォルダ名', '最終更新日時', 'オーナー', 'メール送信先', 'エラー'];
  mainSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダー行の書式設定
  var headerRange = mainSheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');

  // 列幅を調整
  mainSheet.setColumnWidth(1, 80);   // 有効化
  mainSheet.setColumnWidth(2, 300);  // フォルダURL/ID
  mainSheet.setColumnWidth(3, 150);  // フォルダ名
  mainSheet.setColumnWidth(4, 150);  // 最終更新日時
  mainSheet.setColumnWidth(5, 200);  // オーナー
  mainSheet.setColumnWidth(6, 200);  // メール送信先
  mainSheet.setColumnWidth(7, 250);  // エラー

  // A列にデータ入力規則（チェックボックス）を設定
  var checkboxRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .build();
  mainSheet.getRange(2, 1, 100, 1).setDataValidation(checkboxRule);

  // 1行目を固定
  mainSheet.setFrozenRows(1);

  // mainシートをアクティブにする
  ss.setActiveSheet(mainSheet);

  SpreadsheetApp.getUi().alert(
    'セットアップ完了',
    'mainシートを作成しました。\n\n' +
    '【次のステップ】\n' +
    '1. A列のチェックボックスをONにする\n' +
    '2. B列に監視したいフォルダのURLを貼り付け\n' +
    '3. F列に通知先メールアドレスを入力\n' +
    '4. メニューから「今すぐチェック実行」を選択',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * サンプル行を追加
 */
function addSampleRow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName('main');

  if (!mainSheet) {
    SpreadsheetApp.getUi().alert('エラー', 'mainシートが見つかりません。\n先に「初期セットアップ」を実行してください。', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // 次の空き行を探す
  var lastRow = mainSheet.getLastRow();
  var newRow = lastRow + 1;

  // サンプルデータを追加
  mainSheet.getRange(newRow, 1).setValue(false);
  mainSheet.getRange(newRow, 2).setValue('ここにフォルダURLまたはIDを貼り付け');
  mainSheet.getRange(newRow, 6).setValue('your-email@example.com');

  // チェックボックスを設定
  var checkboxRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .build();
  mainSheet.getRange(newRow, 1).setDataValidation(checkboxRule);

  SpreadsheetApp.getUi().alert(
    'サンプル行を追加しました',
    '行 ' + newRow + ' にサンプルを追加しました。\n\n' +
    '【設定方法】\n' +
    '1. B列にGoogle ドライブのフォルダURLを貼り付け\n' +
    '2. F列に通知先メールアドレスを入力\n' +
    '3. A列のチェックボックスをONにする',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 使い方を表示
 */
function showHelp() {
  var helpText =
    '【ドライブ監視システムの使い方】\n\n' +
    '■ 初期設定\n' +
    '1. 「初期セットアップ」を実行\n' +
    '2. B列に監視したいフォルダのURLを貼り付け\n' +
    '3. F列に通知先メールアドレスを入力\n' +
    '4. A列のチェックボックスをONにする\n\n' +
    '■ 動作確認\n' +
    '「今すぐチェック実行」で手動テスト\n\n' +
    '■ 自動実行\n' +
    '「自動実行設定」から実行間隔を選択\n\n' +
    '■ 共有ドライブを使う場合\n' +
    'Drive APIの有効化が必要です\n' +
    '（サービス→Drive API を追加）\n\n' +
    '■ 詳細な手順書\n' +
    'https://github.com/itoksk/drive-app';

  SpreadsheetApp.getUi().alert('使い方', helpText, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * チェック実行（通知付き）
 */
function checkNewFilesWithNotification() {
  var ui = SpreadsheetApp.getUi();

  try {
    checkNewFiles();
    ui.alert(
      'チェック完了',
      'フォルダのチェックが完了しました。\n\n' +
      '結果はmainシートと各フォルダのシートを確認してください。\n' +
      'エラーがある場合はG列に表示されます。',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert(
      'エラーが発生しました',
      'チェック中にエラーが発生しました。\n\n' +
      'エラー内容: ' + e.toString() + '\n\n' +
      '【確認事項】\n' +
      '・mainシートが存在するか\n' +
      '・フォルダURLが正しいか\n' +
      '・フォルダへのアクセス権限があるか',
      ui.ButtonSet.OK
    );
  }
}

/**
 * 1時間ごとのトリガーを設定
 */
function setupHourlyTrigger() {
  removeTriggers();
  ScriptApp.newTrigger('checkNewFiles')
    .timeBased()
    .everyHours(1)
    .create();
  SpreadsheetApp.getUi().alert('自動実行設定完了', '1時間ごとに自動実行するように設定しました。', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 6時間ごとのトリガーを設定
 */
function setupSixHourlyTrigger() {
  removeTriggers();
  ScriptApp.newTrigger('checkNewFiles')
    .timeBased()
    .everyHours(6)
    .create();
  SpreadsheetApp.getUi().alert('自動実行設定完了', '6時間ごとに自動実行するように設定しました。', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 毎日のトリガーを設定
 */
function setupDailyTrigger() {
  removeTriggers();
  ScriptApp.newTrigger('checkNewFiles')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
  SpreadsheetApp.getUi().alert('自動実行設定完了', '毎日9時に自動実行するように設定しました。', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * すべてのトリガーを削除
 */
function removeTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkNewFiles') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * トリガー削除（UIから呼び出し用）
 */
function removeTriggersWithNotification() {
  removeTriggers();
  SpreadsheetApp.getUi().alert('自動実行を停止しました', 'すべての自動実行設定を削除しました。', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * スプレッドシート名の末尾にV2を追加
 */
function renameToV2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentName = ss.getName();
  if (!currentName.endsWith('V2')) {
    ss.rename(currentName + ' V2');
    SpreadsheetApp.getUi().alert('リネーム完了', 'スプレッドシート名を「' + currentName + ' V2」に変更しました。', SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert('確認', '既にV2が付いています。', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ============================================
// メイン処理
// ============================================

// Googleドライブの新規ファイルやフォルダの変更をチェックする関数
function checkNewFiles() {
  try {
    // 現在のスプレッドシートを取得します
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName('main');
    
    if (!mainSheet) {
      Logger.log('エラー: mainシートが見つかりません');
      return;
    }
    
    var data = mainSheet.getDataRange().getValues();
    
    // ヘッダー行をスキップ（1行目はヘッダーと仮定）
    var startRow = data.length > 1 && isHeaderRow(data[0]) ? 1 : 0;

    // データの行数分ループします
    for (var i = startRow; i < data.length; i++) {
      var row = data[i];
      // 通知フラグを取得します
      var notifyFlag = row[0];

      // 通知フラグがtrueの場合、以下の処理を行います
      if (notifyFlag === true || notifyFlag === 'TRUE' || notifyFlag === 'true') {
        try {
          // フォルダURLまたはフォルダIDを取得します
          var folderUrlOrId = row[1];
          
          if (!folderUrlOrId || folderUrlOrId === '') {
            Logger.log('行 ' + (i + 1) + ': フォルダURL/IDが空です');
            continue;
          }
          
          // フォルダIDを取得します（URLの場合は抽出、IDの場合はそのまま使用）
          var folderId = extractFolderId(folderUrlOrId);
          
          if (!folderId) {
            Logger.log('行 ' + (i + 1) + ': フォルダIDの抽出に失敗しました: ' + folderUrlOrId);
            updateErrorStatus(mainSheet, i + 1, 'フォルダIDの抽出に失敗しました');
            continue;
          }
          
          // フォルダ情報を取得します（共有ドライブ対応）
          var folderMeta = getFolderMetadata(folderId);
          
          if (!folderMeta) {
            Logger.log('行 ' + (i + 1) + ': フォルダIDが無効です: ' + folderId);
            updateErrorStatus(mainSheet, i + 1, 'フォルダIDが無効です: ' + folderId);
            continue;
          }
          
          if (folderMeta.errorMessage) {
            Logger.log('行 ' + (i + 1) + ': ' + folderMeta.errorMessage);
            updateErrorStatus(mainSheet, i + 1, folderMeta.errorMessage);
            continue;
          }
          
          // フォルダ名を取得します
          var folderName = folderMeta.name;

          // フォルダ名、最終更新日、オーナーを更新します (列 3, 4, 5)
          try {
            mainSheet.getRange(i + 1, 3, 1, 3).setValues([
              [folderName, new Date(), folderMeta.ownerEmail || '取得不可']
            ]);
          } catch (e) {
            Logger.log('行 ' + (i + 1) + ': フォルダ情報の更新でエラー: ' + e.toString());
            mainSheet.getRange(i + 1, 3, 1, 3).setValues([
              [folderName, new Date(), '取得不可']
            ]);
          }

          // 新しいフォルダの内容を出力し、新しいURLを取得します
          var sheet = getOrCreateFolderSheet(ss, folderId, folderName);
          if (!sheet) {
            Logger.log('行 ' + (i + 1) + ': シートの作成に失敗しました: ' + folderName);
            updateErrorStatus(mainSheet, i + 1, 'シートの作成に失敗しました');
            continue;
          }

          // 既存のURLを取得します
          var existingUrls = getExistingUrls(sheet);
          
          // フォルダ内容を出力し、新規のファイルとフォルダを取得します
          var newFilesAndFolders = outputFolderContents(
            folderId,
            folderName,
            sheet,
            existingUrls,
            folderMeta.useDriveApiOnly,
            folderMeta.driveId
          );

          // 新しいファイルやフォルダがある場合、メールを送信します
          if (newFilesAndFolders.length > 0) {
            var emailRecipient = row[5];
            if (emailRecipient && emailRecipient !== '') {
              var emailBody = '以下のファイルまたはフォルダが新たに追加されました：\n\n';
              for (var j = 0; j < newFilesAndFolders.length; j++) {
                var fileData = newFilesAndFolders[j];
                emailBody += '名前: ' + fileData.name + '\n';
                emailBody += 'URL: ' + fileData.url + '\n';
                emailBody += '種類: ' + fileData.type + '\n';
                emailBody += '最終更新日時: ' + fileData.lastUpdated + '\n';
                emailBody += 'オーナー: ' + fileData.owner + '\n';
                emailBody += 'フォルダ構成: ' + fileData.ancestry + '\n\n';
              }

              try {
                MailApp.sendEmail(emailRecipient, '新しいファイルまたはフォルダが追加されました: ' + folderName, emailBody);
                Logger.log('行 ' + (i + 1) + ': メール送信成功 - ' + emailRecipient);
              } catch (e) {
                Logger.log('行 ' + (i + 1) + ': メール送信エラー: ' + e.toString());
              }
            } else {
              Logger.log('行 ' + (i + 1) + ': メール送信先が設定されていません');
            }
          }
          
          // エラー状態をクリア
          clearErrorStatus(mainSheet, i + 1);
          
        } catch (e) {
          Logger.log('行 ' + (i + 1) + ': 処理中にエラーが発生しました: ' + e.toString());
          updateErrorStatus(mainSheet, i + 1, 'エラー: ' + e.toString());
        }
      }
    }
  } catch (e) {
    Logger.log('checkNewFiles関数でエラーが発生しました: ' + e.toString());
    throw e;
  }
}

// ヘッダー行かどうかを判定する関数
function isHeaderRow(row) {
  // 最初の列がチェックボックス的な値でない場合、ヘッダー行と判断
  return row[0] !== true && row[0] !== false && row[0] !== 'TRUE' && row[0] !== 'FALSE';
}

// フォルダ情報を取得する関数（DriveApp優先、共有ドライブはDrive APIで対応）
function getFolderMetadata(folderId) {
  try {
    var folder = DriveApp.getFolderById(folderId);
    return {
      name: folder.getName(),
      ownerEmail: getOwnerEmail(folder),
      useDriveApiOnly: false,
      driveId: null
    };
  } catch (e) {
    Logger.log('DriveAppで取得失敗、Drive APIを試行: ' + e.toString());
  }

  var fileData = fetchDriveFileMetadata(folderId, 'id,name,mimeType,owners,driveId');
  if (!fileData) {
    return null;
  }
  if (fileData.mimeType !== 'application/vnd.google-apps.folder') {
    return { errorMessage: '指定IDはフォルダではありません: ' + folderId };
  }

  return {
    name: fileData.name || '（名称不明）',
    ownerEmail: extractOwnerEmailFromApi(fileData),
    useDriveApiOnly: true,
    driveId: fileData.driveId || null
  };
}

// Drive APIでファイル/フォルダのメタデータを取得する関数（共有ドライブ対応）
function fetchDriveFileMetadata(fileId, fields) {
  try {
    var url = 'https://www.googleapis.com/drive/v3/files/' + fileId +
      '?fields=' + encodeURIComponent(fields) +
      '&supportsAllDrives=true';
    var response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      return JSON.parse(response.getContentText());
    }

    Logger.log('Drive API files.get エラー: ' + response.getResponseCode() + ' ' + response.getContentText());
  } catch (e) {
    Logger.log('Drive API files.get 失敗: ' + e.toString());
  }
  return null;
}

// Drive APIレスポンスからオーナーのメールアドレスを取得
function extractOwnerEmailFromApi(fileData) {
  if (fileData && fileData.owners && fileData.owners.length > 0) {
    return fileData.owners[0].emailAddress || '取得不可';
  }
  return '取得不可';
}

// エラー状態を更新する関数
function updateErrorStatus(sheet, row, errorMessage) {
  try {
    // G列にエラーメッセージを記録（存在しない場合は作成）
    var lastCol = sheet.getLastColumn();
    if (lastCol < 7) {
      sheet.getRange(1, 7).setValue('エラー');
    }
    sheet.getRange(row, 7).setValue(errorMessage);
  } catch (e) {
    Logger.log('エラー状態の更新に失敗: ' + e.toString());
  }
}

// エラー状態をクリアする関数
function clearErrorStatus(sheet, row) {
  try {
    var lastCol = sheet.getLastColumn();
    if (lastCol >= 7) {
      sheet.getRange(row, 7).setValue('');
    }
  } catch (e) {
    // エラーは無視
  }
}

// フォルダIDまたはURLからフォルダIDを抽出する関数（改善版）
function extractFolderId(urlOrId) {
  if (!urlOrId) {
    return null;
  }
  
  // 既にフォルダIDの形式（英数字とハイフン、アンダースコアのみ）の場合はそのまま返す
  if (/^[a-zA-Z0-9_-]+$/.test(urlOrId.trim())) {
    return urlOrId.trim();
  }
  
  // URLからフォルダIDを抽出
  var url = urlOrId.trim();
  
  // 複数のURLパターンに対応
  var patterns = [
    /\/folders\/([a-zA-Z0-9_-]+)/,  // 標準パターン: /folders/ID
    /id=([a-zA-Z0-9_-]+)/,          // id=パラメータパターン
    /#folders\/([a-zA-Z0-9_-]+)/    // ハッシュパターン: #folders/ID
  ];
  
  for (var i = 0; i < patterns.length; i++) {
    var match = url.match(patterns[i]);
    if (match && match[1]) {
      return match[1];
    }
  }
  
  return null;
}

// Drive APIを使用してフォルダ内のファイル一覧を取得する関数（共有ドライブ用）
function getFilesFromDriveAPI(folderId, pageToken, driveId) {
  try {
    var query = "'" + folderId + "' in parents and trashed=false";
    var url = 'https://www.googleapis.com/drive/v3/files' +
      '?q=' + encodeURIComponent(query) +
      '&fields=nextPageToken,files(id,name,mimeType,modifiedTime,webViewLink,owners)' +
      '&pageSize=1000' +
      '&supportsAllDrives=true' +
      '&includeItemsFromAllDrives=true' +
      '&spaces=drive';

    if (driveId) {
      url += '&corpora=drive&driveId=' + encodeURIComponent(driveId);
    }
    
    if (pageToken) {
      url += '&pageToken=' + encodeURIComponent(pageToken);
    }
    
    var response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      return JSON.parse(response.getContentText());
    } else {
      Logger.log('Drive APIリクエストエラー: ' + response.getResponseCode());
      Logger.log(response.getContentText());
      return null;
    }
  } catch (e) {
    Logger.log('Drive APIでのファイル取得エラー: ' + e.toString());
    return null;
  }
}

// オーナーのメールアドレスを取得する関数（エラーハンドリング強化）
function getOwnerEmail(folder) {
  try {
    var owner = folder.getOwner();
    if (owner) {
      return owner.getEmail();
    }
  } catch (e) {
    Logger.log('オーナー情報の取得に失敗: ' + e.toString());
  }
  
  try {
    // 代替方法: Drive APIを使用（共有ドライブ対応）
    var folderId = folder.getId();
    var fileData = fetchDriveFileMetadata(folderId, 'owners');
    return extractOwnerEmailFromApi(fileData);
  } catch (e) {
    Logger.log('Drive APIでのオーナー取得も失敗: ' + e.toString());
  }
  
  return '取得不可';
}

// シートに存在するURLを取得する関数
function getExistingUrls(sheet) {
  // シートが存在しない、またはシートの行数が2未満の場合、空の配列を返します
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  try {
    var urls = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
    return urls.flat().filter(function(url) {
      return url !== null && url !== '';
    });
  } catch (e) {
    Logger.log('既存URLの取得でエラー: ' + e.toString());
    return [];
  }
}

// フォルダの内容を出力し、新しいファイルとフォルダをフィルタリングする関数
function outputFolderContents(folderId, folderName, sheet, existingUrls, useDriveApiOnly, driveId) {
  try {
    var folder;
    var resolvedFolderName = folderName;
    
    // DriveAppで取得を試行（使用可能な場合のみ）
    if (!useDriveApiOnly) {
      try {
        folder = DriveApp.getFolderById(folderId);
        resolvedFolderName = folder.getName();
      } catch (e) {
        Logger.log('DriveAppで取得失敗、Drive APIを使用: ' + e.toString());
      }
    }

    if (!folder) {
      // Drive APIを使用した処理に切り替え
      return outputFolderContentsWithDriveAPI(folderId, resolvedFolderName, sheet, existingUrls, driveId);
    }

    // フォルダの階層を取得します
    var ancestry = resolvedFolderName;

    var fileDataList = [];
    outputFolderInfo(folder, ancestry, fileDataList);

    // シートの内容をクリアし、すべてのファイルデータを一度に書き込みます
    sheet.clearContents();
    sheet.appendRow(['名前', 'URL', '種類', '最終更新日時', 'オーナー', 'フォルダ構成']);
    
    // バッチ書き込みでパフォーマンス向上
    if (fileDataList.length > 0) {
      var rows = fileDataList.map(function(fileData) {
        return [fileData.name, fileData.url, fileData.type, fileData.lastUpdated, fileData.owner, fileData.ancestry];
      });
      sheet.getRange(2, 1, rows.length, 6).setValues(rows);
    }

    // 新しいファイルとフォルダをフィルタリングします
    var newFilesAndFolders = fileDataList.filter(function (fileData) {
      return !existingUrls.includes(fileData.url);
    });

    return newFilesAndFolders;
  } catch (e) {
    Logger.log('outputFolderContentsでエラー: ' + e.toString());
    throw e;
  }
}

// Drive APIを使用してフォルダの内容を出力する関数（共有ドライブ用）
function outputFolderContentsWithDriveAPI(folderId, folderName, sheet, existingUrls, driveId) {
  try {
    var fileDataList = [];
    var ancestry = folderName;
    
    // Drive APIを使用してファイル一覧を取得
    outputFolderInfoWithDriveAPI(folderId, ancestry, fileDataList, driveId);

    // シートの内容をクリアし、すべてのファイルデータを一度に書き込みます
    sheet.clearContents();
    sheet.appendRow(['名前', 'URL', '種類', '最終更新日時', 'オーナー', 'フォルダ構成']);
    
    // バッチ書き込みでパフォーマンス向上
    if (fileDataList.length > 0) {
      var rows = fileDataList.map(function(fileData) {
        return [fileData.name, fileData.url, fileData.type, fileData.lastUpdated, fileData.owner, fileData.ancestry];
      });
      sheet.getRange(2, 1, rows.length, 6).setValues(rows);
    }

    // 新しいファイルとフォルダをフィルタリングします
    var newFilesAndFolders = fileDataList.filter(function (fileData) {
      return !existingUrls.includes(fileData.url);
    });

    return newFilesAndFolders;
  } catch (e) {
    Logger.log('outputFolderContentsWithDriveAPIでエラー: ' + e.toString());
    throw e;
  }
}

// Drive APIを使用してフォルダ情報を出力する関数（共有ドライブ用）
function outputFolderInfoWithDriveAPI(folderId, ancestry, fileDataList, driveId) {
  try {
    var pageToken = null;
    
    do {
      var result = getFilesFromDriveAPI(folderId, pageToken, driveId);
      
      if (!result || !result.files) {
        break;
      }
      
      for (var i = 0; i < result.files.length; i++) {
        var file = result.files[i];
        var type = file.mimeType === 'application/vnd.google-apps.folder' ? 'フォルダ' : 'ファイル';
        var fallbackUrl = type === 'フォルダ'
          ? 'https://drive.google.com/drive/folders/' + file.id
          : 'https://drive.google.com/file/d/' + file.id + '/view';
        var fileData = {
          name: file.name,
          url: file.webViewLink || fallbackUrl,
          type: type,
          lastUpdated: file.modifiedTime ? new Date(file.modifiedTime) : new Date(),
          owner: file.owners && file.owners.length > 0 ? file.owners[0].emailAddress : '取得不可',
          ancestry: ancestry
        };
        fileDataList.push(fileData);
        
        // サブフォルダの場合は再帰的に処理
        if (type === 'フォルダ') {
          var subFolderAncestry = ancestry + ' > ' + file.name;
          outputFolderInfoWithDriveAPI(file.id, subFolderAncestry, fileDataList, driveId);
        }
      }
      
      pageToken = result.nextPageToken;
    } while (pageToken);
    
  } catch (e) {
    Logger.log('outputFolderInfoWithDriveAPIでエラー: ' + e.toString());
    throw e;
  }
}

// フォルダの情報を出力し、その内容をリストに追加する関数
function outputFolderInfo(folder, ancestry, fileDataList) {
  try {
    var files = folder.getFiles();
    var childFolders = folder.getFolders();

    // フォルダ内のファイルをリストに追加します
    while (files.hasNext()) {
      try {
        var file = files.next();
        var fileData = getFileData(file, 'ファイル', ancestry);
        fileDataList.push(fileData);
      } catch (e) {
        Logger.log('ファイル情報の取得でエラー: ' + e.toString());
        // エラーが発生しても処理を続行
      }
    }

    // フォルダ内のサブフォルダをリストに追加し、その内容を出力します
    while (childFolders.hasNext()) {
      try {
        var subFolder = childFolders.next();
        var subFolderAncestry = ancestry + ' > ' + subFolder.getName();
        var subFolderData = getFileData(subFolder, 'フォルダ', subFolderAncestry);
        fileDataList.push(subFolderData);
        outputFolderInfo(subFolder, subFolderAncestry, fileDataList);
      } catch (e) {
        Logger.log('サブフォルダ情報の取得でエラー: ' + e.toString());
        // エラーが発生しても処理を続行
      }
    }
  } catch (e) {
    Logger.log('outputFolderInfoでエラー: ' + e.toString());
    throw e;
  }
}

// ファイルまたはフォルダの情報を取得する関数
function getFileData(fileOrFolder, type, ancestry) {
  try {
    var fileData = {};
    fileData.name = fileOrFolder.getName();
    fileData.url = fileOrFolder.getUrl();
    fileData.type = type;
    fileData.lastUpdated = fileOrFolder.getLastUpdated();
    
    try {
      fileData.owner = fileOrFolder.getOwner().getEmail();
    } catch (e) {
      fileData.owner = '取得不可';
    }
    
    fileData.ancestry = ancestry;
    return fileData;
  } catch (e) {
    Logger.log('getFileDataでエラー: ' + e.toString());
    throw e;
  }
}

// フォルダIDに紐づくシートを取得または作成する関数
function getOrCreateFolderSheet(ss, folderId, folderName) {
  var sheet = resolveFolderSheet(ss, folderId, folderName);
  if (sheet) {
    markSheetWithFolderId(sheet, folderId);
    setFolderSheetMapping(folderId, sheet.getName());
    return sheet;
  }

  var baseName = sanitizeSheetName(folderName);
  var newSheetName = generateUniqueSheetName(ss, baseName, folderId);
  var newSheet = createSheetByName(ss, newSheetName);
  if (newSheet) {
    markSheetWithFolderId(newSheet, folderId);
    setFolderSheetMapping(folderId, newSheet.getName());
  }
  return newSheet;
}

// 既存シートの解決を行う
function resolveFolderSheet(ss, folderId, folderName) {
  var mappedName = getFolderSheetMapping(folderId);
  if (mappedName) {
    var mappedSheet = ss.getSheetByName(mappedName);
    if (mappedSheet) {
      return mappedSheet;
    }
  }

  var sheetByNote = findSheetByFolderId(ss, folderId);
  if (sheetByNote) {
    return sheetByNote;
  }

  var baseName = sanitizeSheetName(folderName);
  var candidates = findSheetCandidates(ss, baseName, folderName);
  if (candidates.length === 1 && !hasDifferentFolderIdNote(candidates[0], folderId)) {
    return candidates[0];
  }

  return null;
}

// シート名候補を探す
function findSheetCandidates(ss, baseName, rawName) {
  var sheets = ss.getSheets();
  var candidates = [];
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name === rawName || name === baseName || name.indexOf(baseName + '_') === 0) {
      candidates.push(sheets[i]);
    }
  }
  return candidates;
}

// フォルダIDからシートを検索（A1ノートを使用）
function findSheetByFolderId(ss, folderId) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var noteId = getFolderIdNote(sheets[i]);
    if (noteId === folderId) {
      return sheets[i];
    }
  }
  return null;
}

// シートにフォルダIDのノートを記録
function markSheetWithFolderId(sheet, folderId) {
  try {
    var currentId = getFolderIdNote(sheet);
    if (currentId !== folderId) {
      sheet.getRange(1, 1).setNote('folderId:' + folderId);
    }
  } catch (e) {
    Logger.log('シートへのフォルダID記録に失敗: ' + e.toString());
  }
}

// A1ノートからフォルダIDを取得
function getFolderIdNote(sheet) {
  try {
    var note = sheet.getRange(1, 1).getNote();
    if (note && note.indexOf('folderId:') === 0) {
      return note.substring('folderId:'.length);
    }
  } catch (e) {
    // 無視
  }
  return null;
}

// 別フォルダのノートが付いているか確認
function hasDifferentFolderIdNote(sheet, folderId) {
  var noteId = getFolderIdNote(sheet);
  return noteId && noteId !== folderId;
}

// シート名を安全な形に整形
function sanitizeSheetName(name) {
  var safeName = (name || '').toString().replace(/[\/\\\?\*\[\]:]/g, '_').trim();
  if (safeName === '') {
    safeName = 'folder';
  }
  if (safeName.length > 31) {
    safeName = safeName.substring(0, 31);
  }
  return safeName;
}

// シート名のマッピング（Document Properties）を取得/保存
function getFolderSheetMapping(folderId) {
  return PropertiesService.getDocumentProperties().getProperty('folderSheet_' + folderId);
}

function setFolderSheetMapping(folderId, sheetName) {
  PropertiesService.getDocumentProperties().setProperty('folderSheet_' + folderId, sheetName);
}

// 重複回避のためのシート名生成
function generateUniqueSheetName(ss, baseName, folderId) {
  var base = sanitizeSheetName(baseName);
  var shortId = (folderId || '').replace(/[^a-zA-Z0-9]/g, '').substring(0, 6);
  if (shortId === '') {
    shortId = 'id';
  }
  var suffix = '_' + shortId;
  var name = buildSheetNameWithSuffix(base, suffix);
  if (!ss.getSheetByName(name)) {
    return name;
  }

  var counter = 1;
  while (true) {
    var numericSuffix = '_' + counter;
    var candidate = buildSheetNameWithSuffix(base, numericSuffix);
    if (!ss.getSheetByName(candidate)) {
      return candidate;
    }
    counter++;
  }
}

function buildSheetNameWithSuffix(baseName, suffix) {
  var maxLength = 31 - suffix.length;
  var trimmed = baseName;
  if (trimmed.length > maxLength) {
    trimmed = trimmed.substring(0, maxLength);
  }
  return trimmed + suffix;
}

// 指定名で新しいシートを作成する関数
function createSheetByName(ss, sheetName) {
  try {
    var sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['名前', 'URL', '種類', '最終更新日時', 'オーナー', 'フォルダ構成']);
    
    Logger.log('シートを作成しました: ' + sheetName);
    return sheet;
  } catch (e) {
    Logger.log('シート作成でエラー: ' + e.toString());
    return null;
  }
}
