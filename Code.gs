/**
 * Google Driveの特定のフォルダ内のファイルとフォルダ情報を取得し、各フォルダ名のシートに出力します
 * 新しいファイルやフォルダが追加された場合には、メールで通知されます
 * 
 * 【改善版】共有ドライブ対応、URL処理改善、エラーハンドリング強化
 *
 * 使い方:
 * 1. ドライブ内の監視したいフォルダディレクトリのURLまたはフォルダIDをmainシートのB列に入力します
 * 2. mainシートのA列には、その行のフォルダ監視が有効かどうかを表すTRUE/FALSEを入力します
 * 3. mainシートのF列にはメール送信先のアドレスを入力します
 * 4. スクリプトのトリガーを設定して、定期的にcheckNewFiles()関数が実行されるように設定します
 *
 * 注意点:
 * - Google App Scriptのプロジェクト名は任意です 
 * - フォルダの階層が変わると影響してしまうため、可能であればフォルダ構成は変更しないでください
 * - フォルダに新しいファイルやフォルダが追加されずに更新された場合、メール送信は行われません
 * - 共有ドライブのフォルダを監視する場合は、Drive APIが有効になっている必要があります
 */

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
          
          // フォルダを取得します（共有ドライブ対応）
          var folder = getFolderById(folderId);
          
          if (!folder) {
            Logger.log('行 ' + (i + 1) + ': フォルダIDが無効です: ' + folderId);
            updateErrorStatus(mainSheet, i + 1, 'フォルダIDが無効です: ' + folderId);
            continue;
          }
          
          // フォルダ名を取得します
          var folderName = folder.getName();

          // フォルダ名、最終更新日、オーナーを更新します (列 3, 4, 5)
          try {
            var ownerEmail = getOwnerEmail(folder);
            mainSheet.getRange(i + 1, 3, 1, 3).setValues([
              [folderName, new Date(), ownerEmail]
            ]);
          } catch (e) {
            Logger.log('行 ' + (i + 1) + ': フォルダ情報の更新でエラー: ' + e.toString());
            mainSheet.getRange(i + 1, 3, 1, 3).setValues([
              [folderName, new Date(), '取得不可']
            ]);
          }

          // 新しいフォルダの内容を出力し、新しいURLを取得します
          var sheet = ss.getSheetByName(folderName);
          if (!sheet) {
            // 指定名のシートが存在しない場合、新しく作成します
            sheet = createSheet(ss, folderName);
            if (!sheet) {
              Logger.log('行 ' + (i + 1) + ': シートの作成に失敗しました: ' + folderName);
              updateErrorStatus(mainSheet, i + 1, 'シートの作成に失敗しました');
              continue;
            }
          }

          // 既存のURLを取得します
          var existingUrls = getExistingUrls(ss, folderName);
          
          // フォルダ内容を出力し、新規のファイルとフォルダを取得します
          var newFilesAndFolders = outputFolderContents(folderId, sheet, existingUrls);

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

// フォルダIDからフォルダを取得する関数（共有ドライブ対応）
function getFolderById(folderId) {
  try {
    // まず通常のDriveAppで試行（マイドライブ用）
    try {
      var folder = DriveApp.getFolderById(folderId);
      if (folder) {
        return folder;
      }
    } catch (e) {
      // DriveAppで取得できない場合は共有ドライブの可能性
      Logger.log('DriveAppで取得失敗、Drive APIを試行: ' + e.toString());
    }
    
    // Drive APIを使用して共有ドライブのフォルダを取得
    try {
      var driveApiUrl = 'https://www.googleapis.com/drive/v3/files/' + folderId + '?fields=id,name,mimeType,parents';
      var response = UrlFetchApp.fetch(driveApiUrl, {
        headers: {
          'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
        }
      });
      
      if (response.getResponseCode() === 200) {
        var fileData = JSON.parse(response.getContentText());
        if (fileData.mimeType === 'application/vnd.google-apps.folder') {
          // Drive APIで取得できた場合でも、DriveAppのオブジェクトが必要な場合は
          // 別の方法で処理する必要がある
          // ここでは簡易的にDriveAppで再試行
          try {
            return DriveApp.getFolderById(folderId);
          } catch (e2) {
            // DriveAppで取得できない場合は、Drive APIの結果を使って処理を続行
            // この場合、getFiles()やgetFolders()が使えないため、Drive APIで再帰的に取得する必要がある
            Logger.log('共有ドライブフォルダを検出: ' + fileData.name);
            // 共有ドライブの場合は、Drive APIを使用した処理に切り替える
            return createSharedDriveFolderWrapper(folderId, fileData);
          }
        }
      }
    } catch (e) {
      Logger.log('Drive APIでの取得も失敗: ' + e.toString());
    }
    
    return null;
  } catch (e) {
    Logger.log('getFolderByIdでエラー: ' + e.toString());
    return null;
  }
}

// 共有ドライブフォルダのラッパーオブジェクト
// Drive APIを使用してファイル一覧を取得する実装
function createSharedDriveFolderWrapper(folderId, fileData) {
  // Drive APIを使用してフォルダを取得する方法を試行
  // ただし、DriveAppのオブジェクトが必要なため、可能な限りDriveAppで取得を試みる
  try {
    // 再度DriveAppで試行（権限が付与された後は取得できる可能性がある）
    return DriveApp.getFolderById(folderId);
  } catch (e) {
    Logger.log('共有ドライブフォルダの取得に失敗: ' + e.toString());
    Logger.log('Drive APIを使用した処理が必要です。Drive APIが有効になっているか確認してください。');
    throw new Error('共有ドライブのフォルダにアクセスできません。Drive APIの有効化と適切な権限が必要です。エラー: ' + e.toString());
  }
}

// Drive APIを使用してフォルダ内のファイル一覧を取得する関数（共有ドライブ用）
function getFilesFromDriveAPI(folderId, pageToken) {
  try {
    var query = "'" + folderId + "' in parents and trashed=false";
    var url = 'https://www.googleapis.com/drive/v3/files?q=' + encodeURIComponent(query) + '&fields=nextPageToken,files(id,name,mimeType,modifiedTime,webViewLink,owners)';
    
    if (pageToken) {
      url += '&pageToken=' + encodeURIComponent(pageToken);
    }
    
    var response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      }
    });
    
    if (response.getResponseCode() === 200) {
      return JSON.parse(response.getContentText());
    } else {
      Logger.log('Drive APIリクエストエラー: ' + response.getResponseCode());
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
    // 代替方法: Drive APIを使用
    var folderId = folder.getId();
    var driveApiUrl = 'https://www.googleapis.com/drive/v3/files/' + folderId + '?fields=owners';
    var response = UrlFetchApp.fetch(driveApiUrl, {
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      }
    });
    
    if (response.getResponseCode() === 200) {
      var fileData = JSON.parse(response.getContentText());
      if (fileData.owners && fileData.owners.length > 0) {
        return fileData.owners[0].emailAddress;
      }
    }
  } catch (e) {
    Logger.log('Drive APIでのオーナー取得も失敗: ' + e.toString());
  }
  
  return '取得不可';
}

// シートに存在するURLを取得する関数
function getExistingUrls(ss, folderName) {
  var sheet = ss.getSheetByName(folderName);
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
function outputFolderContents(folderId, sheet, existingUrls) {
  try {
    var folder;
    var folderName;
    
    // まずDriveAppで取得を試行
    try {
      folder = DriveApp.getFolderById(folderId);
      folderName = folder.getName();
    } catch (e) {
      // DriveAppで取得できない場合はDrive APIを使用
      Logger.log('DriveAppで取得失敗、Drive APIを使用: ' + e.toString());
      try {
        var driveApiUrl = 'https://www.googleapis.com/drive/v3/files/' + folderId + '?fields=id,name,mimeType';
        var response = UrlFetchApp.fetch(driveApiUrl, {
          headers: {
            'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
          }
        });
        
        if (response.getResponseCode() === 200) {
          var fileData = JSON.parse(response.getContentText());
          folderName = fileData.name;
          // Drive APIを使用した処理に切り替え
          return outputFolderContentsWithDriveAPI(folderId, folderName, sheet, existingUrls);
        } else {
          throw new Error('フォルダを取得できませんでした: ' + folderId);
        }
      } catch (e2) {
        Logger.log('Drive APIでも取得失敗: ' + e2.toString());
        throw new Error('フォルダを取得できませんでした: ' + folderId + ' エラー: ' + e2.toString());
      }
    }
    
    if (!folder) {
      throw new Error('フォルダを取得できませんでした: ' + folderId);
    }

    // フォルダの階層を取得します
    var ancestry = folderName;

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
function outputFolderContentsWithDriveAPI(folderId, folderName, sheet, existingUrls) {
  try {
    var fileDataList = [];
    var ancestry = folderName;
    
    // Drive APIを使用してファイル一覧を取得
    outputFolderInfoWithDriveAPI(folderId, ancestry, fileDataList);

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
function outputFolderInfoWithDriveAPI(folderId, ancestry, fileDataList) {
  try {
    var pageToken = null;
    
    do {
      var result = getFilesFromDriveAPI(folderId, pageToken);
      
      if (!result || !result.files) {
        break;
      }
      
      for (var i = 0; i < result.files.length; i++) {
        var file = result.files[i];
        var type = file.mimeType === 'application/vnd.google-apps.folder' ? 'フォルダ' : 'ファイル';
        var fileData = {
          name: file.name,
          url: file.webViewLink || 'https://drive.google.com/drive/folders/' + file.id,
          type: type,
          lastUpdated: file.modifiedTime ? new Date(file.modifiedTime) : new Date(),
          owner: file.owners && file.owners.length > 0 ? file.owners[0].emailAddress : '取得不可',
          ancestry: ancestry
        };
        fileDataList.push(fileData);
        
        // サブフォルダの場合は再帰的に処理
        if (type === 'フォルダ') {
          var subFolderAncestry = ancestry + ' > ' + file.name;
          outputFolderInfoWithDriveAPI(file.id, subFolderAncestry, fileDataList);
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

// 指定名で新しいシートを作成する関数
function createSheet(ss, folderName) {
  try {
    // シート名に使用できない文字を置換
    var safeSheetName = folderName.replace(/[\/\\\?\*\[\]:]/g, '_');
    
    // シート名の長さ制限（31文字）に対応
    if (safeSheetName.length > 31) {
      safeSheetName = safeSheetName.substring(0, 31);
    }
    
    // 既に同名のシートが存在する場合は番号を付加
    var sheetName = safeSheetName;
    var counter = 1;
    while (ss.getSheetByName(sheetName)) {
      var suffix = '_' + counter;
      var maxLength = 31 - suffix.length;
      sheetName = safeSheetName.substring(0, maxLength) + suffix;
      counter++;
    }
    
    var sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['名前', 'URL', '種類', '最終更新日時', 'オーナー', 'フォルダ構成']);
    
    Logger.log('シートを作成しました: ' + sheetName);
    return sheet;
  } catch (e) {
    Logger.log('シート作成でエラー: ' + e.toString());
    return null;
  }
}
