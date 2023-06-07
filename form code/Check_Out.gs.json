function createFormFromExcel() {
    var spreadsheetId = "14Gtg-jIt7QwIOjxj2LEDaTODT73icE7yjKc4iyL8U0Q"; // ExcelファイルがアップロードされるGoogleスプレッドシートのIDを指定してください
    var sheetName = "シート1"; // Excelファイルがアップロードされるシート名を指定してください
    var parentFolderId = "18g8JSrBxNAfnO75BYwNycT59ixlbdps-"; // フォームと回答スプレッドシートを格納する親フォルダのIDを指定してください
  
    // ファイル選択ダイアログの表示
    var files = DriveApp.getFiles();
  
    // 選択されたファイルの情報を取得
    var file;
    while (files.hasNext()) {
      var currentFile = files.next();
      if (currentFile.getId() === spreadsheetId) {
        file = currentFile;
        break;
      }
    }
  
    if (!file) {
      Logger.log("指定されたファイルが見つかりません。");
      return;
    }
  
    // フォームの作成
    var responseSpreadsheet = SpreadsheetApp.create("Check OUT Answer");
    var form = FormApp.create("Check OUT").setDestination(FormApp.DestinationType.SPREADSHEET, responseSpreadsheet.getId());
  
    // Excelファイルのデータを取得し、フォームに追加
    var spreadsheet = SpreadsheetApp.openById(file.getId());
    var sheet = spreadsheet.getSheetByName(sheetName);
  
    if (!sheet) {
      Logger.log("指定されたシートが見つかりません。");
      return;
    }
  
    var data = sheet.getDataRange().getValues();
  
    // ヘッダー行を取得
    var headerRow = data[0];
  
    // プルダウン選択肢を格納する配列
    var choices = [];
  
    // ヘッダー行の要素ごとに選択肢を追加
    for (var i = 0; i < headerRow.length; i++) {
      var question = headerRow[i];
      var options = [];
  
      // 各行の対応する列から選択肢を取得
      for (var j = 1; j < data.length; j++) {
        options.push(data[j][i]);
      }
  
      // 選択肢を追加
      choices.push(options);
  
      // 質問と選択肢をフォームに追加
      var item = form.addListItem();
      item.setTitle(question);
      item.setChoiceValues(options);
    }
  
    // フォームのURLをログに出力
    Logger.log("フォームのURL: " + form.getPublishedUrl());
  
    // yyyy/mm/dd 形式のフォルダ名を作成
    var date = new Date();
    var folderName = Utilities.formatDate(date, "GMT", "yyyy/MM/dd");
  
    // 保存先の親フォルダを取得
    var parentFolder = DriveApp.getFolderById(parentFolderId);
  
    // 日付フォルダを作成
    var dateFolder = parentFolder.createFolder(folderName);
  
    // フォームファイルを保存
    var formFile = DriveApp.getFileById(form.getId());
    formFile.moveTo(dateFolder);
  
    // 回答スプレッドシートを保存
    var responseSpreadsheetId = responseSpreadsheet.getId();
    var responseSheetFile = DriveApp.getFileById(responseSpreadsheetId);
    responseSheetFile.moveTo(dateFolder);
  
    // フォームと回答スプレッドシートのURLをログに出力
    Logger.log("フォームのURL: " + form.getPublishedUrl());
    Logger.log("回答スプレッドシートのURL: " + responseSpreadsheet.getUrl());
  
    // 回答スプレッドシートの共有設定
    dateFolder.addEditor("example@example.com");
    dateFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }
  
  