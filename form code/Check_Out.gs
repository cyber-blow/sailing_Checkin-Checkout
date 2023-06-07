// 定数を定義
var HEADER_ROW_INDEX = 0;
var OPTION_START_ROW_INDEX = 1;
var SPREADSHEET_ID = "14Gtg-jIt7QwIOjxj2LEDaTODT73icE7yjKc4iyL8U0Q"; // ここにスプレッドシートのIDを指定
var SHEET_NAME = "シート1"; // ここにシート名を指定
var PARENT_FOLDER_ID = "18g8JSrBxNAfnO75BYwNycT59ixlbdps-"; // ここに親フォルダのIDを指定

function createFormFromExcel() {
  var file = getFile(SPREADSHEET_ID);
  if (!file) {
    throw new Error("指定されたファイルが見つかりません。");
  }

  var spreadsheet = SpreadsheetApp.openById(file.getId());
  var sheet = getSheet(spreadsheet, SHEET_NAME);
  if (!sheet) {
    throw new Error("指定されたシートが見つかりません。");
  }

  var data = sheet.getDataRange().getValues();

  var responseSpreadsheet = createSpreadsheet("Check OUT Answer");
  var form = createForm("Check OUT", responseSpreadsheet);

  addQuestions(form, data);

  var folder = createAndMoveFiles(form, responseSpreadsheet, PARENT_FOLDER_ID);

  setSharingSettings(folder);
}

function getFile(spreadsheetId) {
  var files = DriveApp.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    if (file.getId() === spreadsheetId) {
      return file;
    }
  }
  return null;
}

function getSheet(spreadsheet, sheetName) {
  return spreadsheet.getSheetByName(sheetName);
}

function createSpreadsheet(name) {
  return SpreadsheetApp.create(name);
}

function createForm(name, spreadsheet) {
  return FormApp.create(name).setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());
}

function addQuestions(form, data) {
  var headerRow = data[HEADER_ROW_INDEX];
  for (var i = 0; i < headerRow.length; i++) {
    var question = headerRow[i];
    var options = data.slice(OPTION_START_ROW_INDEX).map(row => row[i]);
    // 重複する選択肢を取り除く
    var uniqueOptions = [...new Set(options)];
    var item = form.addListItem();
    item.setTitle(question);
    item.setChoiceValues(uniqueOptions);
  }
}

function createAndMoveFiles(form, spreadsheet, parentFolderId) {
  var dateFolder = createDateFolder(parentFolderId);
  moveFile(form, dateFolder);
  moveFile(spreadsheet, dateFolder);
  return dateFolder;
}

function createDateFolder(parentFolderId) {
  var date = new Date();
  var folderName = Utilities.formatDate(date, "GMT", "yyyy/MM/dd");
  var parentFolder = DriveApp.getFolderById(parentFolderId);
  return parentFolder.createFolder(folderName);
}

function moveFile(file, folder) {
  var fileInDrive = DriveApp.getFileById(file.getId());
  fileInDrive.moveTo(folder);
}

function setSharingSettings(folder) {
  folder.addEditor("example@example.com");
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
}
