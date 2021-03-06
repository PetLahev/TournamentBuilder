function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('Custom Menu')
      .addItem('Show sidebar', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('HTML/Settings')
      .setTitle('Tournament builder')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function test() {
  Logger.log('Test');
  var file = SpreadsheetApp.getActiveSpreadsheet();
  var sheet =  file.getActiveSheet();
  var cell = sheet.getActiveCell();
  var builder = new StandardBracket(file, 8, 5, BracketsType.Standard);
  builder.removeBrackets(sheet);
  builder.build(sheet);
}