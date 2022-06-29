function exportStories() {

  // Extract story headings:
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var searchType = DocumentApp.ElementType.PARAGRAPH;
  var searchHeading = DocumentApp.ParagraphHeading.HEADING3;
  var searchResult = null;
  
  var stories = []
  while (searchResult = body.findElement(searchType, searchResult)) {
    var par = searchResult.getElement().asParagraph();
    if (par.getHeading() == searchHeading) {
      stories.push([par.getText()])
    }
  }

  // Export to Backlog Sheet
  var pr = PropertiesService.getDocumentProperties();
  var ss = SpreadsheetApp.openById(pr.getProperty('BacklogSheetID'));
  var tab = ss.getSheetByName("Backlog Export");

  var startRow = 2
  var numRows = tab.getLastRow() - startRow + 1;
  var range = tab.getRange(startRow, 1, numRows);
  range.clear();
  tab.getRange(startRow, 1, stories.length).setValues(stories)
}

function connectSpreadsheet(){
  var pr = PropertiesService.getDocumentProperties();
  var ui = DocumentApp.getUi();
  var response = ui.prompt('Backlog Sheet ID', ui.ButtonSet.OK);
  pr.setProperty('BacklogSheetID', response.getResponseText());
}

function openSpreadsheet(){
  var pr = PropertiesService.getDocumentProperties();
  var ss = SpreadsheetApp.openById(pr.getProperty('BacklogSheetID'));
  var url = ss.getUrl();
  openURL(url);
}

function openURL(url){
  var html = "<script>window.open('" + url + "');google.script.host.close();</script>";
  var output = HtmlService.createHtmlOutput(html).setHeight(10).setWidth(80);
  var ui = DocumentApp.getUi();
  ui.showModalDialog(output, 'Opening...');
}

function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createMenu('Backlog')
      .addItem('Export Stories', 'exportStories')
      .addItem('Connect Spreadsheet', 'connectSpreadsheet')
      .addItem('Open Spreadsheet', 'openSpreadsheet')
      .addToUi();
}

function onInstall(){
  onOpen();
}