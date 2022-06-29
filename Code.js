function exportStories() {

  var doc = DocumentApp.getActiveDocument();
  var ui = DocumentApp.getUi()
  var body = doc.getBody();

   // Extract story headings:
  var searchType = DocumentApp.ElementType.PARAGRAPH;
  var searchHeading = DocumentApp.ParagraphHeading.HEADING3;
  var searchResult = null;
  var stories = []
  var storiesJson = []
  var regexp = /(?<id>[A-Z]+-[0-9]+): (?<name>.*)/
  var storyRaw = null
  var storyObj = null
  var storyMatch = null
  var par = null
  var errorCount = 0
  while (searchResult = body.findElement(searchType, searchResult)) {
    par = searchResult.getElement().asParagraph();
    if (par.getHeading() == searchHeading) {
      storyRaw = par.getText()
      storyMatch = storyRaw.match(regexp)
      if (storyMatch) {
        storyObj = storyMatch.groups
        stories.push([storyRaw])
        storiesJson.push(storyObj)
      } else {
        errorCount ++;
        Logger.log('Story "' + storyRaw + '" does not match REGEX pattern for stories.')
      }
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

  // Export to JSON
  var fileSets = {
    title: doc.getName() + '.json',
    mimeType: 'application/json'
  }
  var blob = Utilities.newBlob(JSON.stringify(storiesJson), "application/vnd.google-apps.script+json");
  file = Drive.Files.insert(fileSets, blob)

  // Display summary
  var successMsg = stories.length + ' stories exported successfully.'
  var errorMsg = null
  if (errorCount > 0) {
    errorMsg = errorCount + ' errors. See execution log for details.'
  } else {
    errorMsg = errorCount + ' errors.'
  }
  ui.alert(successMsg + '\n' + errorMsg)
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
