function exportStories() {

  const doc = DocumentApp.getActiveDocument();
  const ui = DocumentApp.getUi()
  const body = doc.getBody();

   // Extract story headings:
  const searchType = DocumentApp.ElementType.PARAGRAPH;
  const searchHeading = DocumentApp.ParagraphHeading.HEADING3;
  let searchResult = null;
  const stories = []
  const storiesJson = {}
  const regexp = /(?<id>[A-Z]+-[0-9]+): (?<name>.*)/
  let paragraph, storyRaw, storyObj, storyMatch
  let errorCount = 0
  while (searchResult = body.findElement(searchType, searchResult)) {
    paragraph = searchResult.getElement().asParagraph();
    if (paragraph.getHeading() == searchHeading) {
      storyRaw = paragraph.getText()
      storyMatch = storyRaw.match(regexp)
      if (storyMatch) {
        storyObj = storyMatch.groups
        if (storyObj.id in storiesJson) {
          errorCount ++;
          Logger.log(Logger.log('ID "' + storyObj.id + '" for story "' + storyObj.name + '" is already in use.'))
        } else {
          storiesJson[storyObj.id] = storyObj
          stories.push([storyObj.id, storyObj.name])
        }
      } else {
        errorCount ++;
        Logger.log('Story "' + storyRaw + '" does not match REGEX pattern for stories.')
      }
    }
  }

  if (errorCount == 0) {
    // Export to Backlog Sheet
    const pr = PropertiesService.getDocumentProperties();
    const ss = SpreadsheetApp.openById(pr.getProperty('BacklogSheetID'));
    const tab = ss.getSheetByName("Backlog Export");
    const startRow = 2
    const numRows = tab.getLastRow() - startRow + 1;
    const range = tab.getRange(startRow, 1, numRows, 2);
    range.clear();
    tab.getRange(startRow, 1, stories.length, 2).setValues(stories)

    // Export to JSON
    const fileSets = {
      title: doc.getName() + '.json',
      mimeType: 'application/json'
    }
    const blob = Utilities.newBlob(JSON.stringify({"stories": storiesJson}, null, 2), "application/vnd.google-apps.script+json");
    const file = Drive.Files.insert(fileSets, blob)

    // Display summary
    ui.alert(stories.length + ' stories exported successfully.')

  } else {
    ui.alert('Unable to export (encountered ' + errorCount + ' errors). See execution log for details.')
  }

}

function connectSpreadsheet(){
  const pr = PropertiesService.getDocumentProperties();
  const ui = DocumentApp.getUi();
  const response = ui.prompt('Backlog Sheet ID', ui.ButtonSet.OK);
  pr.setProperty('BacklogSheetID', response.getResponseText());
}

function openSpreadsheet(){
  const pr = PropertiesService.getDocumentProperties();
  const ss = SpreadsheetApp.openById(pr.getProperty('BacklogSheetID'));
  const url = ss.getUrl();
  openURL(url);
}

function openURL(url){
  const html = "<script>window.open('" + url + "');google.script.host.close();</script>";
  const output = HtmlService.createHtmlOutput(html).setHeight(10).setWidth(80);
  const ui = DocumentApp.getUi();
  ui.showModalDialog(output, 'Opening...');
}

function onOpen() {
  const ui = DocumentApp.getUi();
  ui.createMenu('Backlog')
      .addItem('Export Stories', 'exportStories')
      .addItem('Connect Spreadsheet', 'connectSpreadsheet')
      .addItem('Open Spreadsheet', 'openSpreadsheet')
      .addToUi();
}
