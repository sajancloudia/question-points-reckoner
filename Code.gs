/**
 * Creates a custom menu in the spreadsheet to launch the viewer dialog.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Question Viewer')
    .addItem('Open Question Reckoner', 'openViewerDialog')
    .addToUi();
}

/**
 * Displays the HTML modal dialog inside Google Sheets.
 */
function openViewerDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Index')
    .setWidth(950)
    .setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, 'Question Points Reckoner');
}

/**
 * Web App entry point (if you choose to deploy as a Web App).
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Question Points Reckoner')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Fetches question numbers and Google Drive links from the active sheet.
 */
function getSheetData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1') 
                || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return [];

  // Read columns A (Questions), B (English), C (Malayalam) starting from row 2
  const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

  return values.map(row => {
    return {
      question: String(row[0]).trim(),
      englishUrl: convertToPreviewUrl(String(row[1]).trim()),
      malayalamUrl: convertToPreviewUrl(String(row[2]).trim())
    };
  }).filter(item => item.question !== '');
}

/**
 * Converts a standard Google Drive sharing link to a preview link for iframe embedding.
 */
function convertToPreviewUrl(url) {
  if (!url) return '';
  const match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match && match[1]) {
    return `https://drive.google.com/file/d/${match[1]}/preview`;
  }
  return url;
}
