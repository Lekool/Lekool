/**
 * @OnlyCurrentDoc
 *
 * The above comment directs App Script to limit the scope of file access for this script to the document
 * that it is bound to.
 */

/**
 * Creates a custom menu in the Google Sheet UI when the spreadsheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Finance Tool')
      .addItem('Setup Sheet', 'setupSheet')
      .addSeparator()
      .addItem('Load New File', 'loadNewFile')
      .addItem('Export Data', 'exportData')
      .addSeparator()
      .addItem('Clear Sheet', 'clearSheet')
      .addItem('Delete Current Row', 'deleteCurrentRow')
      .addToUi();
}

/**
 * Sets up the spreadsheet with the required headers and formatting.
 */
function setupSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Set headers
  const headers = [
    'Transaction Date', 'Post Date', 'Description', 'Category', 'Type', 'Amount', 'Card', 'Notes'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format date columns
  sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('B:B').setNumberFormat('yyyy-mm-dd');

  // Freeze the header row
  sheet.setFrozenRows(1);

  SpreadsheetApp.getUi().alert('Sheet setup complete!');
}

/**
 * Clears all data from the sheet, leaving the header row intact.
 */
function clearSheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Are you sure you want to clear all data?', ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const range = sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn());
    range.clearContent();
    ui.alert('Sheet cleared.');
  }
}

/**
 * Deletes the currently selected row.
 */
function deleteCurrentRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeRange = sheet.getActiveRange();
  if (activeRange.getRow() > 1) { // Do not delete header row
    sheet.deleteRow(activeRange.getRow());
  } else {
    SpreadsheetApp.getUi().alert('Cannot delete the header row.');
  }
}

/**
 * Shows an HTML dialog to allow the user to upload a CSV file.
 */
function loadNewFile() {
  const html = HtmlService.createHtmlOutputFromFile('FileUpload')
      .setWidth(300)
      .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(html, 'Load New Transactions');
}

/**
 * Processes the content of an uploaded CSV file.
 * This function is called from the client-side script in FileUpload.html.
 *
 * @param {string} csvContent The string content of the CSV file.
 * @param {string} cardName The name of the card provided by the user.
 */
function processUploadedCsv(csvContent, cardName) {
  const parsedData = parseCsv(csvContent);
  if (parsedData.length === 0) {
    throw new Error('CSV file is empty or could not be parsed.');
  }

  const headers = parsedData[0].map(h => h.trim());
  const dataRows = parsedData.slice(1);

  let processedData;

  // Detect file type based on headers
  if (headers.includes('Transaction Date') && headers.includes('Post Date')) {
    // Chase format
    processedData = processChaseData(dataRows, headers, cardName);
  } else if (headers.includes('Card Member')) {
    // Amex format
    processedData = processAmexData(dataRows, headers, cardName);
  } else {
    throw new Error('Could not determine file type. Please check the CSV headers.');
  }

  if (processedData.length > 0) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.getRange(sheet.getLastRow() + 1, 1, processedData.length, processedData[0].length)
         .setValues(processedData);
  }
}

function processChaseData(rows, headers, cardName) {
  const transactionDateIndex = headers.indexOf('Transaction Date');
  const postDateIndex = headers.indexOf('Post Date');
  const descriptionIndex = headers.indexOf('Description');
  const categoryIndex = headers.indexOf('Category');
  const typeIndex = headers.indexOf('Type');
  const amountIndex = headers.indexOf('Amount');

  return rows.map(row => {
    const amount = parseFloat(row[amountIndex]);
    return [
      row[transactionDateIndex], // Transaction Date
      row[postDateIndex],       // Post Date
      row[descriptionIndex],    // Description
      row[categoryIndex],       // Category
      row[typeIndex],           // Type
      -amount,                  // Amount (inverted)
      cardName,                 // Card
      ''                        // Notes
    ];
  });
}

function processAmexData(rows, headers, cardName) {
  const dateIndex = headers.indexOf('Date');
  const descriptionIndex = headers.indexOf('Description');
  const cardMemberIndex = headers.indexOf('Card Member');
  const amountIndex = headers.indexOf('Amount');

  return rows.map(row => {
    return [
      row[dateIndex],           // Transaction Date
      '',                       // Post Date
      row[descriptionIndex],    // Description
      '',                       // Category
      '',                       // Type
      parseFloat(row[amountIndex]), // Amount
      cardName,                 // Card
      row[cardMemberIndex]      // Notes
    ];
  });
}

/**
 * A robust CSV parser that handles quoted fields.
 * @param {string} csvContent The string content of the CSV file.
 * @returns {Array<Array<string>>} A 2D array of the parsed data.
 */
function parseCsv(csvContent) {
  const lines = csvContent.trim().split('\n');
  // This regex splits by commas, but ignores commas inside double quotes.
  const regex = /,(?=(?:(?:[^"]*"){2})*[^"]*$)/;
  return lines.map(line => {
    return line.split(regex).map(field => {
      // Remove leading/trailing whitespace and quotes from the field
      return field.trim().replace(/^"|"$/g, '').trim();
    });
  });
}

/**
 * Exports the current data to a new Google Sheet.
 */
function exportData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert('There is no data to export.');
    return;
  }

  const newSpreadsheet = SpreadsheetApp.create('Exported Transactions');
  const newSheet = newSpreadsheet.getSheets()[0];
  newSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  const url = newSpreadsheet.getUrl();
  const htmlOutput = HtmlService.createHtmlOutput(`<p>Data exported successfully. <a href="${url}" target="_blank">Open new sheet</a>.</p>`)
      .setWidth(300)
      .setHeight(80);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Export Complete');
}
