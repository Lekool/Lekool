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
  if (parsedData.length < 2) { // Must have header and at least one data row
    throw new Error('The CSV file appears to be empty or does not contain any transaction data.');
  }

  const headers = parsedData[0].map(h => h.trim());
  const dataRows = parsedData.slice(1);

  let processedData;

  // Detect file type and validate headers
  if (headers.includes('Transaction Date') && headers.includes('Post Date')) {
    // Chase format
    const requiredChaseHeaders = ['Transaction Date', 'Post Date', 'Description', 'Category', 'Type', 'Amount'];
    validateHeaders(headers, requiredChaseHeaders);
    processedData = processChaseData(dataRows, headers, cardName);
  } else if (headers.includes('Card Member')) {
    // Amex format
    const requiredAmexHeaders = ['Date', 'Description', 'Card Member', 'Amount'];
    validateHeaders(headers, requiredAmexHeaders);
    processedData = processAmexData(dataRows, headers, cardName);
  } else {
    throw new Error('Could not determine file type. The file does not seem to be a supported Chase or Amex statement. Please check the file headers.');
  }

  if (processedData.length > 0) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.getRange(sheet.getLastRow() + 1, 1, processedData.length, processedData[0].length)
         .setValues(processedData);
  } else {
    // This can happen if the file has a header but all data rows are invalid (e.g. summary rows)
    throw new Error('The file was processed, but no valid transaction rows were found.');
  }
}

/**
 * Validates that all required headers are present in the actual headers.
 * @param {string[]} actualHeaders - The headers from the parsed CSV file.
 * @param {string[]} requiredHeaders - The headers that are required for a specific format.
 */
function validateHeaders(actualHeaders, requiredHeaders) {
  const missingHeaders = [];
  for (const requiredHeader of requiredHeaders) {
    if (!actualHeaders.includes(requiredHeader)) {
      missingHeaders.push(requiredHeader);
    }
  }
  if (missingHeaders.length > 0) {
    throw new Error(`The uploaded file is missing the following required columns: ${missingHeaders.join(', ')}. Please check the file and try again.`);
  }
}

function processChaseData(rows, headers, cardName) {
  const transactionDateIndex = headers.indexOf('Transaction Date');
  const postDateIndex = headers.indexOf('Post Date');
  const descriptionIndex = headers.indexOf('Description');
  const categoryIndex = headers.indexOf('Category');
  const typeIndex = headers.indexOf('Type');
  const amountIndex = headers.indexOf('Amount');

  const processedRows = rows
    .filter(row => row.length > amountIndex && row[transactionDateIndex] && row[amountIndex])
    .map(row => {
      try {
        // Sanitize amount: remove $, ,, then parse
        const amountStr = row[amountIndex].replace(/[\$,]/g, '');
        const amount = parseFloat(amountStr);
        if (isNaN(amount)) {
          return null; // Skip rows where amount is not a number
        }

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
      } catch (e) {
        // Log error for the specific row and skip it
        console.error(`Skipping invalid Chase row: ${row}. Error: ${e.message}`);
        return null;
      }
    })
    .filter(row => row !== null); // Filter out the rows that were skipped

  return processedRows;
}

function processAmexData(rows, headers, cardName) {
  const dateIndex = headers.indexOf('Date');
  const descriptionIndex = headers.indexOf('Description');
  const cardMemberIndex = headers.indexOf('Card Member');
  const amountIndex = headers.indexOf('Amount');

  const processedRows = rows
    .filter(row => row.length > amountIndex && row[dateIndex] && row[amountIndex])
    .map(row => {
      try {
        // Sanitize amount: remove $, ,, then parse
        const amountStr = row[amountIndex].replace(/[\$,]/g, '');
        const amount = parseFloat(amountStr);
        if (isNaN(amount)) {
          return null; // Skip rows where amount is not a number
        }

        return [
          row[dateIndex],           // Transaction Date
          '',                       // Post Date
          row[descriptionIndex],    // Description
          '',                       // Category
          '',                       // Type
          amount,                   // Amount
          cardName,                 // Card
          row[cardMemberIndex]      // Notes
        ];
      } catch (e) {
        // Log error for the specific row and skip it
        console.error(`Skipping invalid Amex row: ${row}. Error: ${e.message}`);
        return null;
      }
    })
    .filter(row => row !== null); // Filter out the rows that were skipped

  return processedRows;
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
  return lines
    .filter(line => line.trim() !== '') // Filter out blank lines
    .map(line => {
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
