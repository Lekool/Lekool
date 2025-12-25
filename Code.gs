/**
 * @OnlyCurrentDoc
 */

/**
 * Creates a custom menu in the Google Sheet UI when the spreadsheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Finance Tool')
      .addItem('Load New CSV File', 'loadNewFile')
      .addToUi();
}

/**
 * Shows an HTML dialog to allow the user to upload a CSV file.
 */
function loadNewFile() {
  const html = HtmlService.createHtmlOutputFromFile('FileUpload')
      .setWidth(450)
      .setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, 'Load New Transactions');
}

/**
 * Sets up a newly created sheet with headers and formatting.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The new sheet to set up.
 */
function setupNewSheet(sheet) {
  const headers = [
    'Transaction Date', 'Post Date', 'Description', 'Category', 'Type', 'Amount', 'Card', 'Notes'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange('A:B').setNumberFormat('MM/dd/yyyy');
  sheet.setFrozenRows(1);
}

/**
 * Processes the uploaded CSV file, creating a new sheet for its content.
 *
 * @param {string} csvContent The string content of the CSV file.
 * @param {string} cardName The name for the new sheet, provided by the user.
 */
function processUploadedCsv(csvContent, cardName) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Check if a sheet with this name already exists
    if (spreadsheet.getSheetByName(cardName)) {
      throw new Error(`A sheet named "${cardName}" already exists. Please choose a unique name.`);
    }

    // 1. Process the uploaded file content using native Utilities
    let parsedData;
    try {
      // Decode the Base64 string to a blob, then get as string
      const decodedBlob = Utilities.newBlob(Utilities.base64Decode(csvContent, Utilities.Charset.UTF_8));
      const decodedString = decodedBlob.getDataAsString();
      parsedData = Utilities.parseCsv(decodedString);
    } catch (e) {
      throw new Error('Failed to parse CSV file. It may be malformed or the encoding failed.');
    }

    if (!parsedData || parsedData.length < 2) {
      throw new Error('The CSV file appears to be empty or does not contain any transaction data.');
    }

    const headers = parsedData[0].map(h => h.trim());
    const dataRows = parsedData.slice(1);
    let processedRows;

    if (headers.includes('Transaction Date') && headers.includes('Post Date')) {
      const requiredChaseHeaders = ['Transaction Date', 'Post Date', 'Description', 'Category', 'Type', 'Amount'];
      validateHeaders(headers, requiredChaseHeaders);
      processedRows = processChaseData(dataRows, headers, cardName);
    } else if (headers.includes('Card Member')) {
      const requiredAmexHeaders = ['Date', 'Description', 'Card Member', 'Amount'];
      validateHeaders(headers, requiredAmexHeaders);
      processedRows = processAmexData(dataRows, headers, cardName);
    } else {
      throw new Error('Could not determine file type. The file does not seem to be a supported Chase or Amex statement.');
    }

    if (processedRows.length === 0) {
      throw new Error('The file was processed, but no valid transaction rows were found.');
    }

    // 2. Sort the processed data (Post Date at index 1, Transaction Date at index 0)
    processedRows.sort((a, b) => {
      const dateA = (a[1] instanceof Date) ? a[1] : a[0];
      const dateB = (b[1] instanceof Date) ? b[1] : b[0];
      return dateA.getTime() - dateB.getTime();
    });

    // 3. Create a new sheet and set it up
    const newSheet = spreadsheet.insertSheet(cardName);
    setupNewSheet(newSheet);

    // 4. Write the sorted data to the new sheet
    newSheet.getRange(2, 1, processedRows.length, processedRows[0].length).setValues(processedRows);

  } catch (err) {
    // Re-throw the error so it can be caught by the client-side failure handler
    // Ensure the message is clean
    throw new Error(err.message);
  }
}

// --- Helper Functions ---

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

/**
 * A universal date parser that tries multiple formats.
 * @param {string} dateString - The date string to parse.
 * @returns {Date|null} A Date object, or null if parsing fails.
 */
function parseDateUniversal(dateString) {
  if (!dateString || typeof dateString !== 'string') return null;
  let date = new Date(dateString);
  if (date && !isNaN(date.getTime())) return date;
  date = parseYyyyMmDd(dateString);
  if (date && !isNaN(date.getTime())) return date;
  return null;
}

/**
 * Parses a date string in yyyy/mm/dd format into a Date object.
 * @param {string} dateString - The date string to parse.
 * @returns {Date|null} A Date object, or null if the format is invalid.
 */
function parseYyyyMmDd(dateString) {
  if (!dateString || typeof dateString !== 'string') return null;
  const parts = dateString.split('/');
  if (parts.length !== 3) return null;
  const year = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10) - 1;
  const day = parseInt(parts[2], 10);
  if (isNaN(year) || isNaN(month) || isNaN(day)) return null;
  const date = new Date(year, month, day);
  if (date.getFullYear() !== year || date.getMonth() !== month || date.getDate() !== day) return null;
  return date;
}

function processChaseData(rows, headers, cardName) {
  const transactionDateIndex = headers.indexOf('Transaction Date');
  const postDateIndex = headers.indexOf('Post Date');
  const descriptionIndex = headers.indexOf('Description');
  const categoryIndex = headers.indexOf('Category');
  const typeIndex = headers.indexOf('Type');
  const amountIndex = headers.indexOf('Amount');

  return rows
    .filter(row => row.length > amountIndex && row[transactionDateIndex] && row[amountIndex])
    .map(row => {
      try {
        const amount = parseFloat(row[amountIndex].replace(/[\$,]/g, ''));
        if (isNaN(amount)) return null;
        const transactionDate = parseDateUniversal(row[transactionDateIndex]);
        const postDate = parseDateUniversal(row[postDateIndex]);
        if (!transactionDate || !postDate) return null;
        return [transactionDate, postDate, row[descriptionIndex], row[categoryIndex], row[typeIndex], amount, cardName, ''];
      } catch (e) { return null; }
    })
    .filter(row => row !== null);
}

function processAmexData(rows, headers, cardName) {
  const dateIndex = headers.indexOf('Date');
  const descriptionIndex = headers.indexOf('Description');
  const cardMemberIndex = headers.indexOf('Card Member');
  const amountIndex = headers.indexOf('Amount');

  return rows
    .filter(row => row.length > amountIndex && row[dateIndex] && row[amountIndex])
    .map(row => {
      try {
        const amount = parseFloat(row[amountIndex].replace(/[\$,]/g, ''));
        if (isNaN(amount)) return null;
        const transactionDate = parseDateUniversal(row[dateIndex]);
        if (!transactionDate) return null;
        return [transactionDate, '', row[descriptionIndex], '', '', -amount, cardName, row[cardMemberIndex]];
      } catch (e) { return null; }
    })
    .filter(row => row !== null);
}

/**
 * Simple ping function to test client-server connection.
 */
function testConnection() {
  return "Success! The server is reachable.";
}
