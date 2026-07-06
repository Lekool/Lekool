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
 * Returns the names of the existing tabs so the upload dialog can build its
 * card dropdown. The list is self-maintaining: a tab created by a first upload
 * shows up here on the next one.
 * @returns {string[]} Existing sheet (card) names.
 */
function getCardNames() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());
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
    let detectedFormat;

    if (headers.includes('Transaction Date') && headers.includes('Post Date')) {
      detectedFormat = 'Chase';
      const requiredChaseHeaders = ['Transaction Date', 'Post Date', 'Description', 'Category', 'Type', 'Amount'];
      validateHeaders(headers, requiredChaseHeaders);
      processedRows = processChaseData(dataRows, headers, cardName);
    } else if (headers.includes('Card Member')) {
      detectedFormat = 'Amex';
      const requiredAmexHeaders = ['Date', 'Description', 'Card Member', 'Amount'];
      validateHeaders(headers, requiredAmexHeaders);
      processedRows = processAmexData(dataRows, headers, cardName);
    } else {
      throw new Error('Could not determine file type. The file does not seem to be a supported Chase or Amex statement.');
    }

    if (processedRows.length === 0) {
      throw new Error('The file was processed, but no valid transaction rows were found.');
    }

    const existingSheet = spreadsheet.getSheetByName(cardName);

    // 2a. New card: create the tab, set it up, sort, write.
    if (!existingSheet) {
      processedRows.sort(compareByDate);
      const newSheet = spreadsheet.insertSheet(cardName);
      setupNewSheet(newSheet);
      newSheet.getRange(2, 1, processedRows.length, processedRows[0].length).setValues(processedRows);
      applyAmountFormatting(newSheet);
      return { card: cardName, format: detectedFormat, added: processedRows.length, skipped: 0, created: true };
    }

    // 2b. Existing card: guard against a cross-issuer mistake (Chase file into an
    // Amex tab or vice versa). This is Chase-vs-Amex only; the three Chase cards
    // are indistinguishable by content, so it cannot police Chase-vs-Chase.
    const sheetFormat = detectSheetFormat(existingSheet);
    if (sheetFormat && sheetFormat !== detectedFormat) {
      throw new Error(`Format mismatch: "${cardName}" already holds ${sheetFormat} transactions, but this file is a ${detectedFormat} export. Choose the matching card or a different name.`);
    }

    // 2c. Existing card: dedupe against current rows, append survivors, re-sort
    // the whole tab, and rewrite it.
    const lastRow = existingSheet.getLastRow();
    const existingRows = lastRow > 1
      ? existingSheet.getRange(2, 1, lastRow - 1, 8).getValues()
      : [];

    const seen = {};
    existingRows.forEach(r => { seen[dedupKey(r)] = true; });

    const newRows = [];
    let skipped = 0;
    processedRows.forEach(r => {
      const key = dedupKey(r);
      if (seen[key]) { skipped++; return; }
      seen[key] = true; // also collapse duplicates within this same upload
      newRows.push(r);
    });

    if (newRows.length === 0) {
      return { card: cardName, format: detectedFormat, added: 0, skipped: skipped, created: false };
    }

    const combined = existingRows.concat(newRows);
    combined.sort(compareByDate);

    // Clear the old data region first so a shorter rewrite can't leave stragglers.
    if (lastRow > 1) {
      existingSheet.getRange(2, 1, lastRow - 1, 8).clearContent();
    }
    existingSheet.getRange(2, 1, combined.length, 8).setValues(combined);
    applyAmountFormatting(existingSheet);

    return { card: cardName, format: detectedFormat, added: newRows.length, skipped: skipped, created: false };

  } catch (err) {
    // Re-throw the error so it can be caught by the client-side failure handler
    // Ensure the message is clean
    throw new Error(err.message);
  }
}

// --- Helper Functions ---

/**
 * Shared sort comparator. Sorts ascending by Post Date (slot 1) when it is a
 * real Date, otherwise falls back to Transaction Date (slot 0). Amex rows and
 * Chase rows with an unparseable Post Date carry '' in slot 1, which is the
 * sentinel that triggers the fallback.
 */
function compareByDate(a, b) {
  const dateA = (a[1] instanceof Date) ? a[1] : a[0];
  const dateB = (b[1] instanceof Date) ? b[1] : b[0];
  return dateA.getTime() - dateB.getTime();
}

/**
 * Builds the duplicate-detection key for a unified-schema row:
 * Transaction Date + Amount + Description + Type. Two rows sharing this key are
 * treated as the same transaction on append. Tradeoff: two genuinely identical
 * charges (same day, amount, merchant) collapse to one.
 */
function dedupKey(row) {
  const txDate = (row[0] instanceof Date) ? row[0].getTime() : String(row[0]);
  return [txDate, row[5], row[2], row[4]].join('|');
}

/**
 * Best-effort format guess for an existing tab, used only for the cross-issuer
 * mismatch guard. Heuristic: Amex writes the Card Member into the Notes column
 * (slot 7) while Chase leaves it empty, so a tab whose rows mostly have Notes is
 * Amex. Returns 'Chase', 'Amex', or null when the tab has no data rows.
 */
function detectSheetFormat(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const notes = sheet.getRange(2, 8, lastRow - 1, 1).getValues();
  let withNotes = 0;
  notes.forEach(r => { if (r[0] !== '' && r[0] != null) withNotes++; });
  return (withNotes > notes.length / 2) ? 'Amex' : 'Chase';
}

/**
 * Applies conditional formatting to the Amount column (F): light-blue fill for
 * negative amounts (charges), light-green for positive (credits/payments), no
 * fill at zero. Idempotent — replaces any existing rules on the sheet, so it is
 * safe to call on both new and re-written tabs.
 */
function applyAmountFormatting(sheet) {
  const amountRange = sheet.getRange('F2:F');
  const negativeRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground('#cfe2f3') // light blue
      .setRanges([amountRange])
      .build();
  const positiveRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground('#d9ead3') // light green
      .setRanges([amountRange])
      .build();
  sheet.setConditionalFormatRules([negativeRule, positiveRule]);
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
        // Keep the row as long as Transaction Date parses. A bad Post Date falls
        // back to the '' sentinel (same as Amex) instead of dropping the row.
        if (!transactionDate) return null;
        return [transactionDate, postDate || '', row[descriptionIndex], row[categoryIndex], row[typeIndex], amount, cardName, ''];
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
