/**
 * @OnlyCurrentDoc
 */

/**
 * Creates a custom menu in the Google Sheet UI when the spreadsheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Finance Tool')
      .addItem('Setup Sheet', 'setupSheet')
      .addItem('Load New File', 'loadNewFile')
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
  sheet.getRange('A:B').setNumberFormat('dd/MM/yyyy');

  // Freeze the header row
  sheet.setFrozenRows(1);

  SpreadsheetApp.getUi().alert('Sheet setup complete! The date format has been set to dd/MM/yyyy.');
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
  // 1. Process the newly uploaded file
  const parsedData = parseCsv(csvContent);
  if (parsedData.length < 2) {
    throw new Error('The CSV file appears to be empty or does not contain any transaction data.');
  }
  const headers = parsedData[0].map(h => h.trim());
  const dataRows = parsedData.slice(1);
  let newlyProcessedRows;

  if (headers.includes('Transaction Date') && headers.includes('Post Date')) {
    const requiredChaseHeaders = ['Transaction Date', 'Post Date', 'Description', 'Category', 'Type', 'Amount'];
    validateHeaders(headers, requiredChaseHeaders);
    newlyProcessedRows = processChaseData(dataRows, headers, cardName);
  } else if (headers.includes('Card Member')) {
    const requiredAmexHeaders = ['Date', 'Description', 'Card Member', 'Amount'];
    validateHeaders(headers, requiredAmexHeaders);
    newlyProcessedRows = processAmexData(dataRows, headers, cardName);
  } else {
    throw new Error('Could not determine file type. The file does not seem to be a supported Chase or Amex statement.');
  }

  if (newlyProcessedRows.length === 0) {
    throw new Error('The file was processed, but no valid transaction rows were found.');
  }

  // 2. Read existing data from the sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let existingData = [];
  if (sheet.getLastRow() > 1) {
    existingData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  }

  // 3. Combine new and existing data
  const combinedData = existingData.concat(newlyProcessedRows);

  // 4. Sort the combined data (Post Date at index 1, Transaction Date at index 0)
  combinedData.sort((a, b) => {
    const dateA = (a[1] instanceof Date) ? a[1] : a[0];
    const dateB = (b[1] instanceof Date) ? b[1] : b[0];
    return dateA.getTime() - dateB.getTime();
  });

  // 5. Clear the old data from the sheet
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }

  // 6. Write the sorted data back to the sheet
  if (combinedData.length > 0) {
    sheet.getRange(2, 1, combinedData.length, combinedData[0].length).setValues(combinedData);
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

/**
 * A universal date parser that tries multiple formats.
 * @param {string} dateString - The date string to parse.
 * @returns {Date|null} A Date object, or null if parsing fails.
 */
function parseDateUniversal(dateString) {
  if (!dateString || typeof dateString !== 'string') {
    return null;
  }

  // First, try the default constructor, which is flexible (handles MM/DD/YYYY, etc.)
  let date = new Date(dateString);
  if (date && !isNaN(date.getTime())) {
    return date;
  }

  // If that fails, try our specific yyyy/mm/dd parser
  date = parseYyyyMmDd(dateString);
  if (date && !isNaN(date.getTime())) {
    return date;
  }

  return null; // Return null if all attempts fail
}


/**
 * Parses a date string in yyyy/mm/dd format into a Date object.
 * @param {string} dateString - The date string to parse.
 * @returns {Date|null} A Date object, or null if the format is invalid.
 */
function parseYyyyMmDd(dateString) {
  if (!dateString || typeof dateString !== 'string') {
    return null;
  }
  const parts = dateString.split('/');
  if (parts.length !== 3) {
    return null;
  }
  const year = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10) - 1; // Month is 0-indexed in JS Date
  const day = parseInt(parts[2], 10);

  if (isNaN(year) || isNaN(month) || isNaN(day)) {
    return null;
  }

  const date = new Date(year, month, day);
  // Verify that the created date matches the input parts to catch invalid dates like 2023/02/30
  if (date.getFullYear() !== year || date.getMonth() !== month || date.getDate() !== day) {
    return null;
  }
  return date;
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
        const amountStr = row[amountIndex].replace(/[\$,]/g, '');
        const amount = parseFloat(amountStr);
        if (isNaN(amount)) return null;

        const transactionDate = parseDateUniversal(row[transactionDateIndex]);
        const postDate = parseDateUniversal(row[postDateIndex]);
        if (!transactionDate || !postDate) {
          console.error(`Skipping row due to invalid date format: ${row}`);
          return null;
        }

        return [
          transactionDate,
          postDate,
          row[descriptionIndex],
          row[categoryIndex],
          row[typeIndex],
          amount,
          cardName,
          ''
        ];
      } catch (e) {
        console.error(`Skipping invalid Chase row: ${row}. Error: ${e.message}`);
        return null;
      }
    })
    .filter(row => row !== null);

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
        const amountStr = row[amountIndex].replace(/[\$,]/g, '');
        const amount = parseFloat(amountStr);
        if (isNaN(amount)) return null;

        const transactionDate = parseDateUniversal(row[dateIndex]);
        if (!transactionDate) {
          console.error(`Skipping row due to invalid date format: ${row}`);
          return null;
        }

        return [
          transactionDate,
          '',
          row[descriptionIndex],
          '',
          '',
          -amount,
          cardName,
          row[cardMemberIndex]
        ];
      } catch (e) {
        console.error(`Skipping invalid Amex row: ${row}. Error: ${e.message}`);
        return null;
      }
    })
    .filter(row => row !== null);

  return processedRows;
}

/**
 * A robust CSV parser that handles quoted fields.
 * @param {string} csvContent The string content of the CSV file.
 * @returns {Array<Array<string>>} A 2D array of the parsed data.
 */
function parseCsv(csvContent) {
  const lines = csvContent.trim().split(/\r\n?|\n/);
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
