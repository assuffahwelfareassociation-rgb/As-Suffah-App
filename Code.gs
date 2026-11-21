/**
 * Serves the main HTML web application for As-Suffah Welfare Association.
 * Note: Ensure the HTML file in the Apps Script project is named "Index" (Index.html).
 */
function doGet() {
  try {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('As-Suffah Welfare Association')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:20px;color:#b91c1c;">' +
      '<h3>Application Load Error</h3><p>' + err.message + '</p></div>'
    );
  }
}

/**
 * Fetches Total Balance (Dashboard!A26) and Current Balance (Dashboard!H26)
 * from the configured Google Spreadsheet.
 *
 * @return {Object} Balance values and formatted strings.
 */
function getDashboardBalances() {
  const SPREADSHEET_ID = '1VMWxM8b5gqfLz_FqENQfZ7BUFBLZmZQcwAwsrIgvDC0';
  const SHEET_NAME = 'Dashboard';

  try {
    if (!SPREADSHEET_ID) {
      throw new Error('Spreadsheet ID is missing.');
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      throw new Error('Sheet "' + SHEET_NAME + '" was not found in the spreadsheet.');
    }

    // Read A26 (Total Balance) and H26 (Current Balance)
    const rawTotal = sheet.getRange('A26').getDisplayValue();
    const rawCurrent = sheet.getRange('H26').getDisplayValue();

    const parsedTotal = parseBalanceValue(rawTotal, sheet.getRange('A26').getValue());
    const parsedCurrent = parseBalanceValue(rawCurrent, sheet.getRange('H26').getValue());

    return {
      ok: true,
      totalBalance: parsedTotal.formatted,
      currentBalance: parsedCurrent.formatted,
      rawTotal: rawTotal,
      rawCurrent: rawCurrent,
      updatedAt: new Date().toISOString()
    };
  } catch (error) {
    return {
      ok: false,
      error: error && error.message ? error.message : 'Unable to fetch balances.',
      totalBalance: '0.00',
      currentBalance: '0.00'
    };
  }
}

/**
 * Helper to normalize numeric or pre-formatted currency strings cleanly.
 *
 * @param {string} displayValue - Display string from the cell.
 * @param {*} rawCellValue - Raw underlying value from the cell.
 * @return {Object} Object containing clean formatted balance string.
 */
function parseBalanceValue(displayValue, rawCellValue) {
  try {
    if (typeof rawCellValue === 'number' && !isNaN(rawCellValue)) {
      const formattedNum = rawCellValue.toLocaleString('en-US', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
      });
      return { formatted: formattedNum };
    }

    const cleanedStr = String(displayValue || '').trim();
    if (!cleanedStr) {
      return { formatted: '0.00' };
    }

    // Strip currency labels like BDT, ৳, or commas to check if pure number
    const numericCandidate = cleanedStr.replace(/BDT|৳|,|\s/gi, '');
    const numericVal = Number(numericCandidate);

    if (!isNaN(numericVal) && numericCandidate !== '') {
      const formattedNum = numericVal.toLocaleString('en-US', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
      });
      return { formatted: formattedNum };
    }

    // Return original display text if already formatted as custom string
    return { formatted: cleanedStr };
  } catch (e) {
    return { formatted: String(displayValue || '0.00') };
  }
}
