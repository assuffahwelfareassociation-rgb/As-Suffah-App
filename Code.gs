/**
 * Serves JSON balance data when requested via HTTP GET (e.g., from GitHub Pages fetch()).
 */
function doGet(e) {
  try {
    const data = getDashboardBalances();
    return ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    const errorPayload = {
      ok: false,
      error: err && err.message ? err.message : 'Unexpected server error.',
      totalBalance: '0.00',
      currentBalance: '0.00'
    };
    return ContentService
      .createTextOutput(JSON.stringify(errorPayload))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Fetches Total Balance (Dashboard!A26) and Current Balance (Dashboard!H26)
 * from the configured Google Spreadsheet using Script Properties.
 *
 * @return {Object} Balance values and formatted strings.
 */
function getDashboardBalances() {
  const scriptProps = PropertiesService.getScriptProperties();
  const SPREADSHEET_ID = scriptProps.getProperty('SPREADSHEET_ID');
  const SHEET_NAME = 'Dashboard';

  try {
    if (!SPREADSHEET_ID) {
      throw new Error('SPREADSHEET_ID is not configured in Script Properties.');
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      throw new Error('Sheet "' + SHEET_NAME + '" was not found in the spreadsheet.');
    }

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

    const numericCandidate = cleanedStr.replace(/BDT|৳|,|\s/gi, '');
    const numericVal = Number(numericCandidate);

    if (!isNaN(numericVal) && numericCandidate !== '') {
      const formattedNum = numericVal.toLocaleString('en-US', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
      });
      return { formatted: formattedNum };
    }

    return { formatted: cleanedStr };
  } catch (e) {
    return { formatted: String(displayValue || '0.00') };
  }
}
