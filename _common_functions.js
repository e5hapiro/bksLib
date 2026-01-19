/**
 * -----------------------------------------------------------------
 * _common_functions.js
 * Chevra Kadisha Shifts Scheduler
 * Common functions for Google Apps Script (suitable for Google Forms/Sheets integrations)
 * -----------------------------------------------------------------
 * _common_functions.js
 *  Version: 1.0.7 
 *  Last updated: 2026-01-19
 * 
 * CHANGELOG v1.0.1:
 *   - Added enhanced error handling and logging to addToken.
 *   - Improved prevalidation and update detection logic in isFormUpdated.
 *   - Enhanced logging logic in logQCVars.
 *   - Added formattedDateAndTime for consistent date formatting.
 *   v1.0.6:
 *   - Fixed bug in usage of DEBUG
 *   v1.0.7:
 *   - Fixed formattedDateAndTime which required input of separate date and time
 *
 * Utility functions for Google Apps Script (suitable for Google Forms/Sheets integrations)
 * -----------------------------------------------------------------
 */


/**
 * Working around limitations of fixed web AppUrl
 */

function getWebAppUrl() {
  const webAppUrl = "https://script.google.com/macros/s/AKfycbzeF8O_va2qLvjnEYJqDrIfbZUHzPHdNSARHjLv0uFOsNwQpGMv9LrEJgXn5ilDezSr/exec"; 
  return webAppUrl
}



/**
 * Adds a unique token value (UUID) to the specified column in the row that triggered the event.
 * Only works if columnNumber is provided.
 * Logs success or detailed error for debugging.
 * 
 * @function
 * @param {Object} e - The event data object from a Google Sheets trigger.
 * @param {number} columnNumber - The target column number to receive the token.
 */
function addToken(e, columnNumber) {
  if (columnNumber) {
    try {
      var sheet = e.range.getSheet();
      var row = e.range.getRow();
      var uuid = Utilities.getUuid();
      sheet.getRange(row, columnNumber).setValue(uuid);
      Logger.log('Token added successfully for row: ' + row + ' column:' + columnNumber);
    } catch (error) {
      // Stores detailed information for easier debugging
      Logger.log('addToken failed for row: ' + (e && e.range ? e.range.getRow() : 'unknown') + ', error: ' + error.toString());
    }
    Logger.log('addToken failed no column provided ');
  }
}

/**
 * Determines whether a form submission/event represents an "update" condition.
 * Mainly for detecting race conditions or partially completed data.
 * 
 * @function
 * @param {Object} eventData - Object containing submissionDate and email keys at minimum.
 * @returns {boolean} - True if an update is detected, false otherwise.
 */
function isFormUpdated(eventData) {
  let formUpdated = false;

  // Validate required fields for prevalidation
  if (!eventData || !eventData.submissionDate || !eventData.email) {
    Logger.log('Error: Missing required event data fields for checking updates');
    return false;
  }

  // Check for update race condition
  if (eventData.submissionDate !== "" && eventData.email === "") {
    formUpdated = true;
  }

  return formUpdated;
}

/**
 * Quality Control Logger: Logs a set of variables with a context message.
 * ONLY logs if the global constant DEBUG is set to true.
 *
 * @param {string} context - A message describing where in the code this is being called.
 * @param {Object} varsObject - An object where keys are variable names and values are the variables.
 */
function logQCVars(context, varsObject) {
  // --- QA CHECK ---
  if (typeof DEBUG === 'undefined' || DEBUG === false) {
    return;
  }
  // --- END QA CHECK ---

  console.log(`--- QC LOG: ${context} ---`);
  
  if (typeof varsObject !== 'object' || varsObject === null) {
    console.log(`Invalid varsObject: ${varsObject}`);
    console.log(`--- END QC LOG: ${context} ---`);
    return;
  }

  for (const key in varsObject) {
    if (Object.prototype.hasOwnProperty.call(varsObject, key)) {
      const value = varsObject[key];
      
      if (typeof value === 'object' && value !== null) {
        try {
          console.log(`[${key}]: ${JSON.stringify(value)}`);
        } catch (e) {
          console.log(`[${key}] (Object): ${value.toString()}`);
        }
      } else {
        console.log(`[${key}]: ${value}`);
      }
    }
  }
  console.log(`--- END QC LOG: ${context} ---`);
}


/**
 * Returns a formatted English-language string for the supplied Date object.
 * Throws an error if input is not a valid Date.
 * 
 * @function
 * @param {Date} inputDate - JavaScript Date object.
 * @returns {string} - Date formatted as "Weekday, Month Day, Year at HH:MM AM/PM TZ".
 */
function formattedDateAndTime(inputDate, inputTime) {
  // Validate inputs
  if (!(inputDate instanceof Date) || isNaN(inputDate)) {
    throw new Error("Invalid date");
  }
  if (!(inputTime instanceof Date) || isNaN(inputTime)) {
    throw new Error("Invalid time");
  }

  // Combine date and time: use date part from inputDate, time part from inputTime
  const combinedDateTime = new Date(
    inputDate.getFullYear(),
    inputDate.getMonth(),
    inputDate.getDate(),
    inputTime.getHours(),
    inputTime.getMinutes(),
    inputTime.getSeconds()
  );

  const optionsDate = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
  const optionsTime = { hour: 'numeric', minute: '2-digit', hour12: true, timeZoneName: 'short' };

  const dateStr = combinedDateTime.toLocaleDateString('en-US', optionsDate);
  const timeStr = combinedDateTime.toLocaleTimeString('en-US', optionsTime);

  return `${dateStr} at ${timeStr}`;
}



/**
 * Aggressively cleans a string for robust token comparisons.
 *  - Removes all whitespace (space, tabs, newlines)
 *  - Removes invisible and Unicode control characters
 *  - Removes leading/trailing quotes, if present
 *  - Converts to lowercase (optional, for case-insensitive matching)
 *
 * @param {string} str
 * @param {boolean} [toLower] Should convert to lowercase? (Default false)
 * @returns {string}
 */
function normalizeToken(str, toLower) {
  if (typeof str !== 'string') str = String(str);

  // Remove leading/trailing whitespace, quotes, ALL whitespace, & invisible characters
  let cleaned = str
    .replace(/^["']+|["']+$/g, '')        // Remove leading/trailing quotes if present
    .replace(/[\u200B-\u200D\uFEFF]/g, '') // Remove zero-width/unicode invisible chars
    .replace(/\s+/g, '')                   // Remove ALL whitespace (space, tabs, newlines)
    .replace(/[\r\n\t]/g, '');             // Remove specific control chars

  if (toLower) cleaned = cleaned.toLowerCase();
  return cleaned;
}


/**
 * Safely opens the spreadsheet.
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} The spreadsheet object.
 * @private
 */
function getSpreadsheet_(ss_id) {
  return SpreadsheetApp.openById(ss_id);
}


/**
 * Ensures a sheet exists in the spreadsheet. If it does not exist, creates it.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet object.
 * @param {string} sheetName The name of the sheet to ensure exists.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet object (either existing or newly created).
 * @private
 */
function ensureSheetExists_(sheetInputs, ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (typeof sheetInputs !== 'undefined' && sheetInputs.DEBUG) {
      console.log('Created new sheet: ' + sheetName);
    }
  }
  
  return sheet;
}

/**
 * Ensures a sheet has headers in its first row. If the first row is empty or missing,
 * inserts a new row at the top with the provided headers.
 * 
 * If the sheet is empty (no data), inserts the header row.
 * If the sheet has data and row 1 is not empty, assumes headers are already present.
 * If you need to force headers even when data exists, modify the logic accordingly.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to check/modify.
 * @param {Array<string>} headers Array of header column names.
 * @private
 */
function ensureSheetHeaders_(sheetInputs, sheet, headers) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  // If sheet is completely empty, insert header row at row 1
  if (lastRow === 0) {
    sheet.insertRows(1, 1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (typeof sheetInputs !== 'undefined' && sheetInputs.DEBUG) {
      console.log('Inserted headers into empty sheet.');
    }
    return;
  }
  
  // Check if first row looks like headers by examining a few cells
  var firstRowData = sheet.getRange(1, 1, 1, Math.min(headers.length, lastCol)).getValues()[0];
  
  // If first row is all empty cells, it's likely a placeholder row
  var firstRowEmpty = firstRowData.every(function(cell) {
    return cell === '' || cell === null;
  });
  
  if (firstRowEmpty && lastRow >= 1) {
    // First row is empty but data exists; replace row 1 with headers
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (typeof sheetInputs !== 'undefined' && sheetInputs.DEBUG) {
      console.log('Replaced empty first row with headers.');
    }
    return;
  }
  
  // If first row is not empty, assume headers already exist
  if (!firstRowEmpty) {
    if (typeof sheetInputs !== 'undefined' && sheetInputs.DEBUG) {
      console.log('First row already contains data; assuming headers present.');
    }
    return;
  }
  
  // Fallback: if we reach here and lastRow > 0 but data exists below, 
  // insert a new row at the top
  if (lastRow > 0) {
    sheet.insertRows(1, 1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (typeof sheetInputs !== 'undefined' && sheetInputs.DEBUG) {
      console.log('Inserted headers at top of sheet.');
    }
  }
}
