
/**
* -----------------------------------------------------------------
* _select_decedents.js
* Chevra Kadisha Admin Menu functions
* -----------------------------------------------------------------
* _selection_decedents.js
Version: 1.0.1 * Last updated: 2026-03-24
 * 
 * CHANGELOG v1.0.0:
 *   - Initial implementation of Get Decedents.
 * CHANGELOG v1.0.1:
 *   - Fixed bug where it was getting data from the master instead of events
 * -----------------------------------------------------------------
 */

/**
 * Returns all decedent records from the Latest Master sheet.
 * Expected headers: Mortuary Name, Street Address, City, State, Zip, Phone, Map URL
 *
 * @param {Object} sheetInputs - Configuration object containing SPREADSHEET_ID and LOCATIONS_SHEET.
 * @returns {Array<Object>} Array of location objects.
 */
function getDecedents(sheetInputs) {
  Logger.log("Getting Decedents from sheet: " + sheetInputs.LATEST_EVENTS);

  function getSafeValue(row, idx, header) {
    if (idx.hasOwnProperty(header)) {
      return row[idx[header]];
    } else {
      return '';
    }
  }

  const headerRowNumber = 2;

  try {
    const ss = getSpreadsheet_(sheetInputs.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetInputs.LATEST_EVENTS);
    if (!sheet) throw new Error("Sheet not found: " + sheetInputs.LATEST_EVENTS);

    const data = sheet.getDataRange().getValues();
    if (data.length < headerRowNumber) {
      Logger.log("Decedents sheet has no data rows.");
      return [];
    }

    const headers = data[headerRowNumber-1];    // Account for zero

    // Map header names to indices
    var idx = {};
    headers.forEach(function(h, i) { idx[h] = i; });

    var decedents = [];

    for (let i = headerRowNumber; i < data.length; i++) {
      const row = data[i];

      // Skip completely blank rows
      if (row.join('').trim() === '') continue;

      const decedent = {
        name:               getSafeValue(row, idx, 'Deceased Name'),
        location:           getSafeValue(row, idx, 'Location'),
        date:               getSafeValue(row, idx, 'Start Date'),
        time:               getSafeValue(row, idx, 'Start Time'),
        info:               getSafeValue(row, idx, 'Personal Information'),
        pronoun:            getSafeValue(row, idx, 'Pronoun'),
        meta:               getSafeValue(row, idx, 'Met or Meta'),
        rowIndex:    i + 1 // 1-based row index in the sheet
      };

      decedents.push(decedent);
    }

    Logger.log("Found " + decedents.length + " decedents.");
    return decedents;

  } catch (e) {
    Logger.log("Error in getDecedents: " + e.toString());
    throw e;
  }
}
