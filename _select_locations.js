
/**
* -----------------------------------------------------------------
* _select_locations.js
* Chevra Kadisha Admin Menu functions
* -----------------------------------------------------------------
* _selection_locations.js
Version: 1.0.7 * Last updated: 2025-11-27
 * 
 * CHANGELOG v1.0.7:
 *   - Initial implementation of Select Locations.
 * -----------------------------------------------------------------
 */

/**
 * Returns all mortuary location records from the Locations sheet.
 * Expected headers: Mortuary Name, Street Address, City, State, Zip, Phone, Map URL
 *
 * @param {Object} sheetInputs - Configuration object containing SPREADSHEET_ID and LOCATIONS_SHEET.
 * @returns {Array<Object>} Array of location objects.
 */
function getLocations(sheetInputs) {
  Logger.log("Getting Locations from sheet: " + sheetInputs.LOCATIONS_SHEET);

  function getSafeValue(row, idx, header) {
    if (idx.hasOwnProperty(header)) {
      return row[idx[header]];
    } else {
      return '';
    }
  }

  try {
    const ss = getSpreadsheet_(sheetInputs.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetInputs.LOCATIONS_SHEET);
    if (!sheet) throw new Error("Sheet not found: " + sheetInputs.LOCATIONS_SHEET);

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      Logger.log("Locations sheet has no data rows.");
      return [];
    }

    const headers = data[0];

    // Map header names to indices
    var idx = {};
    headers.forEach(function(h, i) { idx[h] = i; });

    var locations = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Skip completely blank rows
      if (row.join('').trim() === '') continue;

      const loc = {
        name:        getSafeValue(row, idx, 'Mortuary Name'),
        street:      getSafeValue(row, idx, 'Street Address'),
        city:        getSafeValue(row, idx, 'City'),
        state:       getSafeValue(row, idx, 'State'),
        zip:         getSafeValue(row, idx, 'Zip'),
        phone:       getSafeValue(row, idx, 'Phone'),
        mapUrl:      getSafeValue(row, idx, 'Map URL'),
        notes:       getSafeValue(row, idx, 'Notes'),
        rowIndex:    i + 1 // 1-based row index in the sheet
      };

      locations.push(loc);
    }

    // Logger.log("Found " + locations.length + " locations.");
    return locations;

  } catch (e) {
    Logger.log("Error in getLocations: " + e.toString());
    throw e;
  }
}
