
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


/**
 * Looks up the confidential physical address based on the location name (e.g., 'Site A').
 * This function retrieves the secret address stored in ADDRESS_CONFIG.
 * @param {string} locationName The short name (e.g., 'Site A' or 'Site B').
 * @returns {string} The full physical address or a helpful message.
 * @private
 */
/*
* Function to find a location by name in an array of location objects 
* and return the full address as a concatenated string.
*/
function getAddressFromLocationName(locations, locationName) {        
  // 1. Find the location object in the array where location.name matches locationName.
  const foundLocation = locations.find(location => location.name === locationName);

  // 2. If a matching location object is found, construct and return the full address string.
  if (foundLocation) {
    // Access properties directly and concatenate: "Street, City, State Zip"
    return `${foundLocation.street}, ${foundLocation.city}, ${foundLocation.state} ${foundLocation.zip}`; 
  }

  // 3. If the location is not configured (e.g. 'Virtual Shift', 'Other'), return the original location name.
  return locationName;
}      

