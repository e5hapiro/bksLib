/**
* -----------------------------------------------------------------
* _select_locations.js
* Chevra Kadisha Admin Menu functions
* -----------------------------------------------------------------
* Version: 1.1.1 * Last updated: 2026-04-06
 * 
 * CHANGELOG v1.1.0:
 *   - Converted to vertical format: Col A=keys (Mortuary Name, Street Address,..., Line 30), Col B+=locations.
 * CHANGELOG v1.1.1:
 *   - Updated getAddressFromLocation eliminating trim of a null string
 * -----------------------------------------------------------------
 */

/**
 * Returns all mortuary location records from the Locations sheet (vertical format).
 * Expected: Row 1 headers ["Key", "Crist Mortuary", "Greenwood & Myers", ...]
 * Col A keys: Mortuary Name, Street Address, City, State, Zip, Phone, Map URL, Notes, Line 1-30
 *
 * @param {Object} sheetInputs - Config with SPREADSHEET_ID, LOCATIONS_SHEET.
 * @returns {Array<Object>} Array of location objects.
 */
function getLocations(sheetInputs) {
  Logger.log("Getting Locations from sheet: " + sheetInputs.LOCATIONS_SHEET);

  try {
    const ss = getSpreadsheet_(sheetInputs.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetInputs.LOCATIONS_SHEET);
    if (!sheet) throw new Error("Sheet not found: " + sheetInputs.LOCATIONS_SHEET);

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      Logger.log("Locations sheet has no data rows.");
      return [];
    }

    const headers = data[0]; // ["Key", "Crist Mortuary", "Greenwood & Myers", ...]
    const locationNames = headers.slice(1); // Location names (Col B+)

    var locations = [];

    // For each location column (Col B+)
    for (let col = 1; col < headers.length; col++) {
      const locName = headers[col].toString().trim();
      if (!locName) continue;

      const location = { name: locName };

      // For each key row (Row 1+: Mortuary Name, Street Address, etc.)
      for (let row = 1; row < data.length; row++) {
        const keyName = data[row][0]?.toString().trim(); // Col A
        if (!keyName) continue;

        const value = data[row][col] || ''; // Current location column
        const propName = keyName.toLowerCase().replace(/ /g, ''); // 'Street Address' -> 'streetaddress', 'Line 1' -> 'line1'

        location[propName] = value;
      }

      location.rowIndex = col + 1; // 1-based column index


      // Dynamic logging for all properties
      //Logger.log('Location "%s" (col %s):', locName, location.rowIndex);
      //Object.keys(location).forEach(prop => {
      //  Logger.log('  %s: %s', prop, location[prop]);
      //});


      locations.push(location);
    }

    Logger.log("Found " + locations.length + " locations.");
    return locations;

  } catch (e) {
    Logger.log("Error in getLocations: " + e.toString());
    throw e;
  }
}

/**
 * Looks up full address by location name from vertical locations array.
 * @param {Array<Object>} locations - From getLocations().
 * @param {string} locationName - e.g., "Crist Mortuary".
 * @returns {string} Formatted address or locationName if not found.
 */
function getAddressFromLocationName(locations, locationName) {        
  const foundLocation = locations.find(loc => loc.name === locationName);
  
  if (!foundLocation) {
    Logger.log(`Location "${locationName}" not found.`);
    return locationName;
  }

  // 1. Collect potential parts
  const rawParts = [
    foundLocation.streetaddress,
    foundLocation.city,
    foundLocation.state,
    foundLocation.zip
  ];

  // 2. Filter, Trim, and Join
  const address = rawParts
    .filter(part => part != null && String(part).trim() !== "") // Remove null/undefined/empty
    .map(part => String(part).trim())                          // Clean whitespace
    .join(', ');

  return address || locationName;
}

