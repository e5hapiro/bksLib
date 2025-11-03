/**
* -----------------------------------------------------------------
* _sync_shifts.js
* Chevra Kadisha Shifts Scheduler
* Shift Synchronization
* -----------------------------------------------------------------
* _sync_shifts.js
Version: 1.0.2 * Last updated: 2025-11-02
 * 
 * CHANGELOG v1.0.1:
 *   - Initial implementation of createHourlyShifts_.
 *   - Added logging and error handling.
 *   - Added shift synchronization to the Shifts Master sheet.
 * Shift Synchronization
 * -----------------------------------------------------------------
 */

/**
 * Updates the shifts master sheet with the latest data from the events sheet.  Creates new shifts and updates the shifts master sheet.
 * @returns {void}
 * @returns {Array<object>} An array of structured shift objects.
 */
function updateShifts(sheetInputs, DEBUG) {

  var shifts = []; 
  var events = [];

  try {

    // The master workbook
    const ss = getSpreadsheet_(sheetInputs.SPREADSHEET_ID);

    // The events sheet
    const eventSheet = ss.getSheetByName(sheetInputs.EVENT_FORM_RESPONSES);
    if (!eventSheet) throw new Error(`Sheet not found: ${sheetInputs.EVENT_FORM_RESPONSES}`);

    var events = getEvents(eventSheet);

    // The shifts master sheet
    const shiftsMasterSheet = ss.getSheetByName(sheetInputs.SHIFTS_MASTER_SHEET);
    if (!shiftsMasterSheet) throw new Error(`Sheet not found: ${sheetInputs.SHIFTS_MASTER_SHEET}`);
    var shifts = getExistingShiftsMasterRows(shiftsMasterSheet);

    syncShiftsToMaster_(sheetInputs, DEBUG, events, shifts);


  } catch (error) {
    QA_Logging("Error in updateShifts_: " + error.toString(), DEBUG);
    throw error;
  }

  return shifts;

}


/**
 * Retrieves the existing shifts master rows from the shifts master sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The shifts master sheet.
 * @returns {Array<object>} - The existing shifts master rows.
 */
function getExistingShiftsMasterRows(sheet) {
  var data = sheet.getDataRange().getValues();
  return data; // include headers
}


/**
 * Syncs events into the master sheet shifts list with proper key mapping.
 */
function syncShiftsToMaster_(sheetInputs, DEBUG, events, shiftsWithHeader) {
  var allNewShifts = [];

  var headers = shiftsWithHeader[0];           // Extract header row from sheet data
  var shiftsData = shiftsWithHeader.slice(1);  // Shift data rows

  // Convert sheet rows to normalized objects with consistent keys
  var shifts = shiftsData.map(row => rowToShiftObj(row, headers));

  // Helper: Creates a unique key based on normalized shift object properties
  function getShiftKey(shift) {
    return [
      shift.deceasedName,
      shift.eventLocation,
      shift.startTimeEpoch,
      shift.endTimeEpoch
    ].join('|');
  }

  var validShiftKeys = new Set();
  var existingShiftKeys = new Set();

  shifts.forEach(shift => existingShiftKeys.add(getShiftKey(shift)));

  var allCurrentEventShifts = [];
  events.forEach(event => {
    var eventShifts = createHourlyShifts_(event, DEBUG);
    eventShifts.forEach(shift => {
      validShiftKeys.add(getShiftKey(shift));
      allCurrentEventShifts.push(shift);
      if (!existingShiftKeys.has(getShiftKey(shift))) {
        allNewShifts.push(shift);
      }
    });
  });

  var rowsToDelete = [];
  shifts.forEach((shift, idx) => {
    var key = getShiftKey(shift);
    if (!validShiftKeys.has(key)) {
      rowsToDelete.push(idx + 2); // Account for header + 1-based indexing
    }
  });

  if (rowsToDelete.length > 0) {
    const ss = getSpreadsheet_(sheetInputs.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetInputs.SHIFTS_MASTER_SHEET);

    rowsToDelete.sort((a, b) => b - a);

    // Prevent deleting all data rows, leave at least one after header
    while (sheet.getLastRow() - rowsToDelete.length < 2 && rowsToDelete.length > 0) {
      rowsToDelete.pop();
    }

    rowsToDelete.forEach(rowNum => sheet.deleteRow(rowNum));
  }

  if (allNewShifts.length > 0) {
    syncShiftsToSheet(sheetInputs, DEBUG, allNewShifts);
  }
}


/**
 * Creates individual hourly shift objects from a start and end time.
 * @param {object} eventData The event data object containing all form inputs and parsed Date objects.
 * @returns {Array<object>} An array of structured shift objects.
 * @private
 */
function createHourlyShifts_(eventData, DEBUG=false) {
  QA_Logging("=== Creating Shifts ===", DEBUG);
  QA_Logging("=== FULL eventData OBJECT ===", DEBUG);
  QA_Logging(JSON.stringify(eventData, null, 2), DEBUG);
  QA_Logging("=== END eventData ===", DEBUG);

  if (!eventData || !eventData.startDate || !eventData.endDate || !eventData.deceasedName || !eventData.locationName) {
    QA_Logging('Error: Missing required event data fields', DEBUG);
    return [];
  }

  const shifts = [];
  let currentStart = new Date(eventData.startDate.getTime());
  const endDate = new Date(eventData.endDate.getTime());

  if (currentStart >= endDate) {
    QA_Logging('Warning: Start date is not before end date. No shifts created.', DEBUG);
    return shifts;
  }

  while (currentStart < endDate) {
    let currentEnd = new Date(currentStart.getTime());
    currentEnd.setHours(currentStart.getHours() + 1);

    if (currentEnd > endDate) {
      currentEnd = new Date(endDate.getTime());
    }
    
    const dateStr = currentStart.toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric' });
    const timeStr = `${currentStart.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' })} - ${currentEnd.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' })}`;

    const shiftId = Utilities.getUuid();

    shifts.push({
      id: shiftId,
      deceasedName: eventData.deceasedName,
      eventLocation: eventData.locationName,
      eventDate: dateStr,
      shiftTime: timeStr,
      maxVolunteers: 1,
      currentVolunteers: 0,
      startTimeEpoch: currentStart.getTime(),
      endTimeEpoch: currentEnd.getTime(),
      pronoun: eventData.pronoun,
      metOrMeita: eventData.metOrMeita,
      personalInfo: eventData.personalInfo
    });

    currentStart = new Date(currentEnd.getTime());
  }

  return shifts;
}



/**
 * Synchronizes the generated shifts to the Shifts Master sheet.
 * @param {Array<object>} allNewShifts Array of shift objects.
 * @private
 */
/**
 * Synchronizes the generated shifts to the Shifts Master sheet.
 * @param {Array<object>} allNewShifts Array of shift objects.
 * @private
 */
function syncShiftsToSheet(sheetInputs, DEBUG, allNewShifts) {
  try {
    const ss = getSpreadsheet_(sheetInputs.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetInputs.SHIFTS_MASTER_SHEET);
    if (!sheet) throw new Error(`Sheet not found: ${sheetInputs.SHIFTS_MASTER_SHEET}`);

    const newRows = allNewShifts.map(shift => [
      shift.id,                 // 0 - Shift ID
      shift.deceasedName,       // 1 - Deceased Name (was shift.eventName)
      shift.eventLocation,      // 2 - Location (was at index 7)
      shift.eventDate,          // 3 - Event Date
      shift.shiftTime,          // 4 - Shift Time
      shift.maxVolunteers,      // 5 - Max Volunteers
      shift.currentVolunteers,  // 6 - Current Volunteers
      shift.startTimeEpoch,     // 7 - Start Epoch
      shift.endTimeEpoch,       // 8 - End Epoch
      shift.pronoun,            // 9 - Pronoun (NEW)
      shift.metOrMeita,         // 10 - Met or Meita (NEW)
      shift.personalInfo        // 11 - Personal Info (NEW)
    ]);

    // Check if the sheet is empty to write headers
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Shift ID',           // 0
        'Deceased Name',      // 1
        'Location',           // 2
        'Event Date',         // 3
        'Shift Time',         // 4
        'Max Volunteers',     // 5
        'Current Volunteers', // 6
        'Start Epoch',        // 7
        'End Epoch',          // 8
        'Pronoun',            // 9
        'Met or Meita',       // 10
        'Personal Info'       // 11
      ]);
    }
    
    // Append all new shift rows in one go
    if (newRows.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      const range = sheet.getRange(startRow, 1, newRows.length, newRows[0].length);
      range.setValues(newRows);
    }
    
  } catch (error) {
    QA_Logging("Error in syncShiftsToSheet: " + error.toString(), DEBUG);
    throw error;
  }
}


// Mapping from sheet header names to createHourlyShifts_ object keys
const headerToObjectKey = {
  "Shift ID": "id",
  "Deceased Name": "deceasedName",
  "Location": "eventLocation",
  "Event Date": "eventDate",
  "Shift Time": "shiftTime",
  "Max Volunteers": "maxVolunteers",
  "Current Volunteers": "currentVolunteers",
  "Start Epoch": "startTimeEpoch",
  "End Epoch": "endTimeEpoch",
  "Pronoun": "pronoun",
  "Met-or-Meita": "metOrMeita",
  "Personal Information": "personalInfo"
};

/**
 * Converts a sheet row array to a standardized shift object using header mapping
 */
function rowToShiftObj(row, headers) {
  var obj = {};
  headers.forEach(function(header, i) {
    var key = headerToObjectKey[header];
    if (key) obj[key] = row[i];
  });
  return obj;
}


