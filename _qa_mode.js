/**
* -----------------------------------------------------------------
* _qa_mode.js
* Chevra Kadisha Shifts Scheduler
* QA mode utilized to debug the library
* -----------------------------------------------------------------
* _common_functions.js
Version: 1.0.6 * Last updated: 2025-11-12
 * 
 * CHANGELOG v1.0.6:
 *   - Fixed bug in usage of DEBUG
 *
 * Utility functions for Google Apps Script (suitable for Google Forms/Sheets integrations)
 * -----------------------------------------------------------------
 */
/**
 * Bootstrap runner for Apps Script IDE. Set constants below, then run QA_bootstrap.
 * It invokes the harness functions in a sensible order and logs outputs.
 */
function QA_bootstrap() {

  QA_configProperties();
  try { QA_triggerUpdates(QA_configProperties()); Logger.log('QA_triggerUpdates: done'); } catch (e) { Logger.log('QA_triggerUpdates error: ' + e.message); }

}

// Harness wrappers
/**
 * Trigger Updates.
 * @param {string|Spreadsheet} spreadsheetUrlOrId
 */
function QA_triggerUpdates(sheetInputs) {
  updateShiftsAndEventMap(sheetInputs);
}


/**
 * Return a Spreadsheet instance by url or id (required).
 * Delegates to getQDSSpreadsheet for QA-friendly resolution.
 * @param {string|Spreadsheet} spreadsheetUrlOrId URL, ID, or Spreadsheet object
 * @returns {Spreadsheet}
 * @example
 * var ss = QA_getSpreadsheet('https://docs.google.com/spreadsheets/d/ID/edit');
 */
function QA_getSpreadsheet(spreadsheetUrlOrId) {
  return getQDSSpreadsheet(spreadsheetUrlOrId);
}

function debugQAconfigproperties(){
  Logger.log(QA_configProperties());
}


function QA_configProperties() {

/**
 * Finds a Google Sheet by name in Google Drive and returns its ID.
 * @param {string} ssName The name of the Google Sheet file to find.
 * @returns {string | null} The Spreadsheet ID, or null if not found or an error occurs.
 */

  function getSpreadsheetByName_(ssName) {
    
    if (!ssName) {
      Logger.log('Error: A spreadsheet name was not provided.');
      return null;
    }

    try {
      // 1. Use the DriveApp service to search for files by name.
      // The query finds files with the exact title and limits the search to Google Sheet files.
      var files = DriveApp.getFilesByName(ssName);
      
      // 2. Check if any file was found
      if (files.hasNext()) {
        var file = files.next();
        
        // 3. Use SpreadsheetApp.openById() which requires a file's ID,
        // or simply use the File object's ID directly.
        var spreadsheetId = file.getId();
        
        Logger.log('Found Spreadsheet: "' + ssName + '", ID: ' + spreadsheetId); 
        return spreadsheetId;
        
      } else {
        // No file found with the exact name
        Logger.log('Error: Spreadsheet with name "' + ssName + '" not found in your Google Drive.');
        return null;
      }
      
    } catch (e) {
      Logger.log('An error occurred while accessing Drive: ' + e.toString());
      return null;
    }
  }

  const ssName = "bck_events_responses-shmira_schedule";
  const ss = getSpreadsheetByName_(ssName);

  const webAppUrl = 
  "https://script.google.com/macros/s/AKfycbyqS-grJJ9NU-YctOP5rYCCFO6Lj6vhUPJieXK_R-MzsgrYuJE/exec"

  // hardcode the names of the sheet databases
  const sheetInputs = {
    DEBUG: true,
    SCRIPT_URL: webAppUrl,
    SPREADSHEET_ID: ss,
    EVENT_FORM_RESPONSES: 'Form Responses 1',
    SHIFTS_MASTER_SHEET: 'Shifts Master',
    VOLUNTEER_LIST_SHEET: 'Volunteer Shifts',
    GUESTS_SHEET: 'Guests',
    MEMBERS_SHEET: 'Members',
    LOCATIONS_SHEET: 'Locations',
    EVENT_MAP: 'Event Map',
    ARCHIVE_EVENT_MAP: 'Archive Event Map',
    TOKEN_COLUMN_NUMBER: 12
  };

  return sheetInputs;

}

function QA_Logging(logMessage, DEBUG=false) {

  // --- QA CHECK ---
  if (typeof DEBUG === 'undefined' || DEBUG === false) {
    return;
  }
  
  console.log(logMessage);

}
