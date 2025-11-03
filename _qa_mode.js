
/**
 * Bootstrap runner for Apps Script IDE. Set constants below, then run QA_bootstrap.
 * It invokes the harness functions in a sensible order and logs outputs.
 */
function QA_bootstrap() {

  QA_setScriptProperties();
  try { QA_triggerUpdates(); Logger.log('QA_triggerUpdates: done'); } catch (e) { Logger.log('QA_triggerUpdates error: ' + e.message); }

}

// Harness wrappers
/**
 * Show the About dialog using the given spreadsheet (QA-friendly).
 * @param {string|Spreadsheet} spreadsheetUrlOrId
 */
function QA_triggerUpdates(spreadsheetUrlOrId) {
  updateShiftsAndEventMap();
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

function QA_setScriptProperties() {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  scriptProperties.setProperty('DEBUG', 'true');
  
  const addressConfig = {
    'Crist Mortuary': '3395 Penrose Pl, Boulder, CO 80301',
    'Greenwood & Myers Mortuary': '2969 Baseline Road, Boulder, CO 80303'
  };
  scriptProperties.setProperty('ADDRESS_CONFIG', JSON.stringify(addressConfig));  

  // Generate URL for email
  const webAppUrl = "https://script.google.com/macros/s/AKfycbxKxTyP7pmN1dspwuEUe1s-UVz4RwbADJiboj4G50w/dev"; 
  scriptProperties.setProperty('SCRIPT_URL', webAppUrl);  


  const sheetInputs = {
    SPREADSHEET_ID: '1cCouQRRpEN0nUhN45m14_z3oaONo7HHgwyfYDkcu2mw',
    EVENT_FORM_RESPONSES: 'Form Responses 1',
    SHIFTS_MASTER_SHEET: 'Shifts Master',
    GUESTS_SHEET: 'Guests',
    MEMBERS_SHEET: 'Members',
    EVENT_MAP: 'Event Map',
    ARCHIVE_EVENT_MAP: 'Archive Event Map'
  };
  scriptProperties.setProperty('SHEET_INPUTS', JSON.stringify(sheetInputs));
}

function QA_Logging(logMessage, DEBUG=false) {

  if (DEBUG) {
    console.log(logMessage);
  };

}
