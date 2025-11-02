
/**
 * Bootstrap runner for Apps Script IDE. Set constants below, then run QA_bootstrap.
 * It invokes the harness functions in a sensible order and logs outputs.
 */
function QA_bootstrap() {

  QA_setScriptProperties();
  try { QA_triggerUpdates(); Logger.log('QA_triggerUpdates: done'); } catch (e) { Logger.log('QA_triggerUpdates error: ' + e.message); }

/*  
  try { QA_geoValidate(SHEET_NAME, SPREADSHEET_URL_OR_ID); Logger.log('geoValidate: done'); } catch (e) { Logger.log('geoValidate error: ' + e.message); }
  try { QA_geoMap(SHEET_NAME, SPREADSHEET_URL_OR_ID); Logger.log('geoMap: done'); } catch (e) { Logger.log('geoMap error: ' + e.message); }
  try { QA_mergeLinks(MERGE_RANGE_NAME, SPREADSHEET_URL_OR_ID); Logger.log('mergeLinks: done'); } catch (e) { Logger.log('mergeLinks error: ' + e.message); }
  try { QA_mergeEmails(SHEET_NAME, true, SPREADSHEET_URL_OR_ID); Logger.log('mergeEmails(draft): done'); } catch (e) { Logger.log('mergeEmails(draft) error: ' + e.message); }
  try { QA_mergeToDocument(SHEET_NAME, DOC_TO_PDF, SPREADSHEET_URL_OR_ID); Logger.log('mergeToDocument: done'); } catch (e) { Logger.log('mergeToDocument error: ' + e.message); }
  try { QA_createIndex(SPREADSHEET_URL_OR_ID); Logger.log('createIndex: done'); } catch (e) { Logger.log('createIndex error: ' + e.message); }
  try { QA_updateIndex(SPREADSHEET_URL_OR_ID); Logger.log('updateIndex: done'); } catch (e) { Logger.log('updateIndex error: ' + e.message); }
  try { QA_exportSheet(SHEET_NAME, SPREADSHEET_URL_OR_ID); Logger.log('exportSheet: done'); } catch (e) { Logger.log('exportSheet error: ' + e.message); }
  try { log('importSS', QA_importSheets(IMPORT_SPREADSHEET_URL, SPREADSHEET_URL_OR_ID)); } catch (e) { Logger.log('importSS error: ' + e.message); }
  try { log('exportRange', QA_exportRange(RANGE_EXPORT_NAME, SPREADSHEET_URL_OR_ID)); } catch (e) { Logger.log('exportRange error: ' + e.message); }
  try { QA_moveNamedRangesToSpreadsheet(SHEET_NAME, SPREADSHEET_URL_OR_ID); Logger.log('moveNamedRangesToSpreadsheet: done'); } catch (e) { Logger.log('moveNamedRangesToSpreadsheet error: ' + e.message); }
  try { QA_moveSheet(SPREADSHEET_URL_OR_ID); Logger.log('moveSheet: done'); } catch (e) { Logger.log('moveSheet error: ' + e.message); }
  try { QA_duplicateSheet(DUP_SOURCE_NAME, DUP_DEST_NAME, SPREADSHEET_URL_OR_ID); Logger.log('duplicateSheet: done'); } catch (e) { Logger.log('duplicateSheet error: ' + e.message); }
  try { QA_flattenSheet(SHEET_NAME, SPREADSHEET_URL_OR_ID); Logger.log('flattenSheet: done'); } catch (e) { Logger.log('flattenSheet error: ' + e.message); }
  try { QA_autoHideRange('', SPREADSHEET_URL_OR_ID); Logger.log('autoHideRange: done'); } catch (e) { Logger.log('autoHideRange error: ' + e.message); }
  try { QA_autoUnhideRange('', SPREADSHEET_URL_OR_ID); Logger.log('autoUnhideRange: done'); } catch (e) { Logger.log('autoUnhideRange error: ' + e.message); }
  try { QA_refreshFormulas(SPREADSHEET_URL_OR_ID); Logger.log('refreshFormulas: done'); } catch (e) { Logger.log('refreshFormulas error: ' + e.message); }

  if (RUN_DESTRUCTIVE) {
    try { QA_detachAll(RANGE_EXPORT_NAME, '', SPREADSHEET_URL_OR_ID); Logger.log('detachAll: done'); } catch (e) { Logger.log('detachAll error: ' + e.message); }
    try { QA_removeSheet(SHEET_NAME, SPREADSHEET_URL_OR_ID); Logger.log('removeSheet: done'); } catch (e) { Logger.log('removeSheet error: ' + e.message); }
  }

*/

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
