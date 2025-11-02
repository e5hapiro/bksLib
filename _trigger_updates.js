/**
* -----------------------------------------------------------------
* _trigger_updates.js
* Chevra Kadisha Shifts Scheduler
* Trigger Updates
* -----------------------------------------------------------------
* _trigger_updates.js
 * Version: 1.0.0
* Last updated: 2025-11-02
 * 
 * CHANGELOG v1.0.1:
 *   - Initial implementation of triggeredFunction.
 *   - Added logging and error handling.
 *   - Added shift and event map updates.
 *   - Added debug logging.
 *   - Added spreadsheet ID and sheet inputs retrieval.
 *   - Added shift and event map updates.
 *   - Added debug logging.
 *   - Added spreadsheet ID and sheet inputs retrieval.
 *   - Added shift and event map updates.
 *   - Added debug logging.
 *   - Added spreadsheet ID and sheet inputs retrieval.
 * Trigger Updates
 * -----------------------------------------------------------------
 */

/**
 * Triggers the updates for the shifts and event map.
 * Depends on the script properties being set.
 * @returns {void}
 */

function updateShiftsAndEventMap() {
  const scriptProperties = PropertiesService.getScriptProperties();
  if (!scriptProperties) throw new Error('Script properties not found');
  const DEBUG = scriptProperties.getProperty('DEBUG') === 'true';
  const addressConfig = JSON.parse(scriptProperties.getProperty('ADDRESS_CONFIG'));
  const SPREADSHEET_ID = scriptProperties.getProperty('SPREADSHEET_ID');
  const sheetInputs = JSON.parse(scriptProperties.getProperty('SHEET_INPUTS'));
  
  // Now use these variables in your function logic
  QA_Logging('triggeredFunction is called at ' + new Date().toISOString(), DEBUG);
  QA_Logging('Spreadsheet ID: ' + sheetInputs.SPREADSHEET_ID, DEBUG);
  updateShifts(sheetInputs, DEBUG);
  updateEventMap(sheetInputs, DEBUG);
}
