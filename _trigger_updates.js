/**
* -----------------------------------------------------------------
* _trigger_updates.js
* Chevra Kadisha Shifts Scheduler
* Trigger Updates
* -----------------------------------------------------------------
* _trigger_updates.js
Version: 1.0.6* Last updated: 2025-11-12
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
 *   v1.0.6:
 *   - Fixed bug in usage of DEBUG

 * Trigger Updates
 * -----------------------------------------------------------------
 */

/**
 * Triggers the updates for the shifts and event map.
 * Depends on the script properties being set.
 * @returns {void}
 */
function debugUpdateShifts(){
  const configProperties = QA_setScriptProperties();
  updateShiftsAndEventMap(configProperties)  

}


function updateShiftsAndEventMap(sheetInputs) {

  if (typeof sheetInputs.DEBUG === 'undefined') {
    console.log ("DEBUG is undefined");
    return;
  }

  console.log(sheetInputs);

  const webAppUrl = sheetInputs.SCRIPT_URL; 
  const addressConfig = sheetInputs.ADDRESS_CONFIG;  


  // Now use these variables in your function logic
  QA_Logging('triggeredFunction is called at ' + new Date().toISOString(), sheetInputs.DEBUG);
  QA_Logging('Spreadsheet ID: ' + sheetInputs.SPREADSHEET_ID, sheetInputs.DEBUG);
  updateShifts(sheetInputs);
  updateEventMap(sheetInputs,addressConfig,webAppUrl);
}
