/**
* -----------------------------------------------------------------
* _importSpreadsheet.js
* Chevra Kadisha Shifts Scheduler
* Import Spreadsheet
* -----------------------------------------------------------------
* _importSpreadsheet.js
 * Version: 1.0.1
* Last updated: 2025-11-02
 * -----------------------------------------------------------------
 */
  
function importSpreadsheet(spreadsheetUrl) {

  const importSpreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
  const thisSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // ensures that any errors thrown by Google are passed up and captured by the logger    
  try {

    // Step 1 - Loops through all sheets in the imported spreadsheet
    var importSheets = importSpreadsheet.getSheets();
    var thisSheets = thisSpreadsheet.getSheets();

    importSheets.forEach(function(sheet) {

      // Step 2 - get sheet name in order to rename the sheet when moved to existing workbook
      var sheetName = sheet.getName()

      // Step 3 - Check for existence of destination name 
      if (sheetExists(thisSheets,sheetName) 
        || sheetExists(thisSheets,"copy of" + sheetName)) {
        var eMsg = "(importSS) Error: Destination sheet exists already - " + sheetName;
        throw new Error(eMsg);
      }

      // Step 4 - all other sheets are to be imported to current spreadsheet
      var destinationSheet = sheet.copyTo(thisSpreadsheet);
      destinationSheet.setName(sheetName);
 
    });

  // Step 12 - if any error occurs copying the sheet then pass it back to local sheet to handle
  } catch (error) {
    throw error;
  }

  return true;




    /**
     * QDS Function: sheetExists()
     * Purpose: Accepts list of sheets and checks if sheet name exists
     * Scope: importParamLocalization() only
     * @param {array} sheets - array of sheets object
     * @param {string} sheetName - name of sheet to find
     * @return {importParams} object required by the import functions
     */
  function sheetExists(sheets, sheetName) {

    for (var i = 0; i < sheets.length; i++) {
        if (sheets[i].getName() === sheetName) {
          return true; // Sheet with the given name exists
        }
      }
  
    return false; // Sheet with the given name does not exist

  }

}


/**
 * Function: debugImportSpreadsheet()
 * Purpose: QA tool for debugging importSpreadsheet functions
 *          Note that qdssheets references will need to be added 
 */
function debugImportSpreadsheet(){



  // Step 2 - provide a hardcoded URL for testing
  // var sheetUrl = localInput(localization);
  var sheetUrl = "https://docs.google.com/spreadsheets/d/11C5qlSYgWNLuBP5Bslh-6Qtu8Y_vcBP3fyXZUfBltuM/edit?usp=sharing";

  // Step 3 - validate URL
  if (sheetUrl === null) {
    getAlert("No input received")
    return;
  }
  else
  {

    var response = importSpreadsheet(sheetUrl);
    
    // Step 5 - Log success and indicate same to user
    if (response) {
      Logger.log("Sheets have been successfully imported from:" + sheetUrl);
    } 
  }


}

