/**
* -----------------------------------------------------------------
* _qa_mode.js
* Chevra Kadisha Shifts Scheduler
* QA mode utilized to debug the library
* -----------------------------------------------------------------
* _common_functions.js
Version: 1.0.8 * Last updated: 2026-03-26
 * 
 * CHANGELOG v1.0.6:
 *   - Fixed bug in usage of DEBUG
 * CHANGELOG v1.0.7:
 *   - Added Decedents
 * CHANGELOG v1.0.8:
 *   - Added Locations and Emails
 *
 * Utility functions for Google Apps Script (suitable for Google Forms/Sheets integrations)
 * -----------------------------------------------------------------
 */
/**
 * Bootstrap runner for Apps Script IDE. Set constants below, then run QA_bootstrap.
 * It invokes the harness functions in a sensible order and logs outputs.
 */
function QA_bootstrap() {

  try { QA_triggerUpdates(QA_configProperties()); Logger.log('QA_triggerUpdates: done'); } catch (e) { Logger.log('QA_triggerUpdates error: ' + e.message); }
  //try { QA_archiveUpdates(QA_configProperties()); Logger.log('QA_archiveUpdates: done'); } catch (e) { Logger.log('QA_archiveUpdates error: ' + e.message); }

  //try { QA_getEmails(QA_configProperties()); Logger.log('QA_getEmails: done'); } catch (e) { Logger.log('QA_getEmails error: ' + e.message); }
  // try { QA_getLocations(QA_configProperties()); Logger.log('QA_getLocations: done'); } catch (e) { Logger.log('QA_getLocations error: ' + e.message); }
  //try { QA_getDecedents(QA_configProperties()); Logger.log('QA_getDecedents: done'); } catch (e) { Logger.log('QA_getDecedents error: ' + e.message); }
  //try { QA_getShifts(QA_configProperties()); Logger.log('QA_getShifts: done'); } catch (e) { Logger.log('QA_getShifts error: ' + e.message); }
    
  //try { QA_triggerVolunteerShiftRemoval(QA_configProperties()); Logger.log('QA_triggerVolunteerShiftRemoval: done'); } catch (e) { Logger.log('QA_triggerVolunteerShiftRemoval error: ' + e.message); }



}

// Harness wrappers
/**
 * Trigger Updates.
 */
function QA_triggerUpdates(sheetInputs) {
  updateShiftsAndEventMap(sheetInputs);
}


function QA_archiveUpdates(sheetInputs) {
  updateArchive(sheetInputs);
}

function QA_getShifts(sheetInputs){
  const volunteerToken = "39e26527-eff4-4b6e-9e2c-1031a187d7fb";
  const isMember = false;
  const shiftFlags =1;
  const nameOnly = false;

  getShifts(sheetInputs, volunteerToken, isMember, shiftFlags, nameOnly)

}

function QA_getEmails(sheetInputs){
  getEmails(sheetInputs);
}

function QA_getDecedents(sheetInputs){
  getDecedents(sheetInputs);
}

function QA_getLocations(sheetInputs){
  getLocations(sheetInputs);
}


/**
 * Shift Removals
 */
function QA_triggerVolunteerShiftRemoval(sheetInputs) {

  let selectedShiftIds = ["cc4fab4e-a488-4ef7-8ba1-d88ff4c50b24"];
  let volunteerData = {"isMember":true,"name":"Lou Shapiro","email":"eshapiro@gmail.com","token":"6f39a3e5-5e33-4ada-8e3d-6cfc044ac1ba","selectedEvents":[{"End Date":"Thu Nov 20 2025 00:00:00 GMT-0700 (Mountain Standard Time)","Personal Information":"Owner of Boulder's favorite restaurant ","selectedShifts":[{"Current Volunteers":0,"Deceased Name":"Josephine Levin","Event Date":{},"Met or Meta":"meta","Shift Time":"12:00 AM - 1:00 AM","Event Token":"af9762d3-42da-4b27-94cd-ad1d77f664e0","End Epoch":1763452800000,"Max Volunteers":1,"Shift ID":"cc4fab4e-a488-4ef7-8ba1-d88ff4c50b24","Start Epoch":1763449200000,"Location":"Greenwood & Myers Mortuary","Personal Information":"Owner of Boulder's favorite restaurant ","Pronoun":"Her"},{"Max Volunteers":1,"Pronoun":"Her","Deceased Name":"Josephine Levin","Start Epoch":1763492400000,"Location":"Greenwood & Myers Mortuary","Shift Time":"12:00 PM - 1:00 PM","Met or Meta":"meta","Event Date":{},"Shift ID":"5d6d88fd-046e-458c-86e0-012eddbf2263","Personal Information":"Owner of Boulder's favorite restaurant ","Current Volunteers":0,"End Epoch":1763496000000,"Event Token":"af9762d3-42da-4b27-94cd-ad1d77f664e0"},{"End Epoch":1763499600000,"Location":"Greenwood & Myers Mortuary","Personal Information":"Owner of Boulder's favorite restaurant ","Start Epoch":1763496000000,"Pronoun":"Her","Shift ID":"46e4b0ad-6f83-4fe0-974f-1a9d5060f250","Deceased Name":"Josephine Levin","Current Volunteers":0,"Met or Meta":"meta","Event Date":{},"Shift Time":"1:00 PM - 2:00 PM","Event Token":"af9762d3-42da-4b27-94cd-ad1d77f664e0","Max Volunteers":1}],"Location":"Greenwood & Myers Mortuary","availableShifts":[],"End Time":"Sat Dec 30 1899 23:00:00 GMT-0700 (Mountain Standard Time)","Token":"af9762d3-42da-4b27-94cd-ad1d77f664e0","eventShifts":[],"Start Date":"Tue Nov 18 2025 00:00:00 GMT-0700 (Mountain Standard Time)","Pronoun":"Her","Start Time":"Sat Dec 30 1899 09:00:00 GMT-0700 (Mountain Standard Time)","Deceased Name":"Josephine Levin","Timestamp":{},"Met or Meta":"meta","Email Address":"eshapiro@gmail.com"},{"Location":"Crist Mortuary","Met or Meta":"meta","availableShifts":[],"eventShifts":[],"Personal Information":"Dominance","Email Address":"eshapiro@gmail.com","End Date":"Sat Nov 29 2025 00:00:00 GMT-0700 (Mountain Standard Time)","Token":"57bcb924-aa0a-49d2-92b0-556b38276f99","Start Date":"Fri Nov 28 2025 00:00:00 GMT-0700 (Mountain Standard Time)","End Time":"Sat Dec 30 1899 15:30:00 GMT-0700 (Mountain Standard Time)","Timestamp":{},"Pronoun":"Her","Start Time":"Sat Dec 30 1899 14:30:00 GMT-0700 (Mountain Standard Time)","selectedShifts":[],"Deceased Name":"Maryanne Meitzner"}]};



  let response = null;

  try {

    console.log("--- START triggerVolunteerShiftRemoval ---");
    console.log ("selectedShiftIds:" + selectedShiftIds);
    console.log ("volunteerData:" + JSON.stringify(volunteerData));

    if (selectedShiftIds && volunteerData.token) {
      response = removeVolunteerShifts(sheetInputs, selectedShiftIds, volunteerData.name, volunteerData.token);

      if (response) {
          console.log("Volunteer shifts removed: Name:" + volunteerData.name);
          sendShiftEmail(sheetInputs, volunteerData, selectedShiftIds, "Removal");
          return true;
        } else {
          console.log("Volunteer shifts failed to remove: Name:" + volunteerData.name);
          return true;
        }
    }
    else {
        console.log("Volunteer shifts failed to remove: Name:" + volunteerData.name);
        return false;
    }

  } catch (error) {
    console.log("FATAL ERROR in triggerVolunteerShiftRemoval: " + error.toString());
    return null;
  }
  
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

  let sheetInputs = getSheetInputs();
  // now override the ID if necessary
  const ssName = "bck_events_responses-shmira_schedule";
  const ssId = getSpreadsheetByName_(ssName);
  sheetInputs.SPREADSHEET_ID = ssId;

  return sheetInputs;

}

function QA_Logging(logMessage, DEBUG=false) {

  // --- QA CHECK ---
  if (typeof DEBUG === 'undefined' || DEBUG === false) {
    return;
  }
  
  console.log(logMessage);

}
