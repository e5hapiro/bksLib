/**
* -----------------------------------------------------------------
* _selection_form.js
* Chevra Kadisha Selection Form Handler
* -----------------------------------------------------------------
* _selection_form.js
Version: 1.0.3 * Last updated: 2025-11-04
 * 
 * CHANGELOG v1.0.3:
 *   - Initial implementation of Selection Form.
 * -----------------------------------------------------------------
 */


// Used to get specific types of shifts in event information
const SHIFT_FLAGS = {
  NONE: 0,
  AVAILABLE: 1,
  SELECTED: 2,
  EVENT: 4,
};


/**
 * Server-side stub to get available shifts data for "Schedule Shifts" tab.
 */
function getShifts(sheetInputs, volunteerToken, isMember, shiftFlags, nameOnly ) {
  // Fetch and return available shifts data filtered or personalized as needed by volunteerToken
  // Example return format: Array of event objects with availableShifts arrays

  try {

    console.log("--- START getShifts ---");
    console.log("GetShifts-volunteerToken:"+ volunteerToken)

    var info = null;
    var volunteerData = null;

    if (isMember) {

      info = getMemberInfoByToken(sheetInputs, volunteerToken, shiftFlags, nameOnly);

      //logQCVars_('Member info', info);
      if (info) {

        var fullName = info.firstName + " " + info.lastName;
        var email = info.email;

        if(fullName.trim = "") {
          fullName = "Volunteer";      
        }

        volunteerData = {
          name: fullName,
          token: volunteerToken,
          email: email,
          isMember: true,
          events: info.events
          // Add more attributes as needed
        };

        console.log(`Member authenticated: ${volunteerData.name} (Token: ${volunteerData.token.substring(0, 5)}...)`);

        return volunteerData;

      } else {
        console.log("Member token found but failed validation.");
      }
    } else if (volunteerToken) {
      
      info = getGuestInfoByToken(sheetInputs, volunteerToken, shiftFlags, nameOnly );

      var fullName = info.firstName + " " + info.lastName;

      if(fullName.trim = "") {
        fullName = "Volunteer";      
      }

      //logQCVars_('Guest info', info);
      if (info) {

        var fullName = info.firstName + " " + info.lastName;
        var email = info.email;

        if(fullName.trim = "") {
          fullName = "Volunteer";      
        }

        volunteerData = {
          name: fullName,
          token: volunteerToken,
          email: email,
          isMember: false,
          events: info.events
          // Add more attributes as needed
        };

        console.log(`Guest authenticated: ${volunteerData.name} (Token: ${volunteerData.token.substring(0, 5)}...)`);
        return volunteerData;

      } else {
        console.log("Guest token found but failed validation.");
        return null;
      }
    } else {
      console.log("No member or guest token found in URL parameters.");
      return null;
    }

  } catch (error) {
    console.log("FATAL ERROR in getAvailableShifts: " + error.toString());
    return null;
  }

}


function getMemberInfoByToken(sheetInputs, token, shiftFlags , nameOnly, ) {
  Logger.log("Getting getMemberInfoByToken " + token);

  function getSafeValue(row, idx, header) {
    if (idx.hasOwnProperty(header)) {
      return row[idx[header]];
    } else {
      return '';
    }
  }

  try {
    const ss = getSpreadsheet_(sheetInputs.SPREADSHEET_ID);

    const sheet = ss.getSheetByName(sheetInputs.MEMBERS_SHEET);
    if (!sheet) throw new Error(`Sheet not found: ${sheetInputs.MEMBERS_SHEET}`);

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    // Map header names to indices
    var idx = {};
    headers.forEach(function(h, i) { idx[h] = i; });

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      if (normalizeToken(row[idx['Token']]) === normalizeToken(token)) {

        // For bootstrap, only need the name and not all of the other data
        if (nameOnly) {
          const info = {
            firstName: getSafeValue(row, idx, 'First Name'),
            lastName: getSafeValue(row, idx, 'Last Name'),
            email: getSafeValue(row, idx, 'Email Address'),
            events: null
          };
          return info;
        } else {
          const info = {
            timestamp: getSafeValue(row, idx, 'Timestamp'),
            email: getSafeValue(row, idx, 'Email Address'),
            firstName: getSafeValue(row, idx, 'First Name'),
            lastName: getSafeValue(row, idx, 'Last Name'),
            token: getSafeValue(row, idx, 'Token'),
            events: getEventsForToken_(sheetInputs, token, shiftFlags),
            rowIndex: i + 1 // Sheet row (1-based)
          };

          /*
          // Full set of information for debugging only
          const info = {
            timestamp: getSafeValue(row, idx, 'Timestamp'),
            email: getSafeValue(row, idx, 'Email Address'),
            firstName: getSafeValue(row, idx, 'First Name'),
            lastName: getSafeValue(row, idx, 'Last Name'),
            address: getSafeValue(row, idx, 'Address'),
            city: getSafeValue(row, idx, 'City'),
            state: getSafeValue(row, idx, 'State'),
            zip: getSafeValue(row, idx, 'Zip'),
            phone: getSafeValue(row, idx, 'Phone'),
            canText: getSafeValue(row, idx, 'Can we text you at the above phone number?'),
            shmiraVolunteer: getSafeValue(row, idx, 'Are you interested in volunteering to sit shmira?'),
            taharaVolunteer: getSafeValue(row, idx, 'Are you interested in volunteering to do tahara?'),
            tachrichimVolunteer: getSafeValue(row, idx, 'Are you interested in helping to make tachrichim (no sewing experience needed).'),
            hasSatShmiraBoulder: getSafeValue(row, idx, 'Have you sat shmira with the Boulder Chevra Kadisha before?'),
            shmiraTraining: getSafeValue(row, idx, 'Do you need/want training training for sitting shmira'),
            hasSatShmiraOther: getSafeValue(row, idx, 'Have you sat shmira with another chevra kadisha before.'),
            hasDoneTaharaBoulder: getSafeValue(row, idx, 'Have you participated in a tahara with the Boulder Chevra Kadisha?'),
            taharaTraining: getSafeValue(row, idx, 'Do you need/want training training on tahara?'),
            taharaPreference: getSafeValue(row, idx, 'Preferred way to receive requests for tahara (check all that apply)'),
            sewingMachine: getSafeValue(row, idx, 'Do you have a sewing machine?'),
            affiliation: getSafeValue(row, idx, 'What is your affiliation - The Boulder Chevra Kadisha is a community-wide chevra kadisha. We serve all Jews in Boulder County - affiliated of not. '),
            synagogue: getSafeValue(row, idx, 'What is the name of your synagogue (if not a local synagogue also include city, state)'),
            agreement: getSafeValue(row, idx, 'By submitting this application, I certify the information is true and accurate and I agree with the terms and conditions of volunteering with the Boulder Chevra Kadisha. '),
            communicationPreference: getSafeValue(row, idx, 'Preferred way to receive communication (including shmira and tahara notifications) (check all that apply)'),
            hasSatShmiraOtherAgain: getSafeValue(row, idx, 'Have you sat shmira with another chevra kadisha before?'),
            hasDoneTaharaOther: getSafeValue(row, idx, 'Have you participated in a tahara with another chevra kadisha?'),
            notes: getSafeValue(row, idx, 'Is there anything you want us to know about you, your skills or past chevra kadisha experience?'),
            token: getSafeValue(row, idx, 'Token'),
            approvals: getSafeValue(row, idx, 'Approvals'),
            events: getEventsForToken_(sheetInputs, token),
            rowIndex: i + 1 // Sheet row (1-based)
          };
          */

          //logQCVars_('getMemberInfoByToken_', info);
          return info;
        } // Closing else
      }
    }
    Logger.log("Did not find a member with that token");
    return null;
  } catch (e) {
    Logger.log("Error in getMemberInfoByToken: " + e.toString());
    throw e;
  }
}



function getGuestInfoByToken(sheetInputs, token, shiftFlags, nameOnly, ) {
  Logger.log("Getting getGuestInfoByToken[token]: " + token);

  function getSafeValue(row, idx, header) {
    if (idx.hasOwnProperty(header)) {
      return row[idx[header]];
    } else {
      return '';
    }
  }

  try {
    // The master workbook
    const ss = getSpreadsheet_(sheetInputs.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetInputs.GUESTS_SHEET);
    if (!sheet) throw new Error(`Sheet not found: ${sheetInputs.GUESTS_SHEET}`);

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    var info = null;

    // Map header names to indices
    var idx = {};
    headers.forEach(function(h, i) { idx[h] = i; });

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (normalizeToken(row[idx['Token']]) === normalizeToken(token)) {

        // For bootstrap, only need the name and not all of the other data
        if (nameOnly) {

          info = {
            firstName: getSafeValue(row, idx, 'First Name'),
            lastName: getSafeValue(row, idx, 'Last Name'),
            email: getSafeValue(row, idx, 'Email Address'),
            events: null
          }
          return info;

        } 
        else {

          info = {
            timestamp: getSafeValue(row, idx, 'Timestamp'),
            email: getSafeValue(row, idx, 'Email Address'),
            firstName: getSafeValue(row, idx, 'First Name'),
            lastName: getSafeValue(row, idx, 'Last Name'),
            token: getSafeValue(row, idx, 'Token'),
            events: getEventsForToken_(sheetInputs, token, shiftFlags),
            rowIndex: i + 1 // Sheet row (1-based)
          };
          
          
          // Full set of information for debugging only
          /*
          info = {
            timestamp: getSafeValue(row, idx, 'Timestamp'),
            email: getSafeValue(row, idx, 'Email Address'),
            firstName: getSafeValue(row, idx, 'First Name'),
            lastName: getSafeValue(row, idx, 'Last Name'),
            address: getSafeValue(row, idx, 'Address'),
            city: getSafeValue(row, idx, 'City'),
            state: getSafeValue(row, idx, 'State'),
            zip: getSafeValue(row, idx, 'Zip'),
            phone: getSafeValue(row, idx, 'Phone'),
            canText: getSafeValue(row, idx, 'Can we text you at the above phone number?'),
            names: (getSafeValue(row, idx, 'Name of Deceased') || "")
              .toString()
              .split(',')
              .map(function(n) { return n.trim().toLowerCase(); }),
            relationship: getSafeValue(row, idx, 'Relationship to Deceased'),
            over18: getSafeValue(row, idx, 'Are you over 18 years old?'),
            sitAlone: getSafeValue(row, idx, "To sit shmira alone with the Boulder Chevra Kadisha, you must be over 18 years old. If you are under 18 years old, you can sit shmira with a Boulder Chevra Kadisha Member or a parent/guardian. If you will sit shmira with a parent or guardian, have them fill out the form. If you would like to be matched up with a Member of the Boulder Chevra Kadisha, you can continue to complete the form."),
            canSitDuringBusiness: getSafeValue(row, idx, "Are you able to sit shmira during the mortuary's normal business hours? (Business hours are Monday through Friday 9am - 5pm)"),
            sitAfterHours: getSafeValue(row, idx, "To sit shmira alone after business hours of the mortuary, you must be a Boulder Chevra Kadisha Member. If you are not a member and still want to sit shmira and you are only able to sit shmira outside of regular business hours, we can match you with a member. Would you like to discuss sitting shmira with a Boulder Chevra Kadisha Member?"),
            affiliation: getSafeValue(row, idx, "What is your affiliation? (The Boulder Chevra Kadisha is a community wide chevra kadisha. We serve all Jews in Boulder County - affiliated of not.)"),
            synagogue: getSafeValue(row, idx, 'Name, City and State of synagogue.'),
            onMailingList: getSafeValue(row, idx, 'Do you want to be on our mailing list for events and training?'),
            agreement: getSafeValue(row, idx, 'By submitting this application, I certify the information is true and accurate and I agree with the terms and conditions of sitting shmira with the Boulder Chevra Kadisha.'),
            token: getSafeValue(row, idx, 'Token'),
            approvals: getSafeValue(row, idx, 'Approvals'),
            events: getEventsForToken_(sheetInputs, token),
            rowIndex: i + 1 // Sheet row (1-based)
          };
          */

          //logQCVars_('getGuestInfoByToken', info);
          return info;
        }
      }

    }
    Logger.log("Did not find a guest with that token");
    return null;
  } catch (e) {
    Logger.log("Error in getGuestInfoByToken: " + e.toString());
    throw e;
  }
}


// SHARED event look up for both guests and members
function getEventsForToken_(sheetInputs, guestOrMemberToken, shiftFlags = 0) {    // The master workbook
  Logger.log("Getting getEventsForToken_[guestOrMemberToken]: " + guestOrMemberToken);

  // FIELD NAMES TO NORMALIZE
  const dateFields = ["Start Date", "Start Time", "End Date", "End Time"];
  const ss = getSpreadsheet_(sheetInputs.SPREADSHEET_ID);

  // The map sheet
  const eventMapSheet = ss.getSheetByName(sheetInputs.EVENT_MAP);
  if (!eventMapSheet) throw new Error(`Sheet not found: ${sheetInputs.EVENT_MAP}`);

    // The events sheet
    const eventSheet = ss.getSheetByName(sheetInputs.EVENT_FORM_RESPONSES);
    if (!eventSheet) throw new Error(`Sheet not found: ${sheetInputs.EVENT_FORM_RESPONSES}`);


  const eventMapData = eventMapSheet.getDataRange().getValues();
  const eventTokenColIdx = eventMapData[0].indexOf("Event Token");
  const guestMemberTokenColIdx = eventMapData[0].indexOf("Guest/Member Token");
  // Find all Event Tokens for this token
  var matchedEventTokens = [];
  for (let i = 1; i < eventMapData.length; i++) {
    if (normalizeToken(eventMapData[i][guestMemberTokenColIdx]) === normalizeToken(guestOrMemberToken)) {
      matchedEventTokens.push(eventMapData[i][eventTokenColIdx]);
    }
  }

  // Now collect ALL Form Responses 1 events that have a Token matching any event token
  const eventData = eventSheet.getDataRange().getValues();
  const eventHeaders = eventData[0];
  const eventTokenIdx = eventHeaders.indexOf("Token");
  var events = [];

  for (let i = 1; i < eventData.length; i++) {
    if (matchedEventTokens.indexOf(eventData[i][eventTokenIdx]) > -1) {
      var eventObj = {};
      for (var col = 0; col < eventHeaders.length; col++) {
        eventObj[eventHeaders[col]] = eventData[i][col];
      }
      // --- Normalize date fields to string ---
      dateFields.forEach(field => {
        if (
          eventObj[field] != null &&
          typeof eventObj[field] === "object" &&
          eventObj[field].toString
        ) {
          eventObj[field] = eventObj[field].toString();
        }
      });

      console.log("shiftFlags= " + shiftFlags);

      const eventToken = eventObj['Token'];

      if ((shiftFlags & SHIFT_FLAGS.AVAILABLE) !== 0) {
          eventObj.availableShifts = getAvailableShiftsForEvent(sheetInputs, eventToken);
        } else {
          eventObj.availableShifts = null;
        }

        if ((shiftFlags & SHIFT_FLAGS.SELECTED) !== 0) {
          eventObj.selectedShifts = getSelectedShiftsByToken(sheetInputs, eventToken, guestOrMemberToken);
        } else {
          eventObj.selectedShifts = null;
        }

        if ((shiftFlags & SHIFT_FLAGS.EVENT) !== 0) {
          eventObj.eventShifts = getEventShifts(sheetInputs, eventToken);
        } else {
          eventObj.eventShifts = null;
        }

      events.push(eventObj);
    }
  }

  //logQCVars_("getEventsForToken_.events", events);

  return events;
}

function getAvailableShiftsForEvent(sheetInputs, eventToken) {
  Logger.log("Getting getAvailableShiftsForEvent[eventToken]: " + eventToken);

  const ss = getSpreadsheet_(sheetInputs.SPREADSHEET_ID);

  // The shifts master sheet
  const shiftsSheet = ss.getSheetByName(sheetInputs.SHIFTS_MASTER_SHEET);
  if (!shiftsSheet) throw new Error(`Sheet not found: ${sheetInputs.SHIFTS_MASTER_SHEET}`);

  // The shifts master sheet
  const volunteerShiftsSheet = ss.getSheetByName(sheetInputs.VOLUNTEER_LIST_SHEET);
  if (!volunteerShiftsSheet) throw new Error(`Sheet not found: ${sheetInputs.VOLUNTEER_LIST_SHEET}`);


  const shiftsData = shiftsSheet.getDataRange().getValues();
  const shiftsHeaders = shiftsData[0];
  const eventTokenIdx = shiftsHeaders.indexOf("Event Token");
  const shiftIdIdx = shiftsHeaders.indexOf("Shift ID");


  // List ONLY non-duplicated, shift-specific columns to keep
  const keepColumns = ["Shift ID", "Event Token", "Shift Time", "Start Epoch", "End Epoch"];
  const keepIdxs = keepColumns.map(header => shiftsHeaders.indexOf(header));

  let eventShifts = [];
  for (let i = 1; i < shiftsData.length; i++) {
    if (shiftsData[i][eventTokenIdx] === eventToken) {
      let shiftObj = {};
      for (let k = 0; k < keepColumns.length; k++) {
        shiftObj[keepColumns[k]] = shiftsData[i][keepIdxs[k]];
      }
      eventShifts.push(shiftObj);
    }
  }
  // Find claimed shifts
  const volunteerData = volunteerShiftsSheet.getDataRange().getValues();
  const volunteerHeaders = volunteerData[0];
  const volunteerShiftIdIdx = volunteerHeaders.indexOf("Shift ID");
  const claimedShiftIds = new Set();
  for (let i = 1; i < volunteerData.length; i++) {
    claimedShiftIds.add(volunteerData[i][volunteerShiftIdIdx]);
  }

  const availableShifts = eventShifts.filter(shift => !claimedShiftIds.has(shift[shiftsHeaders[shiftIdIdx]]));

  //logQCVars_("AvailableShifts:", availableShifts);

  // Return only shifts that are NOT already claimed
  return availableShifts;
}

function getSelectedShiftsByToken(sheetInputs, eventToken, userToken) {
  Logger.log("Getting getSelectedShiftsByToken[eventToken]: " + eventToken);

  const ss = getSpreadsheet_(sheetInputs.SPREADSHEET_ID);

  const shiftsSheet = ss.getSheetByName(sheetInputs.SHIFTS_MASTER_SHEET);
  if (!shiftsSheet) throw new Error(`Sheet not found: ${sheetInputs.SHIFTS_MASTER_SHEET}`);

  const volunteerShiftsSheet = ss.getSheetByName(sheetInputs.VOLUNTEER_LIST_SHEET);
  if (!volunteerShiftsSheet) throw new Error(`Sheet not found: ${sheetInputs.VOLUNTEER_LIST_SHEET}`);

  // Get Volunteer Shifts data
  const volunteerData = volunteerShiftsSheet.getDataRange().getValues();
  const volunteerHeaders = volunteerData[0];
  const shiftIdIdx = volunteerHeaders.indexOf("Shift ID");
  const userTokenIdx = volunteerHeaders.indexOf("Volunteer Token");

  // Get all shift IDs for this user
  const userShiftIds = new Set();
  for (let i = 1; i < volunteerData.length; i++) {
      if (volunteerData[i][userTokenIdx] === userToken) {
          userShiftIds.add(volunteerData[i][shiftIdIdx]);
      }
  }

  // Get all shift info for this event token
  const shiftsData = shiftsSheet.getDataRange().getValues();
  const shiftsHeaders = shiftsData[0];
  const shiftIdCol = shiftsHeaders.indexOf("Shift ID");
  const eventTokenCol = shiftsHeaders.indexOf("Event Token");

  let selectedShifts = [];
  for (let i = 1; i < shiftsData.length; i++) {
      // Only include a shift if it matches both user and event
      if (
          shiftsData[i][eventTokenCol] === eventToken &&
          userShiftIds.has(shiftsData[i][shiftIdCol])
      ) {
          let shiftObj = {};
          for (let j = 0; j < shiftsHeaders.length; j++) {
              shiftObj[shiftsHeaders[j]] = shiftsData[i][j];
          }
          selectedShifts.push(shiftObj);
      }
  }
  return selectedShifts;
}



function getEventShifts(sheetInputs, eventToken) {
  Logger.log("Getting getEventShifts[eventToken]: " + eventToken);  
  const ss = getSpreadsheet_(sheetInputs.SPREADSHEET_ID);

  const shiftsSheet = ss.getSheetByName(sheetInputs.SHIFTS_MASTER_SHEET);
  if (!shiftsSheet) throw new Error(`Sheet not found: ${sheetInputs.SHIFTS_MASTER_SHEET}`);

  // Get all shift info for this event token
  const shiftsData = shiftsSheet.getDataRange().getValues();
  const shiftsHeaders = shiftsData[0];
  const eventTokenCol = shiftsHeaders.indexOf("Event Token");

  let eventShifts = [];
  for (let i = 1; i < shiftsData.length; i++) {
      // Only include a shift if it matches both user and event
      if (
          shiftsData[i][eventTokenCol] === eventToken 
      ) {
          let shiftObj = {};
          for (let j = 0; j < shiftsHeaders.length; j++) {
              shiftObj[shiftsHeaders[j]] = shiftsData[i][j];
          }
          eventShifts.push(shiftObj);
      }
  }
  return eventShifts;
}



function setVolunteerShifts(sheetInputs, selectedShiftIds, volunteerName, volunteerToken) {
  Logger.log("Getting setVolunteerShifts[volunteerToken]: " + volunteerToken);  

  Logger.log("setVolunteerShifts:");
  Logger.log("selectedShiftIds:");
  Logger.log(selectedShiftIds);
  Logger.log("volunteerName:");
  Logger.log(volunteerName);
  Logger.log("volunteerToken:");
  Logger.log(volunteerToken);

  try {
 
   const ss = getSpreadsheet_(sheetInputs.SPREADSHEET_ID);

    // The volunteer shifts sheet
    const volunteerShiftsSheet = ss.getSheetByName(sheetInputs.VOLUNTEER_LIST_SHEET);
    if (!volunteerShiftsSheet) throw new Error(`Sheet not found: ${sheetInputs.VOLUNTEER_LIST_SHEET}`);

    const now = new Date();

    // Get headers and indices
    const data = volunteerShiftsSheet.getDataRange().getValues();
    const headers = data[0];
    const shiftIdIdx = headers.indexOf("Shift ID");
    const volunteerTokenIdx = headers.indexOf("Volunteer Token");

    // To ensure idempotency, collect all assigned (Shift ID, Volunteer Token) pairs
    const assigned = new Set();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      assigned.add(row[shiftIdIdx] + ":" + row[volunteerTokenIdx]);
    }

    let addedCount = 0;
    selectedShiftIds.forEach(shiftId => {
      // Avoid duplicate assignment
      if (assigned.has(shiftId + ":" + volunteerToken)) return;

      var newRow = [
        Utilities.formatDate(now, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm:ss'),
        shiftId,
        volunteerToken,
        volunteerName
      ];

      //Logger.log(newRow);

      volunteerShiftsSheet.appendRow(newRow);

      assigned.add(shiftId + ":" + volunteerToken);
      addedCount += 1;
    });

    if (addedCount === 0) {
      return "All selected shifts were already assigned to you or unavailable.";
    }
    return true; // Indicate success for frontend callback

  } catch (e) {
    Logger.log("Error in setVolunteerShifts: " + e.message);
    throw e;
  }
}


function removeVolunteerShifts(sheetInputs, shiftIds, volunteerName, volunteerToken) {
  Logger.log("Getting removeVolunteerShifts[volunteerToken]: " + volunteerToken);  

  Logger.log("bckLib.setVolunteerShifts:");
  Logger.log("bckLib.shiftIds:");
  Logger.log(shiftIds);
  Logger.log("bckLib.volunteerToken:");
  Logger.log(volunteerToken);
  Logger.log("bckLib.volunteerName:");
  Logger.log(volunteerName);

  try {
 
   const ss = getSpreadsheet_(sheetInputs.SPREADSHEET_ID);

    // The volunteer shifts sheet
    const volunteerShiftsSheet = ss.getSheetByName(sheetInputs.VOLUNTEER_LIST_SHEET);
    if (!volunteerShiftsSheet) throw new Error(`Sheet not found: ${sheetInputs.VOLUNTEER_LIST_SHEET}`);

    const data = volunteerShiftsSheet.getDataRange().getValues();
    const headers = data[0];
    const shiftIdIdx = headers.indexOf("Shift ID");
    const volunteerTokenIdx = headers.indexOf("Volunteer Token");

    // Gather the rows to delete (row numbers in Apps Script are 1-based!)
    let rowsToDelete = [];
    for (let i = 1; i < data.length; i++) {
      if (shiftIds.includes(data[i][shiftIdIdx]) && data[i][volunteerTokenIdx] === volunteerToken) {
        rowsToDelete.push(i + 1); // add 1 for 1-based row index + 1 for header
      }
    }

    // Delete from bottom up to preserve indexes
    rowsToDelete.reverse().forEach(rowNum => {
      volunteerShiftsSheet.deleteRow(rowNum);
    });

  } catch (e) {
    Logger.log("Error in removeVolunteerShifts: " + e.message);
    throw e;
  }


  return true;
}


function sendShiftEmail(sheetInputs, volunteerData, shifts, actionType) {

  console.log("--- START sendShiftEmail ---");
  console.log("volunteerData:" + JSON.stringify(volunteerData));
  console.log("shifts (IDs):" + JSON.stringify(shifts));

  const emailAction = actionType === "Addition" ? "added to" : "removed from";

  const urlParam = volunteerData.isMember ? "m" : "g";
  const sheetUrl = sheetInputs["SCRIPT_URL"];
  const personalizedUrl = `${sheetUrl}?${urlParam}=${volunteerData.token}`;
  const recipientEmail = volunteerData.email;
  //      const info = bckLib.getMemberInfoByToken(sheetInputs, volunteerData.token, nameOnly);
  const nameOnly = false;
  const info = getMemberInfoByToken(sheetInputs, volunteerData.token, SHIFT_FLAGS.EVENT, nameOnly  );

  const events = info.events || [];
  // Aggregate all available shifts from all events (since each event has an array of availableShifts)
  const allAvailableShifts = events.flatMap(event => {
    return (event.eventShifts || []).map(shift => {
      // Include event info in the shift for email detail composition
      return {
        ...shift,
        eventLocation: event.Location,
        deceasedName: event["Deceased Name"],
        eventStartDate: event["Start Date"],
        eventToken: event.Token || event["Token"] || event["Event Token"] || event.Token,
      };
    });
  });

  // Filter to get details for only those shift IDs passed in the shifts array
  const matchedShiftDetails = allAvailableShifts.filter(shift => shifts.includes(shift["Shift ID"]));

  if (matchedShiftDetails.length === 0) {
    console.warn("No matching shift details found for provided shift IDs");
  }

  // Compose subject line based on matched shifts and using shift times or deceased names for clarity
  const subjectPrefix = "Chevra Kadisha Volunteer: Confirmation of Shift " + actionType
  const subjectSuffix = "at " + `${matchedShiftDetails.map(s => s["Location"]).join(', ')}`

  const subject = `${subjectPrefix}: Shift${matchedShiftDetails.length > 1 ? 's' : ''} ${matchedShiftDetails.map(s => s["Shift Time"]).join(', ')} ${subjectSuffix}`;
  console.log("subject: " + subject);

  // Build the body with all matched shift details
  let shiftDetails = '';
  const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };

  matchedShiftDetails.forEach(shift => {
    console.log("processing matched shift: " + JSON.stringify(shift));
    const fullAddress = getAddressFromLocationName_(shift.eventLocation);

    // Format the event date neatly (assumes eventStartDate as ISO string or date string)
    //const eventDateFormatted = shift.eventStartDate ? new Date(shift.eventStartDate).toLocaleDateString() : 'Date Unknown';
    let eventDateFormatted = 'Date Unknown';

    if (shift.eventStartDate) {
      const date = new Date(shift.eventStartDate);
      if (!isNaN(date)) {
        eventDateFormatted = date.toLocaleDateString('en-US', options);
      }
    }

    shiftDetails += `
    - Deceased: ${shift.deceasedName || 'N/A'}
    - Location: ${shift.eventLocation}
    - Address: ${fullAddress}
    - Date: ${eventDateFormatted}
    - Time: ${shift["Shift Time"]}
    `;
  });

  const body = `
    Dear ${volunteerData.name},

    This is an automatic confirmation that your request to be ${emailAction} the following shift${matchedShiftDetails.length > 1 ? 's' : ''} has been processed successfully:

    Shift Details:
    ${shiftDetails}

    If you need to cancel or change your confirmation, go to Your Volunteer Portal Link: ${personalizedUrl}. Remember, this link is unique to you. Please do not share it.

    Thank you for providing this mitzvah.
  `;

  console.log("body: " + body);

  try {
    if (!recipientEmail || !String(recipientEmail).includes('@')) {
      Logger.log(`Skipping email: Invalid recipient email address: ${recipientEmail}`);
      return;
    }

    MailApp.sendEmail({
      to: recipientEmail,
      subject: subject,
      body: body
    });
    Logger.log(`Email sent successfully for ${actionType} to ${recipientEmail}`);
  } catch (e) {
    Logger.log(`ERROR sending email for ${actionType}: ${e.toString()}`);
  }
}