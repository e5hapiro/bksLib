/**
 * -----------------------------------------------------------------
 * _selection_form.js
 * Chevra Kadisha Selection Form Handler
 * -----------------------------------------------------------------
 * _selection_form.js
 * Version: 1.0.8 
 * Last updated: 2025-12-22
 * 
 * CHANGELOG v1.0.3:
 *   - Initial implementation of Selection Form.
 *   v1.0.6:
 *   - Fixed bug in usage of DEBUG
 *   v1.0.7
 *   - Fixed regression in that location needed to be found in sendEmails for removal and addition
 *   - Added ics mail attachments to emails
 *   v1.0.8
 *   - Filtered available events and shifts to only those that are relevent based on date/time 
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
    Logger.log("Getting getShifts " + volunteerToken);

  // Fetch and return available shifts data filtered or personalized as needed by volunteerToken
  // Example return format: Array of event objects with availableShifts arrays

  if (typeof sheetInputs.DEBUG === 'undefined') {
    console.log ("DEBUG is undefined");
    return;
  }

  try {

    console.log("--- START getShifts ---");
    console.log("GetShifts-volunteerToken:"+ volunteerToken)

    var info = null;
    var volunteerData = null;

    if (isMember) {

      info = getMemberInfoByToken(sheetInputs, volunteerToken, shiftFlags, nameOnly);

      //logQCVars('Member info', info);
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

      //logQCVars('Guest info', info);
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

  if (typeof sheetInputs.DEBUG === 'undefined') {
    console.log ("DEBUG is undefined");
    return;
  }


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

          //logQCVars('getMemberInfoByToken_', info);
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

  if (typeof sheetInputs.DEBUG === 'undefined') {
    console.log ("DEBUG is undefined");
    return;
  }

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

          //logQCVars('getGuestInfoByToken', info);
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

function getEventsForToken_(sheetInputs, guestOrMemberToken, shiftFlags = 0) {
  Logger.log("getEventsForToken_ START, guestOrMemberToken=%s, shiftFlags=%s", guestOrMemberToken, shiftFlags);
  console.log("getEventsForToken_ START", { guestOrMemberToken, shiftFlags });

  if (typeof sheetInputs.DEBUG === 'undefined') {
    console.log("DEBUG is undefined");
    return;
  }

  const dateFields = ["Start Date", "Start Time", "End Date", "End Time"];
  const ss = getSpreadsheet_(sheetInputs.SPREADSHEET_ID);

  const eventMapSheet = ss.getSheetByName(sheetInputs.EVENT_MAP);
  if (!eventMapSheet) throw new Error(`Sheet not found: ${sheetInputs.EVENT_MAP}`);

  const eventSheet = ss.getSheetByName(sheetInputs.EVENT_FORM_RESPONSES);
  if (!eventSheet) throw new Error(`Sheet not found: ${sheetInputs.EVENT_FORM_RESPONSES}`);

  const eventMapData = eventMapSheet.getDataRange().getValues();
  const eventTokenColIdx = eventMapData[0].indexOf("Event Token");
  const guestMemberTokenColIdx = eventMapData[0].indexOf("Guest/Member Token");

  // Collect all event tokens for this guest/member
  const matchedEventTokens = [];
  for (let i = 1; i < eventMapData.length; i++) {
    if (normalizeToken(eventMapData[i][guestMemberTokenColIdx]) === normalizeToken(guestOrMemberToken)) {
      matchedEventTokens.push(eventMapData[i][eventTokenColIdx]);
    }
  }
  console.log("matchedEventTokens:", matchedEventTokens);

  const eventData = eventSheet.getDataRange().getValues();
  const eventHeaders = eventData[0];
  const eventTokenIdx = eventHeaders.indexOf("Token");

  const events = [];
  const now = new Date();

  const tz = Session.getScriptTimeZone(); // for Utilities.formatDate [web:94]

  for (let i = 1; i < eventData.length; i++) {
    if (matchedEventTokens.indexOf(eventData[i][eventTokenIdx]) > -1) {
      const eventObj = {};
      for (let col = 0; col < eventHeaders.length; col++) {
        eventObj[eventHeaders[col]] = eventData[i][col];
      }

      // Raw values coming from the sheet (usually Date objects)
      const rawStartDate = eventObj["Start Date"];
      const rawEndDate   = eventObj["End Date"];
      const rawEndTime   = eventObj["End Time"];

      // Build Date objects for filtering (do NOT overwrite original fields)
      const endDateVal = rawEndDate instanceof Date ? rawEndDate : new Date(rawEndDate);
      const endTimeVal = rawEndTime instanceof Date ? rawEndTime : new Date(rawEndTime);

      if (!(endDateVal instanceof Date) || isNaN(endDateVal.getTime()) ||
          !(endTimeVal instanceof Date) || isNaN(endTimeVal.getTime())) {
        console.warn("Skipping event due to invalid End Date/Time for filtering", {
          endDateVal, endTimeVal
        });
        continue;
      }

      const eventEnd = new Date(
        endDateVal.getFullYear(),
        endDateVal.getMonth(),
        endDateVal.getDate(),
        endTimeVal.getHours(),
        endTimeVal.getMinutes(),
        endTimeVal.getSeconds()
      );

      // Skip this event if parsing failed or it already ended
      if (!(eventEnd instanceof Date) || isNaN(eventEnd.getTime()) || eventEnd <= now) {
        console.log("Skipping past/invalid event");
        continue;
      }

      // ---- add display strings for client ----
      // Example format: Monday Dec 22, 2025
      var displayPattern = "EEEE MMM d, yyyy";

      if (rawStartDate instanceof Date) {
        eventObj.startDateDisplay = Utilities.formatDate(rawStartDate, tz, displayPattern);
      } else if (rawStartDate) {
        eventObj.startDateDisplay = String(rawStartDate);
      } else {
        eventObj.startDateDisplay = "";
      }

      if (rawEndDate instanceof Date) {
        eventObj.endDateDisplay = Utilities.formatDate(rawEndDate, tz, displayPattern);
      } else if (rawEndDate) {
        eventObj.endDateDisplay = String(rawEndDate);
      } else {
        eventObj.endDateDisplay = "";
      }

      const eventToken = eventObj["Token"];

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

  console.log("getEventsForToken_ END, events length:", events.length);
  return events;
}
function getAvailableShiftsForEvent(sheetInputs, eventToken) {
  Logger.log("getAvailableShiftsForEvent START, eventToken=%s", eventToken);
  console.log("getAvailableShiftsForEvent START", { eventToken });

  if (typeof sheetInputs.DEBUG === "undefined") {
    console.log("DEBUG is undefined");
    Logger.log("DEBUG is undefined");
    return;
  }

  const ss = getSpreadsheet_(sheetInputs.SPREADSHEET_ID);

  const shiftsSheet = ss.getSheetByName(sheetInputs.SHIFTS_MASTER_SHEET);
  if (!shiftsSheet) {
    console.error(`Sheet not found: ${sheetInputs.SHIFTS_MASTER_SHEET}`);
    Logger.log("ERROR: Sheet not found: %s", sheetInputs.SHIFTS_MASTER_SHEET);
    throw new Error(`Sheet not found: ${sheetInputs.SHIFTS_MASTER_SHEET}`);
  }

  const volunteerShiftsSheet = ss.getSheetByName(sheetInputs.VOLUNTEER_LIST_SHEET);
  if (!volunteerShiftsSheet) {
    console.error(`Sheet not found: ${sheetInputs.VOLUNTEER_LIST_SHEET}`);
    Logger.log("ERROR: Sheet not found: %s", sheetInputs.VOLUNTEER_LIST_SHEET);
    throw new Error(`Sheet not found: ${sheetInputs.VOLUNTEER_LIST_SHEET}`);
  }

  const shiftsData = shiftsSheet.getDataRange().getValues();
  const shiftsHeaders = shiftsData[0];
  const eventTokenIdx = shiftsHeaders.indexOf("Event Token");
  const shiftIdIdx = shiftsHeaders.indexOf("Shift ID");
  const startEpochIdx = shiftsHeaders.indexOf("Start Epoch");
  const endEpochIdx = shiftsHeaders.indexOf("End Epoch");

  const keepColumns = ["Shift ID", "Event Token", "Shift Time", "Start Epoch", "End Epoch"];
  const keepIdxs = keepColumns.map(header => shiftsHeaders.indexOf(header));

  const now = Date.now();
  const tz = Session.getScriptTimeZone(); // for formatting [web:12][web:94]

  let eventShifts = [];
  for (let i = 1; i < shiftsData.length; i++) {
    const row = shiftsData[i];
    const rowEventToken = row[eventTokenIdx];

    const startEpochRaw = row[startEpochIdx];
    const endEpochRaw   = row[endEpochIdx];
    const startEpoch = Number(startEpochRaw);
    const endEpoch   = Number(endEpochRaw);

    if (rowEventToken !== eventToken) continue;
    if (isNaN(startEpoch) || isNaN(endEpoch)) continue;

    // Only skip if shift end is before now
    if (now > endEpoch) continue;

    let shiftObj = {};
    for (let k = 0; k < keepColumns.length; k++) {
      shiftObj[keepColumns[k]] = row[keepIdxs[k]];
    }

    // NEW: human-friendly display string: "Dec 22, 2025 9:00 PM - 10:00 PM"
    const startDateObj = new Date(startEpoch);
    const endDateObj   = new Date(endEpoch);

    if (!isNaN(startDateObj.getTime()) && !isNaN(endDateObj.getTime())) {
      const datePart      = Utilities.formatDate(startDateObj, tz, "MMM d, yyyy");
      const startTimePart = Utilities.formatDate(startDateObj, tz, "h:mm a");
      const endTimePart   = Utilities.formatDate(endDateObj,   tz, "h:mm a");
      shiftObj["Shift Display"] = `${datePart} ${startTimePart} - ${endTimePart}`;
    } else {
      shiftObj["Shift Display"] = shiftObj["Shift Time"] || "";
    }

    eventShifts.push(shiftObj);
  }

  // Find claimed shifts
  const volunteerData = volunteerShiftsSheet.getDataRange().getValues();
  const volunteerHeaders = volunteerData[0];
  const volunteerShiftIdIdx = volunteerHeaders.indexOf("Shift ID");

  const claimedShiftIds = new Set();
  for (let i = 1; i < volunteerData.length; i++) {
    const claimedId = volunteerData[i][volunteerShiftIdIdx];
    if (claimedId) claimedShiftIds.add(claimedId);
  }

  let availableShifts;
  if (eventShifts.length === 0) {
    availableShifts = [];
  } else {
    availableShifts = eventShifts.filter(shift => !claimedShiftIds.has(shift["Shift ID"]));
  }

  return availableShifts;
}



function getSelectedShiftsByToken(sheetInputs, eventToken, userToken) {
  Logger.log("Getting getSelectedShiftsByToken[eventToken]: " + eventToken);

  if (typeof sheetInputs.DEBUG === 'undefined') {
    console.log("DEBUG is undefined");
    return;
  }

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
  const startEpochIdx = shiftsHeaders.indexOf("Start Epoch");
  const endEpochIdx   = shiftsHeaders.indexOf("End Epoch");

  const tz = Session.getScriptTimeZone(); // for Utilities.formatDate [web:12][web:94]

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

      // Add formatted display string: "Dec 22, 2025 9:00 PM - 10:00 PM"
      const startEpoch = Number(shiftsData[i][startEpochIdx]);
      const endEpoch   = Number(shiftsData[i][endEpochIdx]);

      if (!isNaN(startEpoch) && !isNaN(endEpoch)) {
        const startDateObj = new Date(startEpoch);
        const endDateObj   = new Date(endEpoch);

        if (!isNaN(startDateObj.getTime()) && !isNaN(endDateObj.getTime())) {
          const datePart      = Utilities.formatDate(startDateObj, tz, "MMM d, yyyy");
          const startTimePart = Utilities.formatDate(startDateObj, tz, "h:mm a");
          const endTimePart   = Utilities.formatDate(endDateObj,   tz, "h:mm a");
          shiftObj["Shift Display"] = `${datePart} ${startTimePart} - ${endTimePart}`;
        }
      }

      selectedShifts.push(shiftObj);
    }
  }
  return selectedShifts;
}

function getEventShifts(sheetInputs, eventToken) {
  Logger.log("Getting getEventShifts[eventToken]: " + eventToken);  

  if (typeof sheetInputs.DEBUG === 'undefined') {
    console.log("DEBUG is undefined");
    return;
  }

  const ss = getSpreadsheet_(sheetInputs.SPREADSHEET_ID);
  const shiftsSheet = ss.getSheetByName(sheetInputs.SHIFTS_MASTER_SHEET);
  if (!shiftsSheet) throw new Error(`Sheet not found: ${sheetInputs.SHIFTS_MASTER_SHEET}`);

  const shiftsData = shiftsSheet.getDataRange().getValues();
  const shiftsHeaders = shiftsData[0];

  const eventTokenCol  = shiftsHeaders.indexOf("Event Token");
  const startEpochIdx  = shiftsHeaders.indexOf("Start Epoch");
  const endEpochIdx    = shiftsHeaders.indexOf("End Epoch");

  const tz = Session.getScriptTimeZone(); // for Utilities.formatDate

  let eventShifts = [];
  for (let i = 1; i < shiftsData.length; i++) {
    if (shiftsData[i][eventTokenCol] === eventToken) {
      let shiftObj = {};
      for (let j = 0; j < shiftsHeaders.length; j++) {
        shiftObj[shiftsHeaders[j]] = shiftsData[i][j];
      }

      // Add formatted display string: "Dec 22, 2025 9:00 PM - 10:00 PM"
      const startEpoch = Number(shiftsData[i][startEpochIdx]);
      const endEpoch   = Number(shiftsData[i][endEpochIdx]);

      if (!isNaN(startEpoch) && !isNaN(endEpoch)) {
        const startDateObj = new Date(startEpoch);
        const endDateObj   = new Date(endEpoch);

        if (!isNaN(startDateObj.getTime()) && !isNaN(endDateObj.getTime())) {
          const datePart      = Utilities.formatDate(startDateObj, tz, "MMM d, yyyy");
          const startTimePart = Utilities.formatDate(startDateObj, tz, "h:mm a");
          const endTimePart   = Utilities.formatDate(endDateObj,   tz, "h:mm a");
          shiftObj["Shift Display"] = `${datePart} ${startTimePart} - ${endTimePart}`;
        }
      }

      eventShifts.push(shiftObj);
    }
  }
  return eventShifts;
}


function setVolunteerShifts(sheetInputs, selectedShiftIds, volunteerName, volunteerToken) {
  Logger.log("Getting setVolunteerShifts[volunteerToken]: " + volunteerToken);  

  if (typeof sheetInputs.DEBUG === 'undefined') {
    console.log ("DEBUG is undefined");
    return;
  }


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

  if (typeof sheetInputs.DEBUG === 'undefined') {
    console.log ("DEBUG is undefined");
    return;
  }


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

  if (typeof sheetInputs.DEBUG === 'undefined') {
    console.log ("DEBUG is undefined");
    return;
  }

  const emailAction = actionType === "Addition" ? "added to" : "removed from";
  const urlParam = volunteerData.isMember ? "m" : "g";
  const sheetUrl = sheetInputs["SCRIPT_URL"];
  const personalizedUrl = `${sheetUrl}?${urlParam}=${volunteerData.token}`;
  const recipientEmail = volunteerData.email;

  const nameOnly = false;
  const info = getMemberInfoByToken(sheetInputs, volunteerData.token, SHIFT_FLAGS.EVENT, nameOnly  );

  const events = info.events || [];
  const allAvailableShifts = events.flatMap(event => {
    return (event.eventShifts || []).map(shift => ({
      ...shift,
      eventLocation: event.Location,
      deceasedName: event["Deceased Name"],
      eventStartDate: event["Start Date"],
      eventToken: event.Token || event["Token"] || event["Event Token"] || event.Token,
      eventId: event.Token || event["Token"] || event["Event Token"] || event.Token
    }));
  });

  // Filter to only selected shift IDs
  const matchedShiftDetails = allAvailableShifts.filter(shift => shifts.includes(shift["Shift ID"]));

  // 1. GROUP SHIFTS BY EVENT
  const groupedByEvent = {};
  matchedShiftDetails.forEach(shift => {
    const eventKey = shift.eventId || shift.eventLocation || shift.deceasedName || "Unknown Event";
    if (!groupedByEvent[eventKey]) {
      groupedByEvent[eventKey] = {
        eventLocation: shift.eventLocation,
        deceasedName: shift.deceasedName,
        eventStartDate: shift.eventStartDate,
        shifts: []
      }
    }
    groupedByEvent[eventKey].shifts.push(shift);
  });

  // 2. BUILD BODY (GROUPED)
  const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
  let shiftDetails = '';
  let totalShifts = 0;
  for (const eventKey in groupedByEvent) {
    const event = groupedByEvent[eventKey];
    let eventDateFormatted = "Date Unknown";
    if (event.eventStartDate) {
      const date = new Date(event.eventStartDate);
      if (!isNaN(date)) {
        eventDateFormatted = date.toLocaleDateString('en-US', options);
      }
    }

    const locations = getLocations(sheetInputs);
    const fullAddress = getAddressFromLocationName(locations, event.eventLocation);

    shiftDetails += `\nEvent: ${event.deceasedName || "N/A"}\nLocation: ${event.eventLocation}\nAddress: ${fullAddress}\nDate: ${eventDateFormatted}\nShifts:\n`;
    event.shifts.forEach(shift => {
      shiftDetails += `  - Time: ${shift["Shift Display"]}\n`;
      totalShifts++;
    });
    shiftDetails += '\n';
  }

  // 3. COMPOSE EMAIL BODY
  let body = `
Dear ${volunteerData.name},

This is an automatic confirmation that your request to be ${emailAction} the following shift${totalShifts > 1 ? 's' : ''} has been processed successfully:

${shiftDetails}

If you need to cancel or change your confirmation, go to your portal link: ${personalizedUrl}.

Thank you for providing this mitzvah.
`;

  // 4. CHECK BODY SIZE; TRIM IF NECESSARY
  const MAX_BODY = 20000; // Apps Script body limit is 20,000 chars
  if (body.length > MAX_BODY) {
    // Determine how many events/shifts fit
    let partialBody = `
Dear ${volunteerData.name},

This is an automatic confirmation that your request to be ${emailAction} the following shift(s) has been processed successfully:

`;
    let omittedCount = 0;
    for (const eventKey in groupedByEvent) {
      if (partialBody.length > MAX_BODY - 500) { // Leave space for note and link
        omittedCount += groupedByEvent[eventKey].shifts.length;
        continue;
      }
      const event = groupedByEvent[eventKey];
      let eventDateFormatted = "Date Unknown";
      if (event.eventStartDate) {
        const date = new Date(event.eventStartDate);
        if (!isNaN(date)) {
          eventDateFormatted = date.toLocaleDateString('en-US', options);
        }
      }

      const locations = getLocations(sheetInputs);
      const fullAddress = getAddressFromLocationName(locations, event.eventLocation);
      let eventHeader = `Event: ${event.deceasedName || "N/A"}\nLocation: ${event.eventLocation}\nAddress: ${fullAddress}\nDate: ${eventDateFormatted}\nShifts:\n`;
      let eventShiftsLines = '';
      event.shifts.forEach(shift => {
        if ((partialBody.length + eventHeader.length + eventShiftsLines.length + 100) > MAX_BODY - 500) {
          omittedCount++;
          return;
        }
        eventShiftsLines += `  - Time: ${shift["Shift Display"]}\n`;
      });
      if (eventShiftsLines) {
        partialBody += eventHeader + eventShiftsLines + '\n';
      }
    }
    partialBody += (omittedCount > 0 ? `\n[Note: Some shifts were omitted from this message due to size constraints. All your sign-ups are recorded.]\n` : '');
    partialBody += `\nIf you need to cancel or change your confirmation, go to your portal link: ${personalizedUrl}.\n\nThank you for providing this mitzvah.\n`;
    body = partialBody;
  }

  // 5. SUBJECT
  // For brevity: list first event, and maybe count of shifts
  let subject = "Chevra Kadisha Volunteer: Confirmation of Shift " + actionType;
  if (matchedShiftDetails.length === 1) {
    subject += `: ${matchedShiftDetails[0]["Shift Display"]} at ${matchedShiftDetails[0].eventLocation}`;
  } else if (matchedShiftDetails.length > 1) {
    subject += `: ${matchedShiftDetails.length} shifts`;
  }

  if (!recipientEmail || !String(recipientEmail).includes('@')) {
    Logger.log(`Skipping email: Invalid recipient email address: ${recipientEmail}`);
    return;
  }

 // Only build ICS files when shifts are ADDED
  let mailOptions = { body: body };

  if (actionType === "Addition") {
    const icsAttachments = buildIcsAttachmentsForShifts(sheetInputs, groupedByEvent);
    if (icsAttachments.length > 0) {
      mailOptions.attachments = icsAttachments;
    }
  }

  try {
    MailApp.sendEmail(recipientEmail, subject, body, mailOptions);
    Logger.log(`Email sent successfully for ${actionType} to ${recipientEmail}`);
  } catch (e) {
    Logger.log(`ERROR sending email for ${actionType}: ${e.toString()}`);
  }

}

/**
 * Build ICS blobs for all matched shifts.
 * @param {Object} groupedByEvent The object built in sendShiftEmail.
 * @param {Object} sheetInputs For location lookup.
 * @returns {Blob[]} Array of text/calendar blobs.
 */
function buildIcsAttachmentsForShifts(sheetInputs, groupedByEvent) {
  const icsAttachments = [];
  const locations = getLocations(sheetInputs);

  function buildIcsForShift(summary, description, location, dtStart, dtEnd) {
    function toIcsUtc(dt) {
      const pad = n => (n < 10 ? '0' + n : '' + n);
      const y = dt.getUTCFullYear();
      const m = pad(dt.getUTCMonth() + 1);
      const d = pad(dt.getUTCDate());
      const hh = pad(dt.getUTCHours());
      const mm = pad(dt.getUTCMinutes());
      const ss = pad(dt.getUTCSeconds());
      return y + m + d + 'T' + hh + mm + ss + 'Z';
    }

    const uid = Utilities.getUuid();
    const dtStamp = toIcsUtc(new Date());

    const ics =
      'BEGIN:VCALENDAR\r\n' +
      'PRODID:-//Chevra Kadisha//Volunteer Shifts//EN\r\n' +
      'VERSION:2.0\r\n' +
      'METHOD:PUBLISH\r\n' +
      'BEGIN:VEVENT\r\n' +
      'UID:' + uid + '\r\n' +
      'DTSTAMP:' + dtStamp + '\r\n' +
      'DTSTART:' + toIcsUtc(dtStart) + '\r\n' +
      'DTEND:' + toIcsUtc(dtEnd) + '\r\n' +
      'SUMMARY:' + summary + '\r\n' +
      'DESCRIPTION:' + description + '\r\n' +
      'LOCATION:' + location + '\r\n' +
      'END:VEVENT\r\n' +
      'END:VCALENDAR\r\n';

    return Utilities.newBlob(ics, 'text/calendar', 'shift.ics');
  }

  function parseTimeOnDate(date, timePart) {
    const t = String(timePart || '').trim();
    const d = new Date(date);
    const match = t.match(/(\d{1,2}):?(\d{2})?\s*(AM|PM)/i);
    if (!match) return null;
    let hour = parseInt(match[1], 10);
    const min = match[2] ? parseInt(match[2], 10) : 0;
    const ampm = match[3].toUpperCase();
    if (ampm === 'PM' && hour < 12) hour += 12;
    if (ampm === 'AM' && hour === 12) hour = 0;
    d.setHours(hour, min, 0, 0);
    return d;
  }

  for (const eventKey in groupedByEvent) {
    const event = groupedByEvent[eventKey];
    if (!event.eventStartDate) continue;

    const baseDate = new Date(event.eventStartDate);
    if (isNaN(baseDate)) continue;

    const fullAddress = getAddressFromLocationName(locations, event.eventLocation);

    event.shifts.forEach(shift => {
      const timeStr = String(shift['Shift Display'] || '');
      const parts = timeStr.split('-');
      if (parts.length !== 2) return;

      const startDate = parseTimeOnDate(baseDate, parts[0]);
      const endDate = parseTimeOnDate(baseDate, parts[1]);
      if (!startDate || !endDate) return;

      const summary = `Chevra Kadisha Shift - ${event.deceasedName || 'Volunteer'}`;
      const description =
        `Shift time: ${timeStr}\nLocation: ${event.eventLocation}\nAddress: ${fullAddress}`;

      const icsBlob = buildIcsForShift(summary, description, fullAddress, startDate, endDate);
      icsAttachments.push(icsBlob);
    });
  }

  return icsAttachments;
}
