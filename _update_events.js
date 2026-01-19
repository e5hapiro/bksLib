/**
 * -----------------------------------------------------------------
 * _update_events.js
 * Chevra Kadisha Shifts Scheduler
 * Event Updates
 * -----------------------------------------------------------------
 * _update_events.js
 * Version: 1.0.8 
 * Last updated: 2026-01-19
 * 
 * CHANGELOG v1.0.1:
 *   - Initial implementation of updateEvents_.
 *   - Added logging and error handling.
 *   - Added event, guest, and member data retrieval.
 *   - Added mapping synchronization.
 *   v1.0.6:
 *   - Fixed bug in usage of DEBUG
 * Event Updates
 *   v1.0.7:
 *   - Added filter to getEvents to optimize time
 *   v1.0.8:
 *   - QA Logger bug, was not including DEBUG so defaulting to false
 *   - Included try / catch for error logging when debug is true
 * -----------------------------------------------------------------
 */

/**
 * Updates the event map with the latest data from the events, guests, and members sheets.
 * Sends emails to guests and members who have not yet received an email.
 * @private
 */
function updateEventMap(sheetInputs) {

  if (typeof sheetInputs.DEBUG === 'undefined') {
    console.log("DEBUG is undefined");
    return;
  }

  if (sheetInputs.DEBUG) {
    QA_Logging('updateEventMap is called at ' + new Date().toISOString(), sheetInputs.DEBUG);
    QA_Logging('Spreadsheet ID: ' + sheetInputs.SPREADSHEET_ID, sheetInputs.DEBUG);
    QA_Logging('Event Form Responses: ' + sheetInputs.LATEST_EVENTS, sheetInputs.DEBUG);
    QA_Logging('Guests Sheet: ' + sheetInputs.GUESTS_SHEET, sheetInputs.DEBUG);
    QA_Logging('Members Sheet: ' + sheetInputs.MEMBERS_SHEET, sheetInputs.DEBUG);
    QA_Logging('Event Map: ' + sheetInputs.EVENT_MAP, sheetInputs.DEBUG);
    QA_Logging('Archive Event Map: ' + sheetInputs.ARCHIVE_EVENT_MAP, sheetInputs.DEBUG);
  }

  try {
    // The master workbook
    const ss = getSpreadsheet_(sheetInputs.SPREADSHEET_ID);

    // The events sheet
    const eventSheet = ss.getSheetByName(sheetInputs.LATEST_EVENTS);
    if (!eventSheet) throw new Error('Sheet not found: ' + sheetInputs.LATEST_EVENTS);

    // The guests sheet
    const guestSheet = ss.getSheetByName(sheetInputs.GUESTS_SHEET);
    if (!guestSheet) throw new Error('Sheet not found: ' + sheetInputs.GUESTS_SHEET);

    // The members sheet
    const memberSheet = ss.getSheetByName(sheetInputs.MEMBERS_SHEET);
    if (!memberSheet) throw new Error('Sheet not found: ' + sheetInputs.MEMBERS_SHEET);

    // The map sheet
    const mapSheet = ss.getSheetByName(sheetInputs.EVENT_MAP);
    if (!mapSheet) throw new Error('Sheet not found: ' + sheetInputs.EVENT_MAP);

    // The archive map sheet
    const archiveSheet = ss.getSheetByName(sheetInputs.ARCHIVE_EVENT_MAP);
    if (!archiveSheet) throw new Error('Sheet not found: ' + sheetInputs.ARCHIVE_EVENT_MAP);

    var events          = getEvents(eventSheet);
    var guests          = getApprovedGuests(guestSheet);      // may throw if Approvals missing
    var members         = getApprovedMembers(memberSheet);
    var locations       = getLocations(sheetInputs);
    var existingMapRows = getExistingMapRows(mapSheet);

    syncMappings(events, guests, members, existingMapRows, mapSheet, archiveSheet);

    // After removals need to refresh the existing Map Rows before printing
    existingMapRows = getExistingMapRows(mapSheet);

    // Now send a mail for any guests and events that have not already been sent the mail
    mailMappings(sheetInputs, events, guests, members, locations, existingMapRows);

  } catch (err) {
    // Centralized QA logging of any upstream failure
    QA_Logging('ERROR in updateEventMap: ' + err.message, sheetInputs.DEBUG);
    // Optional: more detail if helpful
    QA_Logging('Stack trace: ' + (err.stack || 'no stack'), sheetInputs.DEBUG);

    // Re-throw so the failure is visible to triggers / UI
    throw err;
  }
}

/**
 * Retrieves the events data from the events sheet.
 *  *
 * Expected headers (second row): Timestamp, Email Address, Deceased Name, Location,
 * Start Date, Start Time, End Date, End Time, Personal Information,
 * Pronoun, Met-or-Meita, Token.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The events sheet.
 * @returns {Array<Object>} The events data.
 * @private
 */
function getEvents(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers = data[1]; // header row is second row in existing design[file:1][file:7]

  var idx = {};
  headers.forEach(function(h, i) { idx[h] = i; });

  // Enforce required Token column
  if (idx['Token'] === undefined) {
    throw new Error("Required column 'Token' is missing from the events sheet.");
  }

  return data.slice(2) // skip first two rows
    .filter(function(row) {
      var ts = row[idx['Timestamp']];
      return ts !== '' && ts !== null;
    })
    .map(function(row) {
      return {
        timestamp: row[idx['Timestamp']],
        email: row[idx['Email Address']],
        deceasedName: row[idx['Deceased Name']],
        locationName: row[idx['Location']],
        startDate: row[idx['Start Date']],
        startTime: row[idx['Start Time']],
        endDate: row[idx['End Date']],
        endTime: row[idx['End Time']],
        personalInfo: row[idx['Personal Information']],
        pronoun: row[idx['Pronoun']],
        metOrMeita: row[idx['Met-or-Meita']],
        token: row[idx['Token']],
        id: row[idx['Token']]
      };
    });
}

/**
 * Retrieves the approved guests data from the guests sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The guests sheet.
 * @returns {Array<object>} - The approved guests data.
 */
function getApprovedGuests(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length === 0) {
    throw new Error("Guests sheet has no data (no header row found).");
  }

  var headers = data[0];

  // Build header index map
  var idx = {};
  headers.forEach(function(h, i) { idx[h] = i; });

  // Enforce required Approvals column
  if (idx['Approvals'] === undefined) {
    throw new Error("Required column 'Approvals' is missing from the guests sheet.");
  }

  // Enforce required Token column
  if (idx['Token'] === undefined) {
    throw new Error("Required column 'Token' is missing from the guests sheet.");
  }

  return data.slice(1)
    .filter(function(row) {
      var approval = (row[idx['Approvals']] || "").toString().trim().toLowerCase();
      return approval === "yes" || approval === "true";
    })
    .map(function(row) {
      return {
        timestamp: row[idx['Timestamp']],
        email: row[idx['Email Address']],
        firstName: row[idx['First Name']],
        lastName: row[idx['Last Name']],
        address: row[idx['Address']],
        city: row[idx['City']],
        state: row[idx['State']],
        zip: row[idx['Zip']],
        phone: row[idx['Phone']],
        canText: row[idx['Can we text you at the above phone number?']],
        names: (row[idx['Name of Deceased']] || "")
          .toString()
          .split(',')
          .map(function(n) { return n.trim().toLowerCase(); }),
        relationship: row[idx['Relationship to Deceased']],
        over18: row[idx['Are you over 18 years old?']],
        sitAlone: row[idx["To sit shmira alone with the Boulder Chevra Kadisha, you must be over 18 years old. If you are under 18 years old, you can sit shmira with a Boulder Chevra Kadisha Member or a parent/guardian. If you will sit shmira with a parent or guardian, have them fill out the form. If you would like to be matched up with a Member of the Boulder Chevra Kadisha, you can continue to complete the form."]],
        canSitDuringBusiness: row[idx["Are you able to sit shmira during the mortuary's normal business hours? (Business hours are Monday through Friday 9am - 5pm)"]],
        sitAfterHours: row[idx["To sit shmira alone after business hours of the mortuary, you must be a Boulder Chevra Kadisha Member. If you are not a member and still want to sit shmira and you are only able to sit shmira outside of regular business hours, we can match you with a member. Would you like to discuss sitting shmira with a Boulder Chevra Kadisha Member?"]],
        affiliation: row[idx["What is your affiliation? (The Boulder Chevra Kadisha is a community wide chevra kadisha. We serve all Jews in Boulder County - affiliated of not.)"]],
        synagogue: row[idx['Name, City and State of synagogue.']],
        onMailingList: row[idx['Do you want to be on our mailing list for events and training?']],
        agreement: row[idx['By submitting this application, I certify the information is true and accurate and I agree with the terms and conditions of sitting shmira with the Boulder Chevra Kadisha.']],
        token: row[idx['Token']],
        approvals: row[idx['Approvals']]
      };
    });
}

/**
 * Retrieves the approved members data from the members sheet.[file:1]
 *
 * Filters rows where Approvals column is 'yes' or 'true' (case-insensitive).
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The members sheet.
 * @returns {Array<Object>} The approved members data.
 * @private
 */
function getApprovedMembers(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers = data[0];
  var idx = {};
  headers.forEach(function(h, i) { idx[h] = i; });

  // Enforce required Approvals column
  if (idx['Approvals'] === undefined) {
    throw new Error("Required column 'Approvals' is missing from the members sheet.");
  }

  // Enforce required Token column
  if (idx['Token'] === undefined) {
    throw new Error("Required column 'Token' is missing from the members sheet.");
  }


  return data.slice(1)
    .filter(function(row) {
      var approval = String(row[idx['Approvals']] || '').trim().toLowerCase();
      return approval === 'yes' || approval === 'true';
    })
    .map(function(row) {
      return {
        timestamp: row[idx['Timestamp']],
        email: row[idx['Email Address']],
        firstName: row[idx['First Name']],
        lastName: row[idx['Last Name']],
        phone: row[idx['Phone']],
        token: row[idx['Token']]
      };
    });
}

/**
 * Retrieves the existing map rows from the map sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The map sheet.
 * @returns {Array<object>} - The existing map rows.
 */
function getExistingMapRows(sheet) {
  var data = sheet.getDataRange().getValues();
  return data.slice(1); // skip headers
}

/**
 * Synchronizes the mappings between events, guests, and members.
 * @param {Array<object>} events - The events data.
 * @param {Array<object>} guests - The guests data.
 * @param {Array<object>} members - The members data.
 * @param {Array<object>} existingRows - The existing map rows.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} mapSheet - The map sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} archiveSheet - The archive map sheet.
 */
function syncMappings(events, guests, members, existingRows, mapSheet, archiveSheet) {
  // Build lookup for existing [source|eventToken|guest/memberToken] => emailSent
  var mapObj = {};
  existingRows.forEach(function(row) {
    var key = row[0] + '|' + row[1] + '|' + row[2];
    mapObj[key] = row[3];
  });

  // Combine guests and members, tagging source
  var allPeople = guests.map(function(p) {
    return Object.assign({}, p, {source: 'Guest'});
  }).concat(
    members.map(function(p) {
      return Object.assign({}, p, {source: 'Member'});
    })
  );

  // Build required new map keys
var requiredKeys = new Set();
var requiredRows = [];

// Guests: Only include mapping if names match
events.forEach(function(event) {
  guests.forEach(function(guest) {
    var matchNames = Array.isArray(guest.names) ? guest.names : [];
    if (matchNames.includes(event.deceasedName.toString().trim().toLowerCase())) {
      var key = "Guest|" + event.token + "|" + guest.token;
      requiredKeys.add(key);
      var emailSent = mapObj[key] !== undefined ? mapObj[key] : "";
      requiredRows.push(["Guest", event.token, guest.token, emailSent]);
    }
  });
  // Members: Map every member to every event
  members.forEach(function(member) {
    var key = "Member|" + event.token + "|" + member.token;
    requiredKeys.add(key);
    var emailSent = mapObj[key] !== undefined ? mapObj[key] : "";
    requiredRows.push(["Member", event.token, member.token, emailSent]);
  });
});


  // Identify obsolete keys to remove
  var toRemoveIndices = [];
  var toRemoveRows = [];
  existingRows.forEach(function(row, idx) {
    var key = row[0] + '|' + row[1] + '|' + row[2];
    if (!requiredKeys.has(key)) {
      toRemoveIndices.push(idx + 2); // +2 for header
      toRemoveRows.push(row);
    }
  });

  // Archive obsolete rows
  if (toRemoveRows.length > 0) {
    archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, toRemoveRows.length, toRemoveRows[0].length).setValues(toRemoveRows);
  }

  // Remove obsolete map rows in reverse order
  toRemoveIndices.reverse().forEach(function(r) {
    mapSheet.deleteRow(r);
  });

  // Add missing rows
  var alreadyPresent = new Set(existingRows.map(function(row) { return row[0] + '|' + row[1] + '|' + row[2]; }));
  var toAdd = requiredRows.filter(function(row) {
    var key = row[0] + '|' + row[1] + '|' + row[2];
    return !alreadyPresent.has(key);
  });
  if (toAdd.length) {
    mapSheet.getRange(mapSheet.getLastRow() + 1, 1, toAdd.length, 4).setValues(toAdd);
  }
}
