/**
 * -----------------------------------------------------------------
 * update_archive.js
 * Chevra Kadisha Shifts Scheduler
 * Historical Archive Builder
 * -----------------------------------------------------------------
 * Version 1.0.0
 * Last updated 2026-01-06
 *
 * CHANGELOG
 * v1.0.0 - Initial implementation of updateArchive.
 *        - Archives past and current shifts into a denormalized table.
 *        - Uses historical.index sheet to avoid duplicate archive entries.
 * -----------------------------------------------------------------
 */

/**
 * Main entry point for the time-based trigger that builds the historical archive.
 * 
 * This function:
 * - Loads events, master shifts, volunteer shifts, members, and guests.
 * - Filters to shifts that are not in the future (based on shift end epoch).
 * - Joins volunteer-shift records to their shift, event, and person (member/guest).
 * - Computes a deterministic archive key: eventToken|shiftId|volunteerToken.
 * - Skips rows already present in historical.index.
 * - Appends new rows to historical.archive and their keys to historical.index.
 * 
 * @param {Object} sheetInputs Configuration object containing sheet names and spreadsheet ID.
 * @param {string} sheetInputs.SPREADSHEETID Master spreadsheet ID.
 * @param {string} sheetInputs.LATESTEVENTS Events sheet name.
 * @param {string} sheetInputs.MASTERSHIFTSSHEET Master shifts sheet name.
 * @param {string} sheetInputs.VOLUNTEERSHIFTS Volunteer shifts sheet name.
 * @param {string} sheetInputs.MEMBERSSHEET Members sheet name.
 * @param {string} sheetInputs.GUESTSSHEET Guests sheet name.
 * @param {string} sheetInputs.HISTORICALARCHIVE Historical archive sheet name.
 * @param {string} sheetInputs.HISTORICALINDEX Historical index sheet name.
 * @param {boolean} [sheetInputs.DEBUG] Enable verbose logging when true.
 * @private
 */
function updateArchive(sheetInputs) {
  if (typeof sheetInputs.DEBUG !== 'undefined' && sheetInputs.DEBUG) {
    console.log('updateArchive called at', new Date().toISOString());
  }

  var ss = getSpreadsheet_(sheetInputs.SPREADSHEET_ID);

  // ========== Ensure all required sheets exist ==========
  var eventSheet = ensureSheetExists_(sheetInputs, ss, sheetInputs.EVENT_FORM_RESPONSES);
  if (!eventSheet) throw new Error('Sheet not found: ' + sheetInputs.EVENT_FORM_RESPONSES);

  var masterShiftsSheet = ensureSheetExists_(sheetInputs, ss, sheetInputs.SHIFTS_MASTER_SHEET);
  if (!masterShiftsSheet) throw new Error('Sheet not found: ' + sheetInputs.SHIFTS_MASTER_SHEET);

  var volunteerShiftsSheet = ensureSheetExists_(sheetInputs, ss, sheetInputs.VOLUNTEER_LIST_SHEET);
  if (!volunteerShiftsSheet) throw new Error('Sheet not found: ' + sheetInputs.VOLUNTEER_LIST_SHEET);

  var membersSheet = ensureSheetExists_(sheetInputs, ss, sheetInputs.MEMBERS_SHEET);
  if (!membersSheet) throw new Error('Sheet not found: ' + sheetInputs.MEMBERS_SHEET);

  var guestsSheet = ensureSheetExists_(sheetInputs, ss, sheetInputs.GUESTS_SHEET);
  if (!guestsSheet) throw new Error('Sheet not found: ' + sheetInputs.GUESTS_SHEET);

  var archiveSheet = ensureSheetExists_(sheetInputs, ss, sheetInputs.ARCHIVE_HISTORICAL);
  if (!archiveSheet) throw new Error('Sheet not found: ' + sheetInputs.ARCHIVE_HISTORICAL);

  var indexSheet = ensureSheetExists_(sheetInputs, ss, sheetInputs.ARCHIVE_HISTORICAL_INDEX);
  if (!indexSheet) throw new Error('Sheet not found: ' + sheetInputs.ARCHIVE_HISTORICAL_INDEX);


  // ========== Ensure headers exist in archive and index sheets ==========
  var archiveHeaders = [
    'archiveKey', 'insertedAt',
    'eventToken', 'eventTimestamp', 'eventEmail', 'eventDeceasedName', 'eventLocationName',
    'eventStartDate', 'eventStartTime', 'eventEndDate', 'eventEndTime',
    'eventPersonalInfo', 'eventPronoun', 'eventMetOrMeita',
    'shiftId', 'shiftEventDate', 'shiftTime', 'shiftMaxVolunteers', 'shiftCurrentVolunteers',
    'shiftStartEpoch', 'shiftEndEpoch',
    'volunteerShiftTimestamp', 'volunteerToken', 'volunteerNameRaw',
    'personType', 'personEmail', 'personFirstName', 'personLastName', 'personPhone'
  ];
  ensureSheetHeaders_(sheetInputs, archiveSheet, archiveHeaders);

  var indexHeaders = ['archiveKey'];
  ensureSheetHeaders_(sheetInputs, indexSheet, indexHeaders);

  // Load core data using existing patterns for events/members/guests
  var events = getEvents(eventSheet);                 // token-based records[file:1][file:7]
  var members = getApprovedMembers(membersSheet);     // token-based records[file:1]
  var guests = getApprovedGuests(guestsSheet);        // token-based records[file:1]
  var shifts = getMasterShifts_(masterShiftsSheet);    // includes shiftId, eventToken, endEpoch[file:3]
  var volunteerShifts = getVolunteerShifts_(volunteerShiftsSheet); // includes shiftId, volunteerToken[file:6]

  // Build lookup maps
  var eventsByToken = indexByKey_(events, function(e) { return e.token; });
  var membersByToken = indexByKey_(members, function(m) { return m.token; });
  var guestsByToken = indexByKey_(guests, function(g) { return g.token; });
  var shiftsById = indexByKey_(shifts, function(s) { return s.shiftId; });

  // Load existing archive index into a Set for O(1) duplicate checks
  var existingKeys = loadArchiveIndex_(indexSheet);

  var nowEpoch = Date.now();

  var newArchiveRows = [];
  var newIndexRows = [];

  // Iterate over volunteer-shift records and construct denormalized archive rows
  volunteerShifts.forEach(function(vs) {
    var shift = shiftsById[vs.shiftId];
    if (!shift) {
      return;
    }

    // Archive only past and current shifts (endEpoch <= now) to avoid multiple volunteers per future shift.
    if (shift.endEpoch > nowEpoch) {
      return;
    }

    var event = eventsByToken[shift.eventToken];
    if (!event) {
      return;
    }

    var person = membersByToken[vs.volunteerToken] || guestsByToken[vs.volunteerToken];
    var personType = membersByToken[vs.volunteerToken] ? 'Member' :
                     guestsByToken[vs.volunteerToken] ? 'Guest' : 'Unknown';

    // Build deterministic archive key
    var archiveKey = buildArchiveKey_(shift.eventToken, shift.shiftId, vs.volunteerToken);

    if (existingKeys.has(archiveKey)) {
      return;
    }

    var insertedAt = new Date();

    var row = buildArchiveRow_(
      archiveKey,
      insertedAt,
      event,
      shift,
      vs,
      person,
      personType
    );

    newArchiveRows.push(row);
    newIndexRows.push([archiveKey]);
  });

  if (newArchiveRows.length === 0) {
    if (typeof sheetInputs.DEBUG !== 'undefined' && sheetInputs.DEBUG) {
      console.log('updateArchive: no new rows to archive.');
    }
    return;
  }

  // Append to historical.archive
  var archiveStartRow = archiveSheet.getLastRow() + 1;
  archiveSheet.getRange(
    archiveStartRow,
    1,
    newArchiveRows.length,
    newArchiveRows[0].length
  ).setValues(newArchiveRows);

  // Append keys to historical.index
  var indexStartRow = indexSheet.getLastRow() + 1;
  indexSheet.getRange(
    indexStartRow,
    1,
    newIndexRows.length,
    1
  ).setValues(newIndexRows);

  if (typeof sheetInputs.DEBUG !== 'undefined' && sheetInputs.DEBUG) {
    console.log('updateArchive: added', newArchiveRows.length, 'rows to archive.');
  }
}


/**
 * Retrieves the master shifts data from the master shifts sheet.[file:3]
 *
 * Expected headers: Shift ID, Deceased Name, Location, Event Date, Shift Time,
 * Max Volunteers, Current Volunteers, Start Epoch, End Epoch, Pronoun,
 * Met-or-Meita, Personal Information, Event Token.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The master shifts sheet.
 * @returns {Array<Object>} The master shifts data.
 * @private
 */
function getMasterShifts_(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers = data[0];
  var idx = {};
  headers.forEach(function(h, i) { idx[h] = i; });

  return data.slice(1).map(function(row) {
    return {
      shiftId: row[idx['Shift ID']],
      deceasedName: row[idx['Deceased Name']],
      locationName: row[idx['Location']],
      eventDate: row[idx['Event Date']],
      shiftTime: row[idx['Shift Time']],
      maxVolunteers: row[idx['Max Volunteers']],
      currentVolunteers: row[idx['Current Volunteers']],
      startEpoch: Number(row[idx['Start Epoch']]) || 0,
      endEpoch: Number(row[idx['End Epoch']]) || 0,
      pronoun: row[idx['Pronoun']],
      metOrMeita: row[idx['Met-or-Meita']],
      personalInfo: row[idx['Personal Information']],
      eventToken: row[idx['Event Token']]
    };
  });
}

/**
 * Retrieves the volunteer shifts data from the volunteer shifts sheet.[file:6]
 *
 * Expected headers: Timestamp, Shift ID, Volunteer Token, Volunteer Name.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The volunteer shifts sheet.
 * @returns {Array<Object>} The volunteer shifts data.
 * @private
 */
function getVolunteerShifts_(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers = data[0];
  var idx = {};
  headers.forEach(function(h, i) { idx[h] = i; });

  return data.slice(1)
    .filter(function(row) {
      var ts = row[idx['Timestamp']];
      return ts !== '' && ts !== null;
    })
    .map(function(row) {
      return {
        timestamp: row[idx['Timestamp']],
        shiftId: row[idx['Shift ID']],
        volunteerToken: row[idx['Volunteer Token']],
        volunteerName: row[idx['Volunteer Name']]
      };
    });
}

/**
 * Builds a generic index object keyed by a key function.
 *
 * @param {Array<Object>} arr Array of objects.
 * @param {function(Object):string} keyFn Function that returns the key.
 * @returns {Object<string, Object>} Map from key to object.
 * @private
 */
function indexByKey_(arr, keyFn) {
  var map = {};
  arr.forEach(function(item) {
    var k = keyFn(item);
    if (k !== undefined && k !== null && k !== '') {
      map[String(k)] = item;
    }
  });
  return map;
}

/**
 * Loads existing archive keys from the historical.index sheet.
 *
 * Expects a single-column sheet with header in row 1 and keys from row 2 downward.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} indexSheet The index sheet.
 * @returns {Set<string>} Set of existing archive keys.
 * @private
 */
function loadArchiveIndex_(indexSheet) {
  var lastRow = indexSheet.getLastRow();
  if (lastRow < 2) {
    return new Set();
  }

  var range = indexSheet.getRange(2, 1, lastRow - 1, 1);
  var values = range.getValues();
  var set = new Set();

  values.forEach(function(row) {
    var key = row[0];
    if (key !== '' && key !== null && key !== undefined) {
      set.add(String(key));
    }
  });

  return set;
}

/**
 * Builds the deterministic archive key from event token, shift ID, and volunteer token.
 *
 * @param {string} eventToken Event token.
 * @param {string} shiftId Shift ID.
 * @param {string} volunteerToken Volunteer token (member or guest token).
 * @returns {string} The archive key.
 * @private
 */
function buildArchiveKey_(eventToken, shiftId, volunteerToken) {
  return [
    String(eventToken || ''),
    String(shiftId || ''),
    String(volunteerToken || '')
  ].join('|');
}

/**
 * Builds a single archive row for writing into historical.archive.
 *
 * Columns (example order):
 *  1. archiveKey
 *  2. insertedAt
 *  3. eventToken
 *  4. eventTimestamp
 *  5. eventEmail
 *  6. eventDeceasedName
 *  7. eventLocationName
 *  8. eventStartDate
 *  9. eventStartTime
 * 10. eventEndDate
 * 11. eventEndTime
 * 12. eventPersonalInfo
 * 13. eventPronoun
 * 14. eventMetOrMeita
 * 15. shiftId
 * 16. shiftEventDate
 * 17. shiftTime
 * 18. shiftMaxVolunteers
 * 19. shiftCurrentVolunteers
 * 20. shiftStartEpoch
 * 21. shiftEndEpoch
 * 22. volunteerShiftTimestamp
 * 23. volunteerToken
 * 24. volunteerNameRaw
 * 25. personType (Member/Guest/Unknown)
 * 26. personEmail
 * 27. personFirstName
 * 28. personLastName
 * 29. personPhone
 *
 * @param {string} archiveKey Deterministic archive key.
 * @param {Date} insertedAt Date/time when the row is inserted.
 * @param {Object} event Event object.
 * @param {Object} shift Shift object.
 * @param {Object} volunteerShift Volunteer-shift object.
 * @param {Object|null} person Person object (member or guest) or null.
 * @param {string} personType 'Member', 'Guest', or 'Unknown'.
 * @returns {Array<*>} One row for historical.archive.
 * @private
 */
function buildArchiveRow_(archiveKey, insertedAt, event, shift, volunteerShift, person, personType) {
  return [
    archiveKey,                          // 1
    insertedAt,                          // 2
    event.token,                         // 3
    event.timestamp,                     // 4
    event.email,                         // 5
    event.deceasedName,                  // 6
    event.locationName,                  // 7
    event.startDate,                     // 8
    event.startTime,                     // 9
    event.endDate,                       // 10
    event.endTime,                       // 11
    event.personalInfo,                  // 12
    event.pronoun,                       // 13
    event.metOrMeita,                    // 14
    shift.shiftId,                       // 15
    shift.eventDate,                     // 16
    shift.shiftTime,                     // 17
    shift.maxVolunteers,                 // 18
    shift.currentVolunteers,             // 19
    shift.startEpoch,                    // 20
    shift.endEpoch,                      // 21
    volunteerShift.timestamp,            // 22
    volunteerShift.volunteerToken,       // 23
    volunteerShift.volunteerName,        // 24
    personType,                          // 25
    person ? person.email : '',          // 26
    person ? person.firstName : '',      // 27
    person ? person.lastName : '',       // 28
    person ? person.phone : ''           // 29
  ];
}


