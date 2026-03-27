/**
 * -----------------------------------------------------------------
 * _send_shift_email.js
 * Chevra Kadisha Shift Email sender
 * -----------------------------------------------------------------
 * _send_shift_email.js
 * Version: 1.0.0
 * Last updated: 2026-03-26
 * 
 * CHANGELOG v1.0.0:
 *   - Initial implementation of Send Shift Email.
 * -----------------------------------------------------------------
 */

/**
 * Sends shift confirmation emails using sheet templates (shift_add or shift_remove).
 * @param {Object} sheetInputs - Config with EMAILS_SHEET.
 * @param {Object} volunteerData - Volunteer details.
 * @param {Array} shifts - Shift IDs.
 * @param {string} actionType - "Addition" or "Removal".
 * @returns {void}
 */

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
  var info = "";

  if (volunteerData.isMember === true) {
    info = getMemberInfoByToken(sheetInputs, volunteerData.token, SHIFT_FLAGS.EVENT, nameOnly, volunteerData.isMember);
  }
  else {
    info = getGuestInfoByToken(sheetInputs, volunteerData.token, SHIFT_FLAGS.EVENT, nameOnly, volunteerData.isMember);
  }

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

  // 1. GROUP SHIFTS BY EVENT (ORIGINAL)
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

  // Load template (Email_ID values from your sheet)
  const emailTemplates = getEmails(sheetInputs);
  const templateKey = actionType === "Addition" ? "shift_notification_addition" : "shift_notification_removal";
  const template = emailTemplates.find(t => t.key === templateKey);
  if (!template) {
    Logger.log(`Template "${templateKey}" not found.`);
    return;
  }

  // 2. BUILD BODY (GROUPED) - ORIGINAL
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

  // Template replacements
  const replaceText = (text) => {
    if (!text) return '';
    return text
      .replace('[name]', volunteerData.name || '')
      .replace('[emailAction]', emailAction)
      .replace(/shift\[s\]/g, totalShifts > 1 ? 'shifts' : 'shift')
      .replace('[shiftDetails]', shiftDetails)
      .replace('[personalizedUrl]', personalizedUrl);
  };

  // Subject from template + original enhancement
  let subject = replaceText(template.subject) || "Chevra Kadisha Volunteer: Confirmation of Shift " + actionType;
  if (matchedShiftDetails.length === 1) {
    subject += `: ${matchedShiftDetails[0]["Shift Display"]} at ${matchedShiftDetails[0].eventLocation}`;
  } else if (matchedShiftDetails.length > 1) {
    subject += `: ${matchedShiftDetails.length} shifts`;
  }

  // Body from template lines 1-30
  const bodyLines = [];
  for (let i = 1; i <= 30; i++) {
    const lineKey = `line${i}`;
    const lineText = replaceText(template[lineKey]);
    if (lineText.trim()) bodyLines.push(lineText);
  }
  let body = bodyLines.join('\n\n');

  // 4. BODY SIZE CHECK (ORIGINAL LOGIC)
  const MAX_BODY = 20000;
  if (body.length > MAX_BODY) {
    let partialBody = `
Dear ${volunteerData.name},

This is an automatic confirmation that your request to be ${emailAction} the following shift(s) has been processed successfully:

`;
    let omittedCount = 0;
    for (const eventKey in groupedByEvent) {
      if (partialBody.length > MAX_BODY - 500) {
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

  let mailOptions = { body: body };

  // Only build ICS files when shifts are ADDED (ORIGINAL)
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
 * Build ICS blobs for all matched shifts. (ORIGINAL FUNCTION INTACT)
 * @param {Object} sheetInputs For location lookup.
 * @param {Object} groupedByEvent The object built in sendShiftEmail.
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
      const description = `Shift time: ${timeStr}\nLocation: ${event.eventLocation}\nAddress: ${fullAddress}`;

      const icsBlob = buildIcsForShift(summary, description, fullAddress, startDate, endDate);
      icsAttachments.push(icsBlob);
    });
  }

  return icsAttachments;
}

