/**
 * -----------------------------------------------------------------
 * _trigger_emails.js
 * Chevra Kadisha Shifts Scheduler
 * Trigger Emails
 * -----------------------------------------------------------------
 * _trigger_emails.js
 * Version: 1.1.0
 * Last updated: 2026-03-26
 * 
 * CHANGELOG v1.1.0:
 *   - Replaced hardcoded body/subject with shmira_notification template via getEmails().
 * -----------------------------------------------------------------
 */
  
/**
 * Sends emails to guests and members who have not yet received an email.
 * @param {Object} sheetInputs - Config including EMAILS_SHEET for getEmails().
 * @param {Array<object>} events - The events data.
 * @param {Array<object>} guests - The guests data.
 * @param {Array<object>} members - The members data.
 * @param {Array<object>} locations - Location data for addresses.
 * @param {Array<object>} existingMapRows - The existing map rows.
 */
function mailMappings(sheetInputs, events, guests, members, locations, existingMapRows) {

  function isValidEmail(email) {
    const regex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return regex.test(email);
  }

  if (typeof sheetInputs.DEBUG === 'undefined') {
    console.log ("DEBUG is undefined");
    return;
  }

  // Load shmira_notification template ONCE (external getEmails() call)
  const emailTemplates = getEmails(sheetInputs);
  const shmiraTemplate = emailTemplates.find(t => t.key === 'shmira_notification');
  if (!shmiraTemplate) {
    QA_Logging('Error: shmira_notification template missing from sheet.', true);
    return;
  }

  const ss = getSpreadsheet_(sheetInputs.SPREADSHEET_ID);
  const mapSheet = ss.getSheetByName(sheetInputs.EVENT_MAP);

  existingMapRows.forEach((mappingRow, idx) => {
    const source = mappingRow[0]; // "Guest" or "Member"
    const eventId = mappingRow[1];
    const personId = mappingRow[2];
    const emailSent = mappingRow[3];

    if (!emailSent) {
      const event = events.find(e => String(e.token).trim() === String(eventId).trim());
      const isGuest = source === "Guest";
      const person = isGuest 
        ? guests.find(g => String(g.token).trim() === String(personId).trim())
        : members.find(m => String(m.token).trim() === String(personId).trim());

      if (!event || !person) {
        QA_Logging(`Skipping mapping (eventId: ${eventId}, personId: ${personId}) - not found.`, true);
        return;
      }

      const validEmail = isValidEmail(person.email);
      if (!validEmail){
        QA_Logging(`Skipping email as email is invalid`); 
        return;
      }

      // Prepare values for template replacement
      const fullAddress = getAddressFromLocationName(locations, event.locationName);
      var urlParam = isGuest ? "g" : "m";
      var webAppUrl = sheetInputs.SCRIPT_URL;
      var personalizedUrl = `${webAppUrl}?${urlParam}=${person.token}`;
      var startDateTimeStr = formattedDateAndTime(event.startDate, event.startTime);
      var endDateTimeStr = formattedDateAndTime(event.endDate, event.endTime); 

      // Replace placeholders in template
      const replaceText = (text) => {
        if (!text) return '';
        return text
          .replace('[firstName]', person.firstName || '')
          .replace('[lastName]', person.lastName || '')
          .replace('[deceasedName]', event.deceasedName || '')
          .replace('[pronoun]', event.pronoun || '')
          .replace('[metOrMeta]', event.metOrMeta || '')
          .replace('[locationName]', event.locationName || '')
          .replace('[fullAddress]', fullAddress)
          .replace('[startDateTimeStr]', startDateTimeStr)
          .replace('[endDateTimeStr]', endDateTimeStr)
          .replace('[personalInfo]', event.personalInfo || '')
          .replace('[personalizedUrl]', personalizedUrl);
      };

      const subject = replaceText(shmiraTemplate.subject);
      
      // Build body: dynamically collect line1 to line30 + optional signature
      const bodyLines = [];
      for (let i = 1; i <= 30; i++) {
        const lineKey = `line${i}`;
        const lineText = replaceText(shmiraTemplate[lineKey]);
        if (lineText.trim()) {
          bodyLines.push(lineText);
        }
      }
      // Add signature if present
      if (shmiraTemplate.signature && replaceText(shmiraTemplate.signature).trim()) {
        bodyLines.push(replaceText(shmiraTemplate.signature));
      }

      const body = bodyLines.join('\n\n');

      try {
        MailApp.sendEmail({
          to: person.email,
          subject: subject,
          body: body,
        });

        var today = new Date;
        mapSheet.getRange(idx + 2, 4).setValue(true); // Email Sent column
        mapSheet.getRange(idx + 2, 5).setValue(today); // Date Sent column
        QA_Logging(`Sent shmira_notification to ${person.email} for event ${event.deceasedName}`);

      } catch (err) {
        QA_Logging(`Error emailing ${person.email}: ${err}`);
      }
    }
  });
}
