/**
* -----------------------------------------------------------------
* _trigger_emails.js
* Chevra Kadisha Shifts Scheduler
* Trigger Emails
* -----------------------------------------------------------------
* _trigger_emails.js
Version: 1.0.6* Last updated: 2025-11-12
 * 
 * CHANGELOG v1.0.1:
 *   - Initial implementation of mailMappings.
 *   - Added logging and error handling.
 *   - Added event, guest, and member data retrieval.
 *   - Added mapping synchronization.
  *   v1.0.6:
 *   - Fixed bug in usage of DEBUG
* Trigger Emails
 * -----------------------------------------------------------------
 */
  

/**
 * Sends emails to guests and members who have not yet received an email.
 * @param {Array<object>} events - The events data.
 * @param {Array<object>} guests - The guests data.
 * @param {Array<object>} members - The members data.
 * @param {Array<object>} existingMapRows - The existing map rows.
 */
function mailMappings(sheetInputs, events, guests, members, locations, existingMapRows) {

  if (typeof sheetInputs.DEBUG === 'undefined') {
  console.log ("DEBUG is undefined");
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


      /**
       * Looks up the confidential physical address based on the location name (e.g., 'Site A').
       * This function retrieves the secret address stored in ADDRESS_CONFIG.
       * @param {string} locationName The short name (e.g., 'Site A' or 'Site B').
       * @returns {string} The full physical address or a helpful message.
       * @private
       */
      /*
      * Function to find a location by name in an array of location objects 
      * and return the full address as a concatenated string.
      */
      function getAddressFromLocationName_(locations, locationName) {        
        // 1. Find the location object in the array where location.name matches locationName.
        const foundLocation = locations.find(location => location.name === locationName);

        // 2. If a matching location object is found, construct and return the full address string.
        if (foundLocation) {
          // Access properties directly and concatenate: "Street, City, State Zip"
          return `${foundLocation.street}, ${foundLocation.city}, ${foundLocation.state} ${foundLocation.zip}`; 
        }

        // 3. If the location is not configured (e.g. 'Virtual Shift', 'Other'), return the original location name.
        return locationName;
      }      

      // Use person's first and last name for greeting
      const subject = `Baruch Dayan Ha-Emet - Death of ${event.deceasedName} - Chevra Kadisha Services Needed`;
      const fullAddress = getAddressFromLocationName_(locations, event.locationName);

      // Generate URL for email
      var urlParam = isGuest ? "g" : "m";
      var webAppUrl = sheetInputs.SCRIPT_URL;
      var personalizedUrl = `${webAppUrl}?${urlParam}=${person.token}`;
      
      // --- Formatted date strings for email ---
      var startDateTimeStr = formattedDateAndTime(event.startDate);
      var endDateTimeStr = formattedDateAndTime(event.endDate); 

      // Compose body with all fields
      const body = `
Dear ${person.firstName} ${person.lastName},

Baruch Dayan Ha'Emet. We sadly notify you of the death of ${event.deceasedName}.

${event.pronoun} ${event.metOrMeita} is at ${event.locationName} (Address: ${fullAddress}).

Shmira will start on ${startDateTimeStr} and is scheduled to end for the funeral on ${endDateTimeStr}.

${event.personalInfo}

Volunteer Portal Link (unique to you): ${personalizedUrl}

As a reminder, only Boulder Chevra Kadisha Member Volunteers can sit shmira after business hours at the mortuaries.
Log in to the Member Volunteer portal on BoulderChevraKadisha.org for after hours facility access information.

Thank you for your mitzvah of providing shmira for this member of our community.

(If you have questions, reply to this email.)
      `;

      try {
        MailApp.sendEmail({
          to: person.email,
          subject: subject,
          body: body,
        });

        // Mark email as sent in map sheet (row indices are 1-based, +2 for header offset)
        var today = new Date;
        mapSheet.getRange(idx + 2, 4).setValue(true); // Email Sent column
        mapSheet.getRange(idx + 2, 5).setValue(today); // Date Sent column
        QA_Logging(`Sent email to ${person.email} for event ${event.deceasedName}`);

      } catch (err) {
        QA_Logging(`Error emailing ${person.email}: ${err}`);
      }
    }
  });
}



