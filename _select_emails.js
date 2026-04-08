/**
 * -----------------------------------------------------------------
 * _select_emails.js
 * Chevra Kadisha Admin Menu functions
 * -----------------------------------------------------------------
 * _selection_emails.js
 * Version: 1.1.0 * Last updated: 2026-03-26
 * 
 * CHANGELOG v1.1.0:
 *   - Updated for transposed sheet format: Col A=keys (Email_ID,Subject,...), Col B+=templates.
 * -----------------------------------------------------------------
 */

/**
 * Returns all email template records from the emails sheet.
 * Expected headers row 1: Key, [template names like shmira_notification]
 *
 * @param {Object} sheetInputs - Configuration object containing SPREADSHEET_ID and EMAILS_SHEET.
 * @returns {Array<Object>} Array of email objects.
 */
function getEmails(sheetInputs) {
  Logger.log("Getting Emails from sheet: " + sheetInputs.EMAILS_SHEET);

  try {
    const ss = getSpreadsheet_(sheetInputs.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetInputs.EMAILS_SHEET);
    if (!sheet) throw new Error("Sheet not found: " + sheetInputs.EMAILS_SHEET);

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      Logger.log("Emails sheet has no data rows.");
      return [];
    }

    const headers = data[0]; // ['Key', 'shmira_notification', ...]
    const templateNames = headers.slice(1); // Template keys (col B+)

    var emails = [];

    // For each template column (starting col 1, B+)
    for (let col = 1; col < headers.length; col++) {
      const templateKey = headers[col].toString().trim();
      if (!templateKey) continue;

      const email = { key: templateKey };

      // For each key row (row 1+, Email_ID, Subject, etc.)
      for (let row = 1; row < data.length; row++) {
        const keyName = data[row][0]?.toString().trim(); // Col A
        if (!keyName) continue;

        const value = data[row][col] || ''; // Current template col
        const propName = keyName.toLowerCase().replace(/ /g, ''); // 'line 1' -> 'line1'

        email[propName] = value;
      }

      email.rowIndex = col + 1; // 1-based column index for template

/*
      // Debug logging (fixed variable names)
      Logger.log('key: %s', email.key);
      Logger.log('subject: %s', email.subject);
      Logger.log('line1: %s', email.line1);
      Logger.log('line2: %s', email.line2);
      Logger.log('line3: %s', email.line3);
      Logger.log('line4: %s', email.line4);
      Logger.log('line5: %s', email.line5);
      Logger.log('line6: %s', email.line6);
      Logger.log('line7: %s', email.line7);
      Logger.log('line8: %s', email.line8);
      Logger.log('line9: %s', email.line9);
      Logger.log('line10: %s', email.line10);
      Logger.log('rowIndex: %s', email.rowIndex);
*/

      emails.push(email);
    }

    Logger.log("Found " + emails.length + " emails.");
    return emails;

  } catch (e) {
    Logger.log("Error in getEmails: " + e.toString());
    throw e;
  }
}
