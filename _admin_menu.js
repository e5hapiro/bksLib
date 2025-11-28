/**
* -----------------------------------------------------------------
* _admin_menu.js
* Chevra Kadisha Admin Menu functions
* -----------------------------------------------------------------
* _selection_form.js
Version: 1.0.6 * Last updated: 2025-11-12
 * 
 * CHANGELOG v1.0.6:
 *   - Initial implementation of Admin Menu.
 * -----------------------------------------------------------------
 */

function displayAbout() {
  try {
    // Set your time zone (or use Session.getScriptTimeZone())
    var timeZone = Session.getScriptTimeZone();

    // Set the release date as a Date object and format it
    var releaseDateObj = new Date('2025-11-12T00:00:00');
    var qdsDate = Utilities.formatDate(releaseDateObj, timeZone, "dd MMMM yyyy, HH:mm");

    var qdsVersion = "v.1.0.6";
    // Current date and time, formatted
    var now = new Date();
    var qdsCurrent = Utilities.formatDate(now, timeZone, "dd MMMM yyyy, HH:mm");

    var ui = SpreadsheetApp.getUi();

    var result = ui.alert(
      "About Chevra Kadisha Scheduler",
      " " 
      + " \r\n Version " + qdsVersion 
      + " \r\n Released " + qdsDate 
      + " \r\n For bug/feature requests contact marlalshapiro@gmail.com"
      + " \r\n \r\n Current Date Time - " + qdsCurrent,
      ui.ButtonSet.OK
    );

  } catch (error) {
    Logger.log(error.message);
    return false;
  }
}

 