/**
 * ChangeLog:
 * 10/5/24
 * - Added ability to start adding sheets in the future.
 * - Added ability to run on only a given day of the week.
 * - Restored ability to add more than one sheet if we want to from previous versions.
 * 
 * 8/25/23
 * - FEATURE: Added Menu and interactive version of "Clear Sheets" to allow users to clear
 * - BUGFIX: ClearSheets had a RegExp that only worked up to 2020; fixed for the future :)
 * 
*/
/* @OnlyCurrentDoc */

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();

  // Create Sheets Submenu
  const createSheetsMenu = ui
    .createMenu("Create Sheets")
    .addItem("Create sheet for today", "addSheets")
    .addItem("Create sheets for coming days", "createFutureDaysInteractive");

  // Delete Sheets Submenu
  const deleteSheetsMenu = ui
    .createMenu("Delete Sheets")
    .addItem("Delete Some Sheets", "deleteSomeSheets")
    .addItem(
      "Delete All Sheets (run to clear extra sheets)",
      "clearSheetsInteractive"
    );

  // Automation Submenu
  const automationMenu = ui
    .createMenu("Automation")
    .addItem("Test automation now (run now)", "addSheetsAutomated")
    .addItem(
      "Set up automation (run at start of semester)",
      "createTimerTrigger"
    )
    .addItem("Stop automation (run at end of year/semester)", "deleteTriggers")
    .addSeparator()
    .addItem("Update settings for automation", "setScriptProperties");

  // Top-level Menu: Date Tabs
  ui.createMenu("Date Tabs")
    .addSubMenu(createSheetsMenu)
    .addSubMenu(deleteSheetsMenu)
    .addSubMenu(automationMenu)
    .addToUi();
}

const WEEKDAYS = [
  "sunday",
  "monday",
  "tuesday",
  "wednesday",
  "thursday",
  "friday",
  "saturday",
];
const ALL_DAYS = "all";

function getWeekday(text) {
  const cleanedText = text.trim().toLowerCase();
  if (cleanedText === ALL_DAYS) {
    return ALL_DAYS;
  }
  const matchedDay = WEEKDAYS.find((day) => day.startsWith(cleanedText));
  if (ALL_DAYS.startsWith(cleanedText)) {
    return ALL_DAYS;
  }
  return matchedDay || null;
}

function deleteTriggers() {
  let triggers = ScriptApp.getScriptTriggers();
  let ntriggers = triggers.length;
  for (let t of triggers) {
    ScriptApp.deleteTrigger(t);
  }
  SpreadsheetApp.getUi().alert(
    `Deleted ${ntriggers} timer(s) that were set up under this account.`
  );
}

function createFutureDaysInteractive() {
  let response = SpreadsheetApp.getUi().prompt(
    "How many days going forward would you like to create sheets for? (Enter a number greater than 1)"
  );
  let n = Number(response.getResponseText());
  if (n) {
    addSheets(n);
  } else {
    SpreadsheetApp.getUi().alert(
      "You did not enter a number greater than or equal to 1; doing nothing"
    );
  }
}

function createTimerTrigger() {
  let triggers = ScriptApp.getScriptTriggers();
  if (triggers.length) {
    SpreadsheetApp.getUi().alert(
      "There is already a trigger set up under this account. I assume you do not want to add more. Please stop other timers before adding new ones"
    );
    return;
  }
  while (!weekday) {
    let weekdayResponse = SpreadsheetApp.getUi().prompt(
      `This script is designed to add a new tab to the spreadsheet for each weekday, copying
    from the TEMPLATE in each case.
    
    On which weekday should the script run? 
      Enter the full name of the day (e.g., Monday) or just the beginning (e.g., M, Th). 
      Type "All" to run on all weekdays.`
    );
    var weekday = getWeekday(weekdayResponse.getResponseText());
  }

  let response = SpreadsheetApp.getUi().prompt(
    `At what hour would you like it to run each day?

    e.g. enter 7 to add it at 7 am each day.  Do not type am or pm, just the number!
    `
  );
  let n = Number(response.getResponseText());
  if (isNaN(n)) {
    SpreadsheetApp.getUi().alert(
      `I could not interpret the hour as a number: ${response.getResponseText()} => ${n}. Canceling...`
    );
    return;
  } else {
    setScriptProperties();
    let confirmation = SpreadsheetApp.getUi().prompt(
      `Are you sure you want to create this timer? 
      
      Type "Proceed" below to continue!`
    );
    if (confirmation.getResponseText().toUpperCase() == "PROCEED") {
      createTimer(n);
      if (n < 12) {
        ampm = "am";
      } else {
        ampm = "pm";
      }
      let dayMessage;
      if (weekday == ALL_DAYS) {
        dayMessage = "day";
      } else {
        dayMessage = weekday[0].toUpperCase() + weekday.substring(1);
      }
      SpreadsheetApp.getUi().alert(
        `Created new timer to run every ${dayMessage} at ${n}:00 ${ampm}`
      );
    }
  }
}

function createTimer(hour = 5, weekday = "monday") {
  let trigger = ScriptApp.newTrigger("addSheetsAutomated")
    .timeBased()
    .atHour(hour);
  if (weekday !== ALL_DAYS) {
    // Map lowercase day names to ScriptApp.WeekDay enums
    const weekDaysMap = {
      monday: ScriptApp.WeekDay.MONDAY,
      tuesday: ScriptApp.WeekDay.TUESDAY,
      wednesday: ScriptApp.WeekDay.WEDNESDAY,
      thursday: ScriptApp.WeekDay.THURSDAY,
      friday: ScriptApp.WeekDay.FRIDAY,
      saturday: ScriptApp.WeekDay.SATURDAY,
      sunday: ScriptApp.WeekDay.SUNDAY,
    };
    const weekDayEnum = weekDaysMap[weekday];
    if (!weekDayEnum) {
      throw new Error(
        `Invalid day name ${weekday}. Please provide a valid day of the week.`
      );
    }
    trigger.onWeekDay(weekDayEnum);
  }
  let id = trigger.create().getUniqueId();
  console.log("Created trigger: ", id);
}

function clearSheetsInteractive() {
  let ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    "Delete everything but template?",
    'This will DELETE ALL SHEETS: you should only run this at the start of a new year/semester. Are you sure??? Type "Clear" to confirm',
    ui.ButtonSet.YES_NO
  );
  if (response.getSelectedButton() == ui.Button.YES) {
    if (response.getResponseText().toUpperCase() == "CLEAR") {
      clearSheets();
    } else {
      ui.alert("Canceled");
    }
  } else if (response.getSelectedButton() == ui.Button.NO) {
    ui.alert("Canceled");
  }
}

function deleteSomeSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  let text = ui
    .prompt(
      "Delete sheets that match (type text of sheets you want to delete -- regexp syntax):"
    )
    .getResponseText();
  let regex = new RegExp(text);
  console.log("Delete: ", text);
  var sheets = ss.getSheets();
  let matching = [];
  for (let s of sheets) {
    let name = s.getName();
    if (name != "TEMPLATE" && name.search(regex) > -1) {
      matching.push(name);
    }
  }
  var response = ui.alert(
    `Are you sure you want to continue? This will delete ${
      matching.length
    } sheets: ${matching.join(", ")}`,
    ui.ButtonSet.YES_NO
  );
  if (response == ui.Button.YES) {
    for (let toDelete of matching) {
      console.log("Deleting", toDelete);
      let sheet = ss.getSheetByName(toDelete);
      ss.deleteSheet(sheet);
    }
  } else {
    ui.alert("Cancelled");
  }
}

function clearSheets() {
  var sheetMatcher = /[0-9]+\/[0-9]+\/[12345][0-9]/;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  console.log("Got " + sheets.length + " sheets");
  for (var i = 0; i < sheets.length; i++) {
    console.log("Testing: " + i + " - " + sheets[i].getName());
    if (sheetMatcher.test(sheets[i].getName())) {
      console.log("Deleting: " + sheets[i].getName());
      ss.deleteSheet(sheets[i]);
    } // end if
  } // end for
} // end clearSheets

function setScriptProperties() {
  const ui = SpreadsheetApp.getUi();
  const extraDaysToAdd = ui
    .prompt(
      "Enter the number of days to add each time the script runs (i.e. enter 1 to just do a single day, or 5 to do a full work week):"
    )
    .getResponseText();
  const startInFuture = ui
    .prompt(
      "Enter the number of days in the future to start adding days (i.e. enter 0 to create the sheet for the day the script runs, or 7 to create the sheet for one week out):"
    )
    .getResponseText();

  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty(
    "extraDaysToAdd",
    parseInt(extraDaysToAdd, 10) - 1
  );
  scriptProperties.setProperty("futureOffset", startInFuture);

  ui.alert(
    `Settings updated: Adding ${extraDaysToAdd} sheets starting ${startInFuture} days in the future.`
  );
}

function addSheetsAutomated() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const extraDaysProp = scriptProperties.getProperty("extraDaysToAdd");
  const futureOffsetProp = scriptProperties.getProperty("futureOffset");
  let extraDaysToAdd = Number(extraDaysProp);
  let futureOffset = Number(futureOffsetProp);  
  addSheets(extraDaysToAdd, futureOffset);
}

/* Set add extra to a number greater than 0 to add extra days */
function addSheets(addExtra = 0, futureOffset = 0) {
  // First grab the data we'll be copying
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var template_sheet = ss.getSheetByName("TEMPLATE");
  //duplicate = template_sheet.copyTo(ss)
  var range_to_copy = template_sheet.getDataRange();
  // Now we can start creating new sheets...
  // We'll do them for one week out...
  let d = new Date();
  if (futureOffset) {
    d.setDate(d.getDate() + futureOffset);
  }
  // Now let's create one week's worth of days going forward...
  addSheetForDate(d);

  for (var i = 1; i <= addExtra; i++) {
    // We don't actually loop any more -- see RETURN below
    // We'll go through the year...

    let date = new Date(d.getFullYear(), d.getMonth(), d.getDate() + i);
    console.log("Add for future date", date);
    addSheetForDate(date);
  }

  function addSheetForDate(date) {
    let day = date.getDay();
    if ([6, 0].indexOf(day) != -1) {
      console.log("Skipping weekend", date);
      return;
    }
    var sheetName = date.toLocaleDateString({
      weekday: "short",
      month: "numeric",
      day: "numeric",
    });
    if (ss.getSheetByName(sheetName)) {
      console.log("Sheet " + sheetName + " already exists");
    } else {
      ss.setActiveSheet(template_sheet);
      var sheet = ss.duplicateActiveSheet(); // makes duplicate the active sheet
      sheet.setName(sheetName);
      ss.moveActiveSheet(1);
    }
  }
}
