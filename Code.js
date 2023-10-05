/**
 * ChangeLog:
 * 8/25/22 
 * - FEATURE: Added Menu and interactive version of "Clear Sheets" to allow users to clear
 * - BUGFIX: ClearSheets had a RegExp that only worked up to 2020; fixed for the future :)
 */

function onOpen (e) {
  SpreadsheetApp.getUi().createMenu(
    'Homework Board Magic'
  )
  .addItem(
    'Create sheet for today','addSheets'
  )
  .addItem(
    'Create sheets for coming days','createFutureDaysInteractive',
  )
  .addItem(
    'Delete Some Sheets','deleteSomeSheets'
  )
  .addItem(
    'Delete All Sheets (run to clear extra sheets)','clearSheetsInteractive'
  )
  .addItem(
    'Set up timer to add sheets every day (run at start of semester)','createTimerTrigger'
  )
  .addItem(
    'Stop adding sheets (run at end of year/semester)','deleteTriggers',
  )
  .addToUi();
}



function deleteTriggers () {
  let triggers = ScriptApp.getScriptTriggers();
  let ntriggers = triggers.length;
  for (let t of triggers) {
    ScriptApp.deleteTrigger(t);
  }
  SpreadsheetApp.getUi().alert(`Deleted ${ntriggers} timer(s) that were set up under this account.`);
}

function createFutureDaysInteractive () {
  let response = SpreadsheetApp.getUi().prompt(
    'How many days going forward would you like to create sheets for? (Enter a number greater than 1)'
  );
  let n = Number(response.getResponseText());
  if (n) {
    addSheets(n);
  } else {
    SpreadsheetApp.getUi().alert('You did not enter a number greater than or equal to 1; doing nothing')
  }
}

function createTimerTrigger () {
  let triggers = ScriptApp.getScriptTriggers();
  if (triggers.length) {
    SpreadsheetApp.getUi().alert('There is already a trigger set up under this account. I assume you do not want to add more. Please stop other timers before adding new ones');
    return;
  }
  let response = SpreadsheetApp.getUi().prompt(
    `This script is designed to add a new tab to the spreadsheet for each weekday, copying
    from the TEMPLATE in each case.

    At what hour would you like it to add the next day's tab?

    e.g. enter 7 to add it at 7 am each day.  Do not type am or pm, just the number!
    `
  );
  let n = Number(response.getResponseText());
  if (isNaN(n)) {
    SpreadsheetApp.getUi().alert(`I could not interpret the hour as a number: ${response.getResponseText()} => ${n}. Canceling...`);
    return;
  } else {
    let confirmation = SpreadsheetApp.getUi().prompt(
      `Are you sure you want to create this timer? Please note if more than one user set this up we will
      create more than one new tab each day, so make sure you are the right person to be turning this on
      and off!
      
      Type "Proceed" below to continue!`
    );
    if (confirmation.getResponseText().toUpperCase()=='PROCEED') {
      createTimer(n);
      if (n < 12) {
        ampm = 'am'
      } else {
        ampm = 'pm';
      }
      SpreadsheetApp.getUi().alert(`Created new timer to run every week day at ${n}:00 ${ampm}`);
    }
  }    
}

function createTimer (n=5) {
  let trigger = ScriptApp.newTrigger('addSheets').timeBased().everyDays(1).atHour(n).create();
  let id = trigger.getUniqueId()
  console.log('Created trigger',id);
}

function clearSheetsInteractive () {
  let ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Delete everything but template?', 
    'This will DELETE ALL SHEETS: you should only run this at the start of a new year/semester. Are you sure??? Type "Clear" to confirm', 
    ui.ButtonSet.YES_NO
  );
  if (response.getSelectedButton() == ui.Button.YES) {
    if (response.getResponseText().toUpperCase()=='CLEAR') {
      clearSheets();
    } else {
      ui.alert('Canceled');
    }
  } else if (response.getSelectedButton() == ui.Button.NO) {
    ui.alert('Canceled');
  } 
}

function deleteSomeSheets () {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  let text = ui.prompt('Delete sheets that match (type text of sheets you want to delete -- regexp syntax):').getResponseText();
  let regex = new RegExp(text);
  console.log('Delete: ',text)
  var sheets = ss.getSheets();
  let matching = [];
  for (let s of sheets) {
    let name = s.getName();
    if (name != 'TEMPLATE' && name.search(regex) > -1) {
      matching.push(name);
    }
  }
  var response = ui.alert(
    `Are you sure you want to continue? This will delete ${matching.length} sheets: ${matching.join(', ')}`,
    ui.ButtonSet.YES_NO
  );
  if (response == ui.Button.YES) {
    for (let toDelete of matching) {
      console.log('Deleting',toDelete)
      let sheet = ss.getSheetByName(toDelete);
      ss.deleteSheet(sheet);
    }
  } else {
    ui.alert("Cancelled");
  }

}

function clearSheets () {
  var sheetMatcher = /[0-9]+\/[0-9]+\/[12345][0-9]/
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets()
  Logger.log('Got '+sheets.length+' sheets');  
  for (var i=0; i<sheets.length; i++) {
    Logger.log('Testing: '+i+' - '+sheets[i].getName());
    if (sheetMatcher.test(sheets[i].getName())) {
      Logger.log('Deleting: '+sheets[i].getName())
      ss.deleteSheet(sheets[i])      
    } // end if
  } // end for
} // end clearSheets

/* Set add extra to a number greater than 0 to add extra days */
function addSheets (addExtra=0) {
  // First grab the data we'll be copying
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var template_sheet = ss.getSheetByName('TEMPLATE');
  //duplicate = template_sheet.copyTo(ss)
  var range_to_copy = template_sheet.getDataRange()  
  // Now we can start creating new sheets...  
  // We'll do them for one week out...
  d = new Date(); 
  // Now let's create one week's worth of days going forward...  
  addSheetForDate(d);

  for (var i=1; i<addExtra; i++) { // We don't actually loop any more -- see RETURN below    
    // We'll go through the year...
    
    let date = new Date(d.getFullYear(),d.getMonth(),d.getDate()+i);
    console.log('Add for future date',date);
    addSheetForDate(date);
  }

  function addSheetForDate (date) {
    let day = date.getDay();
    if ([6,0].indexOf(day) != -1) {
      console.log('Skipping weekend',date);
      return;
    } 
    var sheetName = date.toLocaleDateString({weekday:'short',month:'numeric','day':'numeric'})
    if (ss.getSheetByName(sheetName)) {
        console.log('Sheet '+sheetName+' already exists');
    }  else {
      ss.setActiveSheet(template_sheet)
      var sheet = ss.duplicateActiveSheet() // makes duplicate the active sheet
      sheet.setName(sheetName)
      ss.moveActiveSheet(1)
    }
  }
}
   
