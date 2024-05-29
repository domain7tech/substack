function onOpen() {
  var ui = SpreadsheetApp.getUi(); // Get the user interface object to add a custom menu
  ui.createMenu('Redact Columns')
      .addItem('Choose Columns to Redact', 'changeFontToRedacted') // Adds an item to the custom menu
      .addItem('Revert Sheet to Calibri 12', 'setFontCalibri') // Adds an item to the custom menu

      .addToUi(); // Adds the custom menu to the UI
}


function changeFontToRedacted() {
  var ui = SpreadsheetApp.getUi();
  try {
    // Access the active spreadsheet
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Access the sheet named 'Redaction'
    var sheet = spreadsheet.getSheetByName('Redaction');

    // Display a dialog box for user input on column letters
    var response = ui.prompt('Change Font for Entire Columns', 'Please enter the column letters separated by commas (e.g., A, B):', ui.ButtonSet.OK_CANCEL);

    // Check if the user clicked "OK" and input is not empty
    if (response.getSelectedButton() == ui.Button.OK && response.getResponseText() !== '') {
      var columnsInput = response.getResponseText().split(',').map(function(column) {
        return column.trim().toUpperCase();
      });

      // Get the last row with content in the sheet
      var lastRow = sheet.getLastRow();

      // Change the font for each specified column
      columnsInput.forEach(function(column) {
        var range = sheet.getRange(column + '2:' + column + lastRow);
        range.setFontFamily('Redacted');
      });

      ui.alert('Font changed to Redacted Script for columns: ' + columnsInput.join(', '));
    } else if (response.getSelectedButton() == ui.Button.CANCEL) {
      ui.alert('Operation cancelled.');
    } else {
      ui.alert('No valid input provided.');
    }
  } catch (e) {
    ui.alert('An error occurred: ' + e.toString());
  }
}

//Set Font to Calibri

function setFontCalibri() {
  // Access the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Access the sheet named "Redaction"
  var sheet = spreadsheet.getSheetByName("Redaction");
  
  // Check if the sheet exists
  if (!sheet) {
    Logger.log("Sheet named 'Redaction' does not exist.");
    return;
  }
  
  // Set the font style and size for the entire sheet
  sheet.getRange('A1:Z1000').setFontFamily('Calibri').setFontSize(12);
}


