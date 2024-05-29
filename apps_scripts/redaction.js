function changeFontToRedacted() {
  var ui = SpreadsheetApp.getUi();
  try {
    // Access the active spreadsheet
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Access the sheet named 'Birthdays'
    var sheet = spreadsheet.getSheetByName('Birthdays');

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




