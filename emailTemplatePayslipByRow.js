function emailTemplatePayslipByRow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // 1. Show the pop-up box asking for the row number(s)
  var response = ui.prompt(
    'Send Payslip',
    'Enter row numbers (e.g., 6,7,8 or 6-8):',
    ui.ButtonSet.OK_CANCEL,
  );

  // Exit if the user clicks Cancel or closes the window
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  var rowInput = response.getResponseText().trim();
  if (!rowInput) {
    ui.alert('Invalid Input', 'Input cannot be empty.', ui.ButtonSet.OK);
    return;
  }

  var dataSheet = ss.getActiveSheet();
  var maxRows = dataSheet.getLastRow();
  var rowsToProcess = [];

  // 2. Parse multi-row formats (comma-separated lists and ranges)
  try {
    var segments = rowInput.split(',');
    for (var i = 0; i < segments.length; i++) {
      var segment = segments[i].trim();
      if (segment.indexOf('-') !== -1) {
        // Handle range logic like "6-8"
        var rangeParts = segment.split('-');
        var startRow = parseInt(rangeParts[0].trim(), 10);
        var endRow = parseInt(rangeParts[1].trim(), 10);

        if (isNaN(startRow) || isNaN(endRow) || startRow > endRow) {
          throw new Error('Invalid range format.');
        }
        for (var r = startRow; r <= endRow; r++) {
          if (rowsToProcess.indexOf(r) === -1) rowsToProcess.push(r);
        }
      } else {
        // Handle single row numbers like "6"
        var singleRow = parseInt(segment, 10);
        if (isNaN(singleRow)) throw new Error('Invalid number format.');
        if (rowsToProcess.indexOf(singleRow) === -1)
          rowsToProcess.push(singleRow);
      }
    }
  } catch (e) {
    ui.alert(
      'Invalid Input',
      'Please use valid numbers, commas, or hyphens (e.g., 6,7,8 or 6-8).',
      ui.ButtonSet.OK,
    );
    return;
  }

  // 3. Filter out invalid rows outside sheet boundary limits
  var validRows = [];
  for (var j = 0; j < rowsToProcess.length; j++) {
    var rNum = rowsToProcess[j];
    if (rNum >= 2 && rNum <= maxRows) {
      validRows.push(rNum);
    }
  }

  if (validRows.length === 0) {
    ui.alert(
      'Invalid Input',
      'No valid row numbers found between 2 and ' + maxRows + '.',
      ui.ButtonSet.OK,
    );
    return;
  }

  var templateSheet = ss.getSheetByName('TEMPLATE Payslip');
  if (!templateSheet) {
    ui.alert("Error: 'TEMPLATE Payslip' tab not found!");
    return;
  }

  // 4. Safely check for and clear out any leftover temporary generation tabs beforehand
  var existingTempSheet = ss.getSheetByName('Temp_Payslip_Generation');
  if (existingTempSheet) {
    ss.deleteSheet(existingTempSheet);
    SpreadsheetApp.flush();
  }

  // 5. Process each valid row inside a loop
  var successCount = 0;
  var errors = [];

  for (var k = 0; k < validRows.length; k++) {
    var currentRow = validRows[k];
    try {
      processSinglePayslip(ss, dataSheet, templateSheet, currentRow);
      successCount++;
    } catch (err) {
      errors.push('Row ' + currentRow + ': ' + err.message);
    }
  }

  // 6. Final summary alert reporting back to user
  var summaryMessage = 'Successfully sent ' + successCount + ' payslip(s).';
  if (errors.length > 0) {
    summaryMessage += '\n\nErrors encountered:\n' + errors.join('\n');
  }
  ui.alert('Process Complete', summaryMessage, ui.ButtonSet.OK);
}
