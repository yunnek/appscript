function duplicateRowsUntilDash() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 4) return; // Exit if there is no data from row 4 onwards

  const columnA = sheet.getRange('A4:A' + lastRow).getValues();
  let stopIndex = 1;

  // Find where the "-" is located
  for (let i = 0; i < columnA.length; i++) {
    if (columnA[i][0] === '-') {
      break;
    }
    stopIndex++;
  }

  if (stopIndex > 0) {
    // 1. Insert blank rows at row 4 to make space
    sheet.insertRowsBefore(4, stopIndex);

    // 2. Define the source range (now shifted down by stopIndex)
    const sourceRange = sheet.getRange(4 + stopIndex, 1, stopIndex, lastColumn);

    // 3. Define the destination range starting at row 4
    const destinationRange = sheet.getRange(4, 1, stopIndex, lastColumn);

    // 4. Copy values and formats programmatically to bypass manual validation errors
    sourceRange.copyTo(
      destinationRange,
      SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
      false,
    );

    // 5. Update Column F in the newly duplicated rows with your specific formula
    // (Excluding the dash row itself, so we subtract 1 from stopIndex)
    if (stopIndex > 1) {
      // Get values of Column F for the newly duplicated rows (excluding the dash row)
      const targetFormulaRange = sheet.getRange(4, 6, stopIndex - 1, 1);
      const fValues = targetFormulaRange.getValues();

      // Loop through each cell value in the range
      for (let r = 3; r < fValues.length; r++) {
        // If the cell is not empty, update it with the formula
        if (fValues[r][0] !== '' && fValues[r][0] !== null) {
          sheet
            .getRange(4 + r, 6)
            .setFormula(
              '=IF(OR($J$4="", $K$4=""), "", NETWORKDAYS.INTL($J$4, $K$4, "0000001"))',
            );
        }
      }
    }

    // 6. NEW LOGIC: Dynamic Date Update for Cell C4 based on its current value
    const cellC4 = sheet.getRange('C4');
    const currentDate = cellC4.getValue();
    if (currentDate instanceof Date) {
      let currentDay = currentDate.getDate();
      let currentMonth = currentDate.getMonth(); // 0 = Jan, 1 = Feb, etc.
      let currentYear = currentDate.getFullYear();

      let nextDate = new Date(currentDate);

      // If day is 20th (e.g., July 20), move to 5th of the NEXT month (e.g., Aug 5)
      if (currentDay === 20) {
        nextDate.setMonth(currentMonth + 1);
        nextDate.setDate(5);
      }
      // If day is 5th (e.g., Aug 5), stay in the SAME month but move to the 20th (e.g., Aug 20)
      else if (currentDay === 5) {
        nextDate.setDate(20);
      }
      // Fallback fallback: If the date structure is different, it jumps ahead 15 days
      else {
        nextDate.setDate(currentDay + 15);
      }

      cellC4.setValue(nextDate);
    }

    // 7. NEW LOGIC: Dynamic Date Update for Cells J4 and K4
    const cellJ4 = sheet.getRange('J4');
    const cellK4 = sheet.getRange('K4');
    const initialK4Value = cellK4.getValue();

    if (initialK4Value instanceof Date) {
      // Set J4 to be +1 day of the initial value of K4
      let newJ4Date = new Date(initialK4Value);
      newJ4Date.setDate(newJ4Date.getDate() + 1);
      cellJ4.setValue(newJ4Date);

      // Determine the new K4 date based on rules
      let newK4Date = new Date(initialK4Value);
      let dayK4 = initialK4Value.getDate();
      let monthK4 = initialK4Value.getMonth();
      let yearK4 = initialK4Value.getFullYear();

      // Check if current K4 is the last day of its month
      // Setting date to 0 of the next month returns the last day of the current month
      let lastDayOfCurrentMonth = new Date(yearK4, monthK4 + 1, 0).getDate();

      if (dayK4 === 15) {
        // Rule: Change to the last day of the current month
        newK4Date.setDate(lastDayOfCurrentMonth);
      } else if (dayK4 === lastDayOfCurrentMonth) {
        // Rule: Change to the 15th day of the next month
        newK4Date.setMonth(monthK4 + 1);
        newK4Date.setDate(15);
      } else {
        // Fallback: Default shift if the date doesn't match standard cycles
        newK4Date.setDate(dayK4 + 15);
      }

      cellK4.setValue(newK4Date);
    }
  }
}

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

// Helper Function: Isolated core processing logic for an individual row
function processSinglePayslip(ss, dataSheet, templateSheet, rowNum) {
  var ui = SpreadsheetApp.getUi();

  // Helper function to pull cells and format currencies nicely
  function getFormattedValue(colLetter) {
    var cellValue = dataSheet.getRange(colLetter + rowNum).getValue();
    if (typeof cellValue === 'number') {
      return cellValue.toLocaleString('en-US', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      });
    }
    return cellValue.toString();
  }

  // Robust Helper function to clean cell values and prevent blanks
  function getCleanedValue(colLetter) {
    var cellValue = dataSheet.getRange(colLetter + rowNum).getValue();

    if (
      cellValue === '' ||
      cellValue === null ||
      cellValue === undefined ||
      cellValue.toString().toLowerCase() === 'n/a'
    ) {
      return '0.00';
    }
    if (typeof cellValue === 'number') {
      return cellValue.toLocaleString('en-US', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      });
    }

    var cleanNum = parseFloat(cellValue.toString().replace(/[^0-9.-]+/g, ''));
    if (!isNaN(cleanNum)) {
      return cleanNum.toLocaleString('en-US', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      });
    }

    return cellValue.toString();
  }

  var timezone = ss.getSpreadsheetTimeZone();

  // Fetch from H4 and I4 directly for the Pay Period
  var rawPeriodFrom = dataSheet.getRange('J4').getValue();
  var rawPeriodTo = dataSheet.getRange('K4').getValue();

  var periodFrom =
    rawPeriodFrom instanceof Date
      ? Utilities.formatDate(rawPeriodFrom, timezone, 'MMMM dd, yyyy')
      : rawPeriodFrom;
  var periodTo =
    rawPeriodTo instanceof Date
      ? Utilities.formatDate(rawPeriodTo, timezone, 'MMMM dd, yyyy')
      : rawPeriodTo;

  var payPeriod = '';
  try {
    var partsFrom = periodFrom.split(' ');
    var partsTo = periodTo.split(' ');

    var monthFrom = partsFrom[0];
    var dayFrom = partsFrom[1].replace(',', '');
    var dayTo = partsTo[1].replace(',', '');
    var yearTo = partsTo[2];

    payPeriod = monthFrom + ' ' + dayFrom + ' - ' + dayTo + ', ' + yearTo;
  } catch (e) {
    payPeriod = periodFrom + ' - ' + periodTo;
  }

  var rawPayDate = dataSheet.getRange('C4').getValue();
  var payDate =
    rawPayDate instanceof Date
      ? Utilities.formatDate(rawPayDate, timezone, 'MMMM dd, yyyy')
      : rawPayDate.toString();

  // Create a temporary copy of the template to merge data
  var tempSheet = templateSheet.copyTo(ss);
  tempSheet.setName('Temp_Payslip_Generation');

  // Loop through data and replace placeholders {{Header Name}} with actual row data
  var employeeName = dataSheet
    .getRange('B' + rowNum)
    .getValue()
    .toString()
    .trim();
  var recipient = dataSheet
    .getRange('D' + rowNum)
    .getValue()
    .toString()
    .trim();

  // Validate email address presence fallback
  if (!recipient || !recipient.includes('@')) {
    recipient = 'kennyagiuy@gmail.com';
  }

  tempSheet.createTextFinder('{{Name}}').replaceAllWith(employeeName);
  tempSheet.createTextFinder('{{Role}}').replaceAllWith(
    dataSheet
      .getRange('C' + rowNum)
      .getValue()
      .toString(),
  );
  tempSheet.createTextFinder('{{Worked Days}}').replaceAllWith(
    dataSheet
      .getRange('F' + rowNum)
      .getValue()
      .toString(),
  );
  tempSheet.createTextFinder('{{Pay Date}}').replaceAllWith(payDate);
  tempSheet.createTextFinder('{{Pay Period}}').replaceAllWith(payPeriod);

  tempSheet.createTextFinder('{{Rate}}').replaceAllWith(getCleanedValue('G'));
  tempSheet.createTextFinder('{{Gross}}').replaceAllWith(getCleanedValue('X'));

  // Deductions mapping
  tempSheet
    .createTextFinder('{{SSS Cont}}')
    .replaceAllWith(getCleanedValue('AA'));
  tempSheet
    .createTextFinder('{{Pagibig}}')
    .replaceAllWith(getCleanedValue('AB'));
  tempSheet
    .createTextFinder('{{Philhealth}}')
    .replaceAllWith(getCleanedValue('AC'));
  tempSheet
    .createTextFinder('{{SSS Loan}}')
    .replaceAllWith(getCleanedValue('AD'));
  tempSheet.createTextFinder('{{CA}}').replaceAllWith(getCleanedValue('AG'));

  // --- ADDED DYNAMIC CODES FOR CA BALANCE PLACEHOLDER ---
  // If your main table doesn't have a CA Balance column, pull it here.
  // Assuming column AH holds CA Balance. Change "AH" to your actual CA Balance column letter if different.
  tempSheet
    .createTextFinder('{{CA BAL}}')
    .replaceAllWith(getCleanedValue('AH'));

  // Calculate Total Deductions dynamically from values in columns Q, R, S, T
  var pVal = parseFloat(dataSheet.getRange('AA' + rowNum).getValue()) || 0; // SSS
  var qVal = parseFloat(dataSheet.getRange('AB' + rowNum).getValue()) || 0; // Pag-ibig
  var rVal = parseFloat(dataSheet.getRange('AC' + rowNum).getValue()) || 0; // Philhealth
  var sVal = parseFloat(dataSheet.getRange('AD' + rowNum).getValue()) || 0; // SSS loan
  var tVal = parseFloat(dataSheet.getRange('AG' + rowNum).getValue()) || 0; // Cash advance
  var totalDeductions = pVal + qVal + rVal + sVal + tVal;

  tempSheet
    .createTextFinder('{{Total Deduction}}')
    .replaceAllWith(
      totalDeductions.toLocaleString('en-US', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      }),
    );

  tempSheet
    .createTextFinder('{{Net Pay}}')
    .replaceAllWith(getCleanedValue('AI'));

  // --- NEW LOGIC: HIDE/DELETE CASH ADVANCE ROWS IF CA IS ZERO ---
  if (tVal === 0) {
    // We look up row numbers by searching for specific string signatures in the temporary sheet
    var textFinderCA = tempSheet.createTextFinder('Cash Advance').findNext();
    var textFinderCABal = tempSheet.createTextFinder('CA Balance').findNext();

    // Delete the "CA Balance" row first if found, so row index order doesn't disrupt structural shifts
    if (textFinderCABal) {
      tempSheet.deleteRow(textFinderCABal.getRow());
    }
    // Delete the "Cash Advance" row second
    if (textFinderCA) {
      tempSheet.deleteRow(textFinderCA.getRow());
    }
  }

  SpreadsheetApp.flush(); // Force changes to render layout updates

  // PDF Generation Process using the Export URL
  var ssId = ss.getId();
  var sheetId = tempSheet.getSheetId();
  var url = ss.getUrl().replace(/edit$/, '') + 'export?';

  var exportOptions = {
    exportFormat: 'pdf',
    format: 'pdf',
    size: 'A6', // Triggers exact 1/4 of an A4 sheet layout
    portrait: 'true', // Strict Portrait mode
    scale: '4', // Fit to Page (scales your data to fit the A6 boundaries)
    top_margin: '0.15', // Small margins maximize space on this narrow canvas
    bottom_margin: '0.15',
    left_margin: '0.3',
    right_margin: '0.15',
    horizontal_alignment: 'CENTER', // Centers your payslip horizontally
    vertical_alignment: 'MIDDLE', // Centers your payslip vertically
    sheetnames: 'false',
    printtitle: 'false',
    gridlines: 'false',
    fzr: 'false',
    gid: sheetId,
  };

  var urlParts = [];
  for (var key in exportOptions) {
    urlParts.push(key + '=' + exportOptions[key]);
  }
  var exportUrl = url + urlParts.join('&');

  // Generate file blob payload
  var responseBlob = UrlFetchApp.fetch(exportUrl, {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
      MuteHttpExceptions: true,
    },
  });

  var pdfBlob = responseBlob
    .getBlob()
    .setName('Payslip_' + employeeName + '_Row' + rowNum + '.pdf');

  // Delete the temporary sheet to keep workbook clean before loop cycles
  ss.deleteSheet(tempSheet);
  SpreadsheetApp.flush();

  // Construct and send the email
  var subject = 'Payslip for ' + employeeName;
  var body =
    'Hello,\n\nPlease find your attached payslip generated from row ' +
    rowNum +
    '.\n\nBest regards,\nSolar Ace Premium Products Incorporated';

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    body: body,
    attachments: [pdfBlob],
  });
}

// Automatically creates a top menu button when you open the Google Sheet
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('HR Tools')
    .addItem('Send Payslip by Row', 'emailTemplatePayslipByRow')
    .addItem('Create next cut-off', 'duplicateRowsUntilDash')
    .addToUi();
}
