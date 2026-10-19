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

  tempSheet.createTextFinder('{{Total Deduction}}').replaceAllWith(
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
