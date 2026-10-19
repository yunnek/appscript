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

    // 8. NEW LOGIC: Deduct CA from BAL in the newly duplicated rows
    if (stopIndex > 1) {
      // Get the range for Columns AG (CA) and AH (BAL) for the new rows
      // (Excluding the dash row itself, so we use stopIndex - 1)
      const targetRange = sheet.getRange(4, 33, stopIndex - 1, 2);
      const financialValues = targetRange.getValues();

      // Loop through each row in the newly duplicated block
      for (let r = 0; r < financialValues.length; r++) {
        const caValue = financialValues[r][0]; // Column AG (CA)
        const balValue = financialValues[r][1]; // Column AH (BAL)

        // Only process the row that actually contains numeric data in both cells
        if (
          typeof caValue === 'number' &&
          caValue > 0 &&
          typeof balValue === 'number'
        ) {
          const newBal = balValue - caValue;

          // Update the BAL cell in that specific row
          sheet.getRange(4 + r, 34).setValue(newBal);
        }
      }
    }
  }
}
