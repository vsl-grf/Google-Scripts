function dragDownFormulas() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Raw Data");
  if (!sheet) {
    throw new Error("Sheet 'Raw Data' not found");
  }

  const startColumn = 16; // Column P
  const checkColumn = 15; // Column O
  const lastColumn = sheet.getLastColumn();

  // Find the last populated row in column O
  const columnOValues = sheet.getRange(1, checkColumn, sheet.getLastRow(), 1).getValues();
  let lastPopulatedRow = 0;

  for (let i = columnOValues.length - 1; i >= 0; i--) {
    if (columnOValues[i][0] !== "") {
      lastPopulatedRow = i + 1; // Convert to 1-based index
      break;
    }
  }

  if (!lastPopulatedRow) {
    Logger.log("Column O has no data. No formulas will be copied.");
    return;
  }

  // Ensure column O of the last populated row is not empty
  const colOValue = sheet.getRange(lastPopulatedRow, checkColumn).getValue();
  if (!colOValue) {
    Logger.log(`Column O in row ${lastPopulatedRow} is empty. No formulas will be copied.`);
    return;
  }

  // Ensure there's a row before the last populated row
  if (lastPopulatedRow < 2) {
    Logger.log("Not enough rows to drag formulas from the row before.");
    return;
  }

  // Get formulas from the row before the last populated row
  const formulaRange = sheet.getRange(lastPopulatedRow - 1, startColumn, 1, lastColumn - startColumn + 1);
  const formulas = formulaRange.getFormulas();

  if (!formulas.flat().some(Boolean)) { // Check if any formulas exist
    Logger.log(`No formulas found in row ${lastPopulatedRow - 1} starting from column P.`);
    return;
  }

  Logger.log(`Original formulas from row ${lastPopulatedRow - 1}: ${JSON.stringify(formulas)}`);

  // Adjust formulas for the new row
  const adjustedFormulas = formulas.map((row) =>
    row.map((formula) => {
      if (formula) {
        // Adjust row references for relative formulas
        return formula.replace(/(\d+)/g, (match, rowNumber) => {
          return parseInt(rowNumber, 10) + 1; // Increment row number by 1
        });
      }
      return ""; // Return empty if no formula
    })
  );

  Logger.log(`Adjusted formulas for row ${lastPopulatedRow}: ${JSON.stringify(adjustedFormulas)}`);

  // Apply the adjusted formulas to the last populated row
  sheet.getRange(lastPopulatedRow, startColumn, 1, lastColumn - startColumn + 1).setFormulas(adjustedFormulas);
  Logger.log(`Formulas copied from row ${lastPopulatedRow - 1} to row ${lastPopulatedRow} starting from column P.`);
}
