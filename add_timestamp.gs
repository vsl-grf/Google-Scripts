function addTimestampToNewRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Raw Data");
  if (!sheet) {
    throw new Error("Sheet 'Raw Data' not found");
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("No data rows to process");
    return;
  }
  
  const timestampCol = 15; // Column O (1-indexed)
  const dataRange = sheet.getRange(2, timestampCol, lastRow - 1, 1); // Only column O, skipping header
  const timestamps = dataRange.getValues();

  const now = new Date();
  let updatedTimestamps = false;

  // Add timestamps only to new rows, leave existing blank rows in column O unchanged
  for (let i = 0; i < timestamps.length; i++) {
    if (timestamps[i][0] === "") {
      const row = sheet.getRange(i + 2, 1, 1, sheet.getLastColumn()).getValues()[0]; // Get the full row
      if (row.some((value, index) => index !== timestampCol - 1 && value !== "")) { // Check for data in other columns
        timestamps[i][0] = now; // Add timestamp to rows with new data
        updatedTimestamps = true;
      }
    }
  }

  // Write back updated timestamps to the sheet
  if (updatedTimestamps) {
    dataRange.setValues(timestamps);
  } else {
    Logger.log("No new rows required a timestamp");
  }
}
