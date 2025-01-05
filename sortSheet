function sortSheet() {
  try {
    var sheetName = 'DCP_payments_checker';
    var columnName = "refresher_updated_at"; // The column to be sorted by
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

    if (!sheet) {
      Logger.log(`Sheet "${sheetName}" not found.`);
      return;
    }

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log("No data available to sort.");
      return;
    }

    // Get the header row (assuming the first row is the header)
    var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Find the column index for "refresher_updated_at"
    var columnToSort = headerRow.indexOf(columnName) + 1; // Add 1 because indexOf() is 0-based

    if (columnToSort === 0) {
      Logger.log(`Column "${columnName}" not found.`);
      return;
    }

    // Get all the data in the sheet (excluding the header)
    var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

    // Filter out rows where the column contains #N/A (skip these rows during sorting)
    var filteredData = data.filter(function(row) {
      return row[columnToSort - 1] !== '#N/A' && row[columnToSort - 1] !== null && row[columnToSort - 1] !== '';
    });

    // Sort the filtered data based on the "refresher_updated_at" column (descending order)
    filteredData.sort(function(a, b) {
      // Check if values are valid (skip #N/A or null)
      if (a[columnToSort - 1] === null) return 1;
      if (b[columnToSort - 1] === null) return -1;
      return b[columnToSort - 1] - a[columnToSort - 1]; // Descending order
    });

    // Reinsert the sorted data back into the sheet
    var sortedData = [];
    var filteredIndex = 0;

    // Rebuild the original data array, placing the sorted values and leaving #N/A values in their positions
    for (var i = 0; i < data.length; i++) {
      if (data[i][columnToSort - 1] === '#N/A' || data[i][columnToSort - 1] === null || data[i][columnToSort - 1] === '') {
        sortedData.push(data[i]); // Keep #N/A rows in their place
      } else {
        sortedData.push(filteredData[filteredIndex]); // Insert sorted values
        filteredIndex++;
      }
    }

    // Write the sorted data back to the sheet
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).setValues(sortedData);

    Logger.log(`Sorted sheet "${sheetName}" by column "${columnName}" in descending order, ignoring #N/A values.`);
  } catch (error) {
    Logger.log("Error in sortSheet: " + error.message);
  }
}
