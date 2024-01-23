/**
 * Filters trading data for a specific day and generates a summary with a pivot table in the active Google Spreadsheet.
 * No parameters are needed as it operates on the active spreadsheet.
 * Expected Data Structure:
 * - Multiple sheets with trading data.
 * - Each sheet should contain dates in the first column and trading data in subsequent columns.
 * - Data is expected to start from the third row.
 * 
 * NOTE:
 *  -  The trading day include options settlement times for assignment and expiration. To rollback to:
 *  const stoppingday = new Date(startingday.getTime() + (6 * 60 + 30) * 60000);
 */
function filterTradingDay() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // Get the active spreadsheet.
  const tradingdate = new Date('2024/01/22'); // Trading date in the timezone of the exchange (for example, New York).
  const startingday = new Date(tradingdate.getTime() + (23 * 60 + 30) * 60000); // Adds 23 hours and 30 minutes to tradingdate.
  const stoppingday = new Date(startingday.getTime() + (23 * 60 + 59) * 60000); // Adds 23 hours and 59 minutes to startingday.
  const coloredtext = ["blue", "brown", "purple", "orange", "green", "red"]; // Define text colors for each trading account.
  let sheetColors = {};

  // Delete existing Trades and Pivot sheets (if any) and create a new Trades
  let tradetimezone = (startingday.getMonth() + 1 ) + '/' + startingday.getDate();
  let existingSheet = spreadsheet.getSheetByName(tradetimezone + ' Trades');
  let existingPivot = spreadsheet.getSheetByName(tradetimezone + ' Trades Pivot Table');
  if (existingSheet) spreadsheet.deleteSheet(existingSheet); // Delete the existing sheet named after the trading date
  if (existingPivot) spreadsheet.deleteSheet(existingPivot); // Delete the existing pivot table sheet named after the trading date
  let combinedSheet = spreadsheet.insertSheet(tradetimezone + ' Trades', 0); // Create a new sheet for the trading date

  // Setup header and formatting
  let header = spreadsheet.getSheets()[1].getRange("A2:G2").getValues()[0]; // Retrieve header from the second sheet
  combinedSheet.appendRow(["Trading Account"].concat(header)); // Append 'Trading Account' to the header and add it to the new sheet
  combinedSheet.getRange(1, 1, 1, header.length + 1).setBackground('#000000').setFontColor('#FFFFFF'); // Set header background and font color

  // Process sheets
  spreadsheet.getSheets().forEach((sheet, index) => {
    sheetColors[sheet.getName()] = coloredtext[index % coloredtext.length]; // Assign a color to each sheet based on the index
    let data = sheet.getDataRange().getValues().slice(2); // Get data excluding headers from each sheet
    let filteredData = data.filter(row => row[0] instanceof Date && row[0] >= startingday && row[0] <= stoppingday); // Filter data within the trading period

    if (filteredData.length) {
      let startRow = combinedSheet.getLastRow() + 1; // Calculate the starting row for the new data
      filteredData = filteredData.map(row => [sheet.getName()].concat(row)); // Add sheet name to each row of filtered data
      combinedSheet.getRange(startRow, 1, filteredData.length, filteredData[0].length).setValues(filteredData); // Insert filtered data into the new sheet
      combinedSheet.getRange(startRow, 1, filteredData.length, filteredData[0].length).setFontColor(sheetColors[sheet.getName()]); // Set text color for each row based on the sheet name
    }
  });

  // Finalize formatting
  let dataRange = combinedSheet.getDataRange();
  dataRange.setBorder(true, true, true, true, true, true).setFontFamily('Oswald'); // Set border and font family for the data range
  combinedSheet.setHiddenGridlines(true).autoResizeColumns(1, combinedSheet.getLastColumn()).setFrozenRows(1); // Hide gridlines, auto-resize columns, and freeze the first row
  combinedSheet.getRange('E:H').setNumberFormat("$#,##0.00"); // Set currency formatting for specific columns
  combinedSheet.getRange('B:B').setNumberFormat("MM/dd/yy, h:mm:ss AM/PM"); // Set date-time formatting for a specific column

  createPivotTable(combinedSheet); // Create Pivot Table.
}

/**
 * Creates a pivot table in a new sheet based on the filtered data provided by filterTradingDay.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet containing filtered trading data to be summarized in a pivot table.
 * Assumes that the first column of the sheet contains 'Trading Account' information for row grouping.
 */
function createPivotTable(sheet) {
  const pivotTableSheet = sheet.getParent().insertSheet(sheet.getName() + ' Pivot Table'); // Create a pivot sheet.
  const sourceDataRange = sheet.getDataRange(); // Define the source data range of the pivot table. 
  const pivotTable = pivotTableSheet.getRange('A1').createPivotTable(sourceDataRange); // Create pivot table in the new sheet.

  // Add row group (Column A: 'Trading Account')
  pivotTable.addRowGroup(1); // Group data by 'Trading Account'

  // Add pivot values (sum of Columns E, F, and G)
  pivotTable.addPivotValue(5, SpreadsheetApp.PivotTableSummarizeFunction.SUM); // Sum of Column E
  pivotTable.addPivotValue(6, SpreadsheetApp.PivotTableSummarizeFunction.SUM); // Sum of Column F
  pivotTable.addPivotValue(7, SpreadsheetApp.PivotTableSummarizeFunction.SUM); // Sum of Column G

  // Format the pivot table sheet
  pivotTableSheet.getDataRange().setBorder(true, true, true, true, true, true).setFontFamily('Oswald'); // Set border and font family for the pivot table
  pivotTableSheet.getRange('A1:D1').setBackground('#000000').setFontColor('#FFFFFF'); // Set header background and font color for the pivot table
  pivotTableSheet.setHiddenGridlines(true).autoResizeColumns(1, pivotTableSheet.getLastColumn()).setFrozenRows(1); // Hide gridlines, auto-resize columns, and freeze the first row of the pivot table
  pivotTableSheet.getRange('B:D').setNumberFormat("$#,##0.00"); // Set currency formatting for sums in the pivot table
}
