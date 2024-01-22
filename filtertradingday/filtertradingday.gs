function filterTradingDay() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // Get the active spreadsheet.
  const tradingdate = new Date('2024/01/22'); // Trading date in the timezone of the exchange (for example, New York).
  const startingday = new Date(tradingdate.getTime() + (23 * 60 + 30) * 60000); // Adds 23 hours and 30 minutes to tradingdate.
  const stoppingday = new Date(startingday.getTime() + (6 * 60 + 30) * 60000); // Adds 6 hours and 30 minutes to startingday.
  const coloredtext = ["blue", "brown", "purple", "orange", "green", "red"]; // Define text colors for each trading account.
  let sheetColors = {};

  // Delete existing Trades and Pivot sheets (if any) and create a new Trades
  let tradetimezone = (startingday.getMonth() + 1 ) + '/' + startingday.getDate();
  let existingSheet = spreadsheet.getSheetByName(tradetimezone + ' Trades');
  let existingPivot = spreadsheet.getSheetByName(tradetimezone + ' Trades Pivot Table');
  if (existingSheet) spreadsheet.deleteSheet(existingSheet);
  if (existingPivot) spreadsheet.deleteSheet(existingPivot);
  let combinedSheet = spreadsheet.insertSheet(tradetimezone + ' Trades', 0);

  // Setup header and formatting
  let header = spreadsheet.getSheets()[1].getRange("A2:G2").getValues()[0];
  combinedSheet.appendRow(["Trading Account"].concat(header));
  combinedSheet.getRange(1, 1, 1, header.length + 1).setBackground('#000000').setFontColor('#FFFFFF');

  // Process sheets
  spreadsheet.getSheets().forEach((sheet, index) => {
    sheetColors[sheet.getName()] = coloredtext[index % coloredtext.length];
    let data = sheet.getDataRange().getValues().slice(2);
    let filteredData = data.filter(row => row[0] instanceof Date && row[0] >= startingday && row[0] <= stoppingday);

    if (filteredData.length) {
      let startRow = combinedSheet.getLastRow() + 1;
      filteredData = filteredData.map(row => [sheet.getName()].concat(row));
      combinedSheet.getRange(startRow, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
      combinedSheet.getRange(startRow, 1, filteredData.length, filteredData[0].length).setFontColor(sheetColors[sheet.getName()]);
    }
  });

  // Finalize formatting
  let dataRange = combinedSheet.getDataRange();
  dataRange.setBorder(true, true, true, true, true, true).setFontFamily('Oswald');
  combinedSheet.setHiddenGridlines(true).autoResizeColumns(1, combinedSheet.getLastColumn()).setFrozenRows(1);
  combinedSheet.getRange('E:H').setNumberFormat("$#,##0.00");
  combinedSheet.getRange('B:B').setNumberFormat("MM/dd/yy, h:mm:ss AM/PM");

  createPivotTable(combinedSheet); // Create Pivot Table.
}

function createPivotTable(sheet) {
  const pivotTableSheet = sheet.getParent().insertSheet(sheet.getName() + ' Pivot Table'); // Create a pivot sheet.
  const sourceDataRange = sheet.getDataRange(); // Define the source data range of the pivot table. 
  const pivotTable = pivotTableSheet.getRange('A1').createPivotTable(sourceDataRange); // Create pivot table in the new sheet.

  // Add row group (Column A: 'Trading Account')
  pivotTable.addRowGroup(1);

  // Add pivot values (sum of Columns E, F, and G)
  pivotTable.addPivotValue(5, SpreadsheetApp.PivotTableSummarizeFunction.SUM); // Sum of Column E
  pivotTable.addPivotValue(6, SpreadsheetApp.PivotTableSummarizeFunction.SUM); // Sum of Column F
  pivotTable.addPivotValue(7, SpreadsheetApp.PivotTableSummarizeFunction.SUM); // Sum of Column G

  // Format the pivot table sheet
  pivotTableSheet.getDataRange().setBorder(true, true, true, true, true, true).setFontFamily('Oswald');
  pivotTableSheet.getRange('A1:D1').setBackground('#000000').setFontColor('#FFFFFF');
  pivotTableSheet.setHiddenGridlines(true).autoResizeColumns(1, pivotTableSheet.getLastColumn()).setFrozenRows(1);
  pivotTableSheet.getRange('B:D').setNumberFormat("$#,##0.00"); // Assuming Columns B, C, D in pivot table will have the sums
}
