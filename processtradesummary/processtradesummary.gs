/**
 * Main function to process spreadsheets in a specified folder and generate a summary report.
 */
function processFolder() {

  let testfolderid = '1z38V8POr9lXNAoFBM-7I5_8GG9t5E2wx'; // Test folder identifier.
  let realfolderid = '1k0FhOtK-3_mGoH5CnC9axY2l2FsP7h1q'; // Real folder identifier.
  const opfolderid = '1zFsCLutd5pboDamsClZZBRsvPHB6QDwu'; // Dump folder identifier.
  const infolderid = realfolderid; // Identifier of the folder of trades to process.
  const reportdate = Utilities.formatDate(new Date(), 'GMT', 'yyyyMMdd'); // Datestamp for output file (yyyyMMddHHmmss is too fine).
  const foldername = DriveApp.getFolderById(infolderid).getName().toLowerCase(); // Foldername.
  const opfilename = `summary-of-${foldername}-${reportdate}`.replace(/\s/g, '-'); // Output filename.
  const opmetadata = {
    name: opfilename, // Declare the output filename.
    mimeType: MimeType.GOOGLE_SHEETS, // Declare that we are creating a spreadsheet.
    parents: [opfolderid] // Folder id of the newly created spreadsheet.
  }; // Metadata for creating the output file in the desired folder.

  const opmetafile = Drive.Files.create(opmetadata); // To use 'Drive' enable the 'Drive API' Advanced Drive Service.
  const opdatafile = SpreadsheetApp.openById(opmetafile.id); // Trade Summary Spreadsheet
  setupSummarySheet(opdatafile.getSheets()[0]); // Prepare and format spreadsheet cast.
  
  const datafileid = opdatafile.getId(); // Define output datafile identifier.
  const inputfiles = DriveApp.getFolderById(infolderid).getFilesByType(MimeType.GOOGLE_SHEETS); // Get folder files.
  while (inputfiles.hasNext()) { processInputFile(inputfiles.next().getId(), datafileid); } // Proces folder files.
  createPivotTables(datafileid); // Create pivot sheets and tables.
}

/**
 * Sets up the summary sheet with headers and formatting.
 *
 * @param {SpreadsheetApp.Sheet} summarysheet - The sheet object to be set up as the summary sheet.
 */
function setupSummarySheet(summarysheet) {
  const rowheaders = [
    ['NUM', '=COUNT(C7:C)'], // Number of items.
    ['SUM', '=SUM(C7:C)'], // Sum of values.
    ['AVG', '=ROUND(AVERAGE(C7:C),2)'], // Mean of values, rounded to 2 decimal places.
    ['MIN', '=MIN(C7:C)'], // Minimum value.
    ['MAX', '=MAX(C7:C)'], // Maximum value.
    ['DEV', '=IFERROR(ROUND(STDEV(C7:C),2),0)'] // Standard deviation, rounded to 2 decimal places.
  ]; // Define row headers and formulas for the summary sheet.
  summarysheet.setName('SUMMARY').setHiddenGridlines(true).getDataRange().setFontFamily('Oswald');
  summarysheet.getRange(1, 1, rowheaders.length, 3).setBackground('#000000').setFontColor('#ffffff');
  summarysheet.getRange(1, 2, rowheaders.length, rowheaders[0].length).setValues(rowheaders);
  summarysheet.setFrozenRows(6);
}

/**
 * Processes each sheet in the input spreadsheet and appends relevant data to the output spreadsheet.
 *
 * @param {string} ipssid - The ID of the input spreadsheet.
 * @param {string} opssid - The ID of the output spreadsheet.
 */
function processInputFile(ipssid, opssid) {
  const ipss = SpreadsheetApp.openById(ipssid); // Open the input spreadsheet by its ID.
  const opss = SpreadsheetApp.openById(opssid); // Open the output spreadsheet by its ID
  const ipssname = ipss.getName(); // Retrieve the name of the input spreadsheet.
  const sssearch = 'Total'; // Define the string to search for in each sheet.
  ipss.getSheets().map(sheet => sheet.getName()).sort()
         .forEach( worksheet => {
           const result = stringSearch(ipss, worksheet, sssearch); // Use stringSearch function to find string in the current sheet.
           opss.getSheetByName('SUMMARY').appendRow([ipssname, worksheet, result]); // Append the found data to the output spreadsheet.
           Logger.log('Processed ' + ipssname + ' TRADE' + worksheet + ' [ ' + result + ' USD ].'); // Log the processing of each sheet.
         }); // Process each sheet in the input spreadsheet
  opss.getSheetByName('SUMMARY').getDataRange().setBorder(true, true, true, true, true, true); // Set border for all data in sheet.
  opss.getSheetByName('SUMMARY').getRange('C2:C').setNumberFormat('$#,##0.00'); // Format numbers as currency.
}

/**
 * Searches for a specific string in a given sheet and returns the value in the adjacent cell.
 *
 * @param {SpreadsheetApp.Spreadsheet} ss - The spreadsheet object.
 * @param {string} tab - The name of the sheet within the spreadsheet to search.
 * @param {string} sssearch - The string to search for within the sheet.
 * @return {string|null} - Returns the value of the cell adjacent to the found string, or null if not found.
 */
function stringSearch(ss, tab, sssearch) {
  const tabObject = ss.getSheetByName(tab); // Retrieve the specific sheet from the spreadsheet by name.
  const totalCell = findCellContainingString(tabObject, sssearch); // Use helper function to find the cell containing the search string.
  return totalCell ? tabObject.getRange(totalCell.getRow(), totalCell.getColumn() + 1).getValue() : null;
} // If the cell is found, return the value of the adjacent cell; otherwise, return null.

/**
 * Finds the first cell in a sheet that contains a given string.
 *
 * @param {SpreadsheetApp.Sheet} tabObject - The sheet object to search within.
 * @param {string} sssearch - The string to search for within the sheet.
 * @return {SpreadsheetApp.Range|null} - The first cell that contains the string, or null if not found.
 */
function findCellContainingString(tabObject, sssearch) {
  const data = tabObject.getDataRange().getValues(); // Get all values from the sheet as a 2D array.
  for (let i = 0; i < data.length; i++) {
    for (let j = 0; j < data[i].length; j++) {
      if (data[i][j] === sssearch) {
        return tabObject.getRange(i + 1, j + 1); // Return the Range object for the cell that contains the string.
      }
    }
  } // Iterate through the array to find the cell that contains the search string.
  return null;
} // Return null if the string is not found in any cell.

function createPivotTables(ssid) {
  createPivotTable(ssid, 'DEVPIVOT', 1, 2, SpreadsheetApp.PivotTableSummarizeFunction.STDEV);
  createPivotTable(ssid, 'MINPIVOT', 1, 2, SpreadsheetApp.PivotTableSummarizeFunction.MIN);
  createPivotTable(ssid, 'MAXPIVOT', 1, 2, SpreadsheetApp.PivotTableSummarizeFunction.MAX);
  createPivotTable(ssid, 'AVEPIVOT', 1, 2, SpreadsheetApp.PivotTableSummarizeFunction.AVERAGE);
  createPivotTable(ssid, 'SUMPIVOT', 1, 2, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  createPivotTable(ssid, 'NUMPIVOT', 2, 1, SpreadsheetApp.PivotTableSummarizeFunction.COUNT);
}

/**
 * Creates a pivot table in a new sheet within the specified spreadsheet.
 *
 * @param {string} ssid - The ID of the spreadsheet where the pivot table will be created.
 * @param {string} sheetName - The name of the new sheet that will contain the pivot table.
 * @param {number} rowGroupIndex - The index of the column to group by in rows (1-indexed).
 * @param {number} colGroupIndex - The index of the column to group by in columns (1-indexed).
 * @param {string} pivotFunction - The function to apply to the pivot table values (e.g., 'SUM', 'COUNTA').
 */
function createPivotTable(ssid, sheetName, rowGroupIndex, colGroupIndex, pivotFunction) {
  const ss = SpreadsheetApp.openById(ssid); // Open the spreadshet by ID (with the ssid identifier).
  const sn = ss.getSheetByName('SUMMARY'); // Access the 'SUMMARY' sheet within the spreadsheet.
  const sc = 6; // Exclude the first 5 (header) rows to define the range of the Pivot Table source data.
  const pr = sn.getRange(sc, 1, sn.getLastRow() - sc + 1, sn.getLastColumn()); // Define the range of the Pivot Table source data.
  const ps = ss.insertSheet(sheetName); // Create a new sheet for the pivot table.
  const pt = ps.getRange('A1').createPivotTable(pr); // Create the pivot table in the new sheet.
  pt.addRowGroup(rowGroupIndex); // Configure the pivot table row group.
  pt.addColumnGroup(colGroupIndex); // Configure the pivot table column group.
  pt.addPivotValue(3, pivotFunction); // Configure the pivot table pivot values.
  ps.setFrozenRows(2); // Freeze the first two (header) rows.
  ps.setFrozenColumns(1); // Freeze the header column.
  ps.getDataRange().setFontFamily('Oswald'); // Use the 'Oswald' font in the data range.
  ps.getDataRange().setBorder(true,true,true,true,true,true) // Apply a border to every cell in the data range.
  ps.setHiddenGridlines(true); // Hide gridlines.

  if (sheetName !== 'NUMPIVOT') { 
    ps.getDataRange().setNumberFormat('$#,##0.00'); // Format currencies as currency.
    ps.getRange("1:2").setNumberFormat("0"); // Format header row numbers as numbers.
  } 
}