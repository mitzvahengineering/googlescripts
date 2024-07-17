/**
 * Main function to process spreadsheets in a specified folder and generate a summary report.
 * It processes all Google Sheets files in the specified folder, extracts relevant data, and creates a summary sheet
 * with statistical information and pivot tables.
 */
function processFolder() {
  const testfolderid = '1z38V8POr9lXNAoFBM-7I5_8GG9t5E2wx'; // Test folder identifier.
  const realfolderid = '1k0FhOtK-3_mGoH5CnC9axY2l2FsP7h1q'; // Real folder identifier.
  const opfolderid = '1zFsCLutd5pboDamsClZZBRsvPHB6QDwu'; // Folder where the summary report will be saved.
  const infolderid = realfolderid; // Identifier of the folder of trades to process.
  const reportdate = Utilities.formatDate(new Date(), 'GMT', 'yyyyMMdd'); // Datestamp for output file (format: yyyyMMdd).
  const foldername = DriveApp.getFolderById(infolderid).getName().toLowerCase(); // Name of the input folder.
  const opfilename = `summary-of-${foldername}-${reportdate}`.replace(/\s/g, '-'); // Output filename.
  const opmetadata = {
    name: opfilename,  // Declare the output filename.
    mimeType: MimeType.GOOGLE_SHEETS, // Declare that we are creating a spreadsheet.
    parents: [opfolderid] // Folder id of the newly created spreadsheet.
  }; // Metadata for creating the output file in the desired folder.

  // Create a new Google Sheets file for the summary report.
  const opmetafile = Drive.Files.create(opmetadata); // Requires 'Drive API' Advanced Drive Service.
  const opdatafile = SpreadsheetApp.openById(opmetafile.id); // Open the newly created spreadsheet.
  const summarized = opdatafile.getSheets()[0]; // Get the first sheet of the newly created spreadsheet.
  setupSummarySheet(summarized); // Prepare and format the summary sheet.

  const inputfiles = DriveApp.getFolderById(infolderid).getFilesByType(MimeType.GOOGLE_SHEETS); // Get all Google Sheets files in the input folder.
  const collection = []; // Array to store the collected data.

  // Loop through each file in the input folder and process it.
  while (inputfiles.hasNext()) {
    const ipssid = inputfiles.next().getId(); // Get the ID of the current input file.
    processInputFile(ipssid, collection); // Collect data from each file.
  }

  // Append collected data to the summary sheet in one go.
  if (collection.length > 0) {
    summarized.getRange(7, 1, collection.length, collection[0].length).setValues(collection); // Write collected data to the sheet.
  }
  summarized.getRange('C2:C').setNumberFormat('$#,##0.00'); // Format numbers as currency.
  summarized.getDataRange().setBorder(true, true, true, true, true, true); // Set border for all data in the sheet.

  createPivotTables(opdatafile.getId()); // Create pivot tables.
  createPivotTableOnSummary(opdatafile.getId(), 1, 3); // Create pivot table on the summary sheet.
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
  ];

  summarysheet.setName('TRADE SUMMARY').setHiddenGridlines(true).getDataRange().setFontFamily('Oswald'); // Set the name of the sheet and hide gridlines.
  summarysheet.getRange(1, 1, rowheaders.length, 3).setBackground('#000000').setFontColor('#ffffff'); // Set the header row with background and font color.
  summarysheet.getRange(1, 2, rowheaders.length, rowheaders[0].length).setValues(rowheaders); // Set the values for the header row.
  summarysheet.setFrozenRows(6); // Freeze the first six rows.
}

/**
 * Processes each sheet in the input spreadsheet and collects relevant data.
 *
 * @param {string} ipssid - The ID of the input spreadsheet.
 * @param {Array} opssid - The array to store the collected data.
 */
function processInputFile(ipssid, opssid) {
  const ipss = SpreadsheetApp.openById(ipssid); // Open the input spreadsheet by its ID.
  const tabs = ipss.getSheets(); // Get all sheets in the input spreadsheet.
  tabs.forEach(sheet => {
    const result = stringSearch(sheet, 'Total'); // Use stringSearch function to find the string in the current sheet.
    if (result !== null) {
      opssid.push([ipss.getName(), sheet.getName(), result]); // Collect the found data.
    }
  }); // Loop through each sheet in the input spreadsheet.
}

/**
 * Searches for a specific string in a given sheet and returns the value in the adjacent cell.
 *
 * @param {SpreadsheetApp.Sheet} ssheet - The sheet object to search within.
 * @param {string} sstring - The string to search for within the sheet.
 * @return {string|null} - Returns the value of the cell adjacent to the found string, or null if not found.
 */
function stringSearch(ssheet, sstring) {
  const values = ssheet.getDataRange().getValues(); // Get all values from the sheet as a 2D array.
  // Loop through each row and column in the data array.
  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      if (values[i][j] === sstring) {
        return values[i][j + 1]; // Return the value of the cell adjacent to the found string.
      }
    }
  } // Iterate through the array to find the cell that contains the search string.
  return null; // Return null if the string is not found in any cell.
}

/**
 * Creates pivot tables in a new sheet within the specified spreadsheet for different statistical functions.
 *
 * @param {string} ssid - The ID of the spreadsheet where the pivot table will be created.
 */
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
 * @param {string} sheetname - The name of the new sheet that will contain the pivot table.
 * @param {number} rowgroupindex - The index of the column to group by in rows (1-indexed).
 * @param {number} colgroupindex - The index of the column to group by in columns (1-indexed).
 * @param {string} pivotfunction - The function to apply to the pivot table values (e.g., 'SUM', 'COUNTA').
 */
function createPivotTable(ssid, sheetname, rowgroupindex, colgroupindex, pivotfunction) {
  const ss = SpreadsheetApp.openById(ssid); // Open the spreadsheet by ID.
  const ts = ss.getSheetByName('TRADE SUMMARY'); // Access the 'TRADE SUMMARY' sheet within the spreadsheet.

  if (!ts) { throw new Error("TRADE SUMMARY sheet does not exist"); } // Check if the 'TRADE SUMMARY' sheet exists.

  const tr = ts.getDataRange(); // Get the data range of the 'TRADE SUMMARY' sheet.
  const ps = ss.insertSheet(sheetname); // Create a new sheet for the pivot table.
  const pt = ps.getRange('A1').createPivotTable(tr); // Create the pivot table in the new sheet.

  
  pt.addRowGroup(rowgroupindex); // Configure the pivot table row group.
  pt.addColumnGroup(colgroupindex); // Configure the pivot table column group.
  pt.addPivotValue(3, pivotfunction); // Configure the pivot table pivot values.

  // Format the pivot table.
  const dr = ps.getDataRange(); // Get the data range after pivot creation.
  ps.setFrozenRows(2); // Freeze the first two (header) rows.
  ps.setFrozenColumns(1); // Freeze the header column.
  dr.setFontFamily('Oswald'); // Use the 'Oswald' font in the data range.
  dr.setBorder(true, true, true, true, true, true); // Apply a border to every cell in the data range.
  ps.setHiddenGridlines(true); // Hide gridlines.

  if (sheetname !== 'NUMPIVOT') {
    dr.setNumberFormat('$#,##0.00'); // Format currencies as currency.
    ps.getRange("1:2").setNumberFormat("0"); // Format header row numbers as numbers.
  } // Format currencies as currency, if applicable.
}

/**
 * Creates a pivot table in the specified spreadsheet.
 *
 * @param {string} ssid - The ID of the spreadsheet where the pivot table will be created.
 * @param {number} rowgroupindex - The index of the column to group by in rows (1-indexed).
 * @param {number} valueindex - The index of the column to summarize in the pivot table (1-indexed).
 */
function createPivotTableOnSummary(ssid, rowgroupindex, valueindex) {
  const ss = SpreadsheetApp.openById(ssid); // Open the spreadsheet by ID.
  const ts = ss.getSheetByName('TRADE SUMMARY'); // Access the 'TRADE SUMMARY' sheet within the spreadsheet.

  if (!ts) { throw new Error("TRADE SUMMARY sheet does not exist"); } // Check if the 'TRADE SUMMARY' sheet exists.

  const dr = ts.getDataRange(); // Get the data range of the 'TRADE SUMMARY' sheet.
  const tp = 'TICKER PERFORMANCE'; // Name of the pivot sheet.

  if (ss.getSheetByName(tp)) { ss.deleteSheet(ss.getSheetByName(tp)); } // Check if the pivot sheet already exists and delete it if it does.

  const ps = ss.insertSheet(tp); // Create a new sheet for the pivot table.
  const pt = ps.getRange('A1').createPivotTable(dr); // Create the pivot table in the new sheet.
  const rg = pt.addRowGroup(rowgroupindex); // Configure the pivot table row group with sorting.

  rg.showTotals(true);
  rg.sortDescending();
  pt.addPivotValue(valueindex, SpreadsheetApp.PivotTableSummarizeFunction.SUM); // Configure the pivot table pivot values.

  // Format the pivot table.
  const pr = ps.getDataRange(); // Get the data range after pivot creation.
  ps.setFrozenRows(2); // Freeze the first two (header) rows.
  ps.setFrozenColumns(1); // Freeze the header column.
  pr.setFontFamily('Oswald'); // Use the 'Oswald' font in the data range.
  pr.setBorder(true, true, true, true, true, true); // Apply a border to every cell in the data range.
  ps.setHiddenGridlines(true); // Hide gridlines.
  pr.setNumberFormat('$#,##0.00'); // Format currencies as currency.
  ps.getRange("1:2").setNumberFormat("0"); // Format header row numbers as numbers.
}