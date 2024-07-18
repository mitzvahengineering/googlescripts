/**
 * Trade Summary and Analysis Script
 * 
 * Purpose:
 * This script processes multiple Google Sheets files containing trade data,
 * consolidates the information into a summary spreadsheet, and generates
 * various pivot tables for analysis.
 * 
 * How to use:
 * 1. Set up the config object with appropriate folder IDs:
 *    - testfolder: ID of the folder containing test data
 *    - realfolder: ID of the folder containing production data
 *    - dumpfolder: ID of the folder where the summary spreadsheet will be saved
 * 2. Set the 'production' flag in the config object:
 *    - false: runs the script on the test folder
 *    - true: runs the script on the real (production) folder
 * 3. Ensure the 'textsearch' value in the config object matches the label for total trade value in your sheets
 * 4. Run the 'processFolder' function to execute the script
 * 5. The script will create a new spreadsheet in the specified dump folder with:
 *    - A summary sheet of all processed trades
 *    - Multiple pivot tables for different analyses
 * 
 * Note: This script requires the 'Drive API' Advanced Drive Service to be enabled.
 */

// Configuration object
const config = {
  production: false, // Set to false for testing
  testfolder: '1z38V8POr9lXNAoFBM-7I5_8GG9t5E2wx',
  realfolder: '1k0FhOtK-3_mGoH5CnC9axY2l2FsP7h1q',
  dumpfolder: '1zFsCLutd5pboDamsClZZBRsvPHB6QDwu',
  textsearch: 'Total'
};

/**
 * Main function to process spreadsheets in a specified folder and generate a summary report.
 * It processes all Google Sheets files in the specified folder, extracts relevant data, and creates a summary sheet
 * with statistical information and pivot tables.
 */
function processFolder() {
  const infolderid = config.production ? config.realfolder : config.testfolder; // Identifier of the folder of trades to process.
  const reportdate = Utilities.formatDate(new Date(), 'GMT', 'yyyyMMdd'); // Datestamp for output file (format: yyyyMMdd).
  const foldername = DriveApp.getFolderById(infolderid).getName().toLowerCase(); // Name of the input folder.
  const opfilename = `summary-of-${foldername}-${reportdate}`.replace(/\s/g, '-'); // Output filename.
  const opmetadata = {
    name: opfilename, // Declare the output filename.
    mimeType: MimeType.GOOGLE_SHEETS, // Declare that we are creating a spreadsheet.
    parents: [config.dumpfolder] // Folder id of the newly created spreadsheet.
  }; // Metadata for creating the output file in the desired folder.

  try {
    // Create a new Google Sheets file for the summary report.
    const opmetafile = Drive.Files.create(opmetadata); // Requires 'Drive API' Advanced Drive Service.
    const opdatafile = SpreadsheetApp.openById(opmetafile.id); // Open the newly created spreadsheet.
    setupSummarySheet(opdatafile.getSheets()[0]); // Prepare and format the summary sheet.
    
    const datafileid = opdatafile.getId(); // Define the ID of the output data file.
    const inputfiles = DriveApp.getFolderById(infolderid).getFilesByType(MimeType.GOOGLE_SHEETS); // Get all Google Sheets files in the input folder.
    
    let filenumber = 0;
    while (inputfiles.hasNext()) { 
      processInputFile(inputfiles.next().getId(), datafileid); 
      filenumber++;
      if (filenumber % 10 === 0) Logger.log(`Processed ${filenumber} files`);
    } // Process each file in the input folder.
    
    const ss = SpreadsheetApp.openById(datafileid); // Open the spreadsheet with datafileid for processing.
    const ts = ss.getSheetByName('TRADE SUMMARY'); // Reopen the summary sheet to ensure we have the latest data.
    
    if (ts && ts.getLastRow() > 1) {
      formatSummarySheet(ts);
    } else {
      Logger.log("No data to format in summary sheet");
    } // Format the summary sheet after processing all input files.
    
    createPivotTables(datafileid); // Create pivot tables in the summary report.

    const tp = ss.getSheetByName('TICKER PERFORMANCE'); // Get the TICKER PERFORMANCE sheet.
    if (tp) { ss.moveActiveSheet(ss.getSheets().indexOf(tp) + 1); } // Move the TICKER PERFORMANCE sheet to first position (index 0).

    Logger.log("Processing completed successfully.");
  } catch (error) {
    Logger.log(`Error in processFolder: ${error.message}`);
  }
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

  summarysheet.setName('TRADE SUMMARY').setHiddenGridlines(true).getDataRange().setFontFamily('Oswald'); 
  summarysheet.getRange(1, 1, rowheaders.length, 3).setBackground('#000000').setFontColor('#ffffff'); // Set background and font color for the header area
  summarysheet.getRange(1, 2, rowheaders.length, rowheaders[0].length).setValues(rowheaders); // Set values for row headers and formulas
  summarysheet.getRange(1, 2, rowheaders.length, 1).setHorizontalAlignment('right'); // Align row headers (first column) to the right
  summarysheet.setFrozenRows(6);
}

/**
 * Processes each sheet in the input spreadsheet and appends relevant data to the output spreadsheet.
 *
 * @param {string} ipssid - The ID of the input spreadsheet.
 * @param {string} opssid - The ID of the output spreadsheet.
 */
function processInputFile(ipssid, opssid) {
  try {
    const ipss = SpreadsheetApp.openById(ipssid); // Open the input spreadsheet by its ID.
    const opss = SpreadsheetApp.openById(opssid); // Open the output spreadsheet by its ID.
    const opssname = opss.getSheetByName('TRADE SUMMARY');
    const ipssname = ipss.getName(); // Retrieve the name of the input spreadsheet.
    
    if (!opssname) { throw new Error('TRADE SUMMARY sheet not found in the output spreadsheet'); }
    
    ipss.getSheets().forEach(sheet => {
      const worksheet = sheet.getName();
      Logger.log(`Processing sheet: ${worksheet}`);
      
      const result = stringSearch(sheet, config.textsearch); // Use stringSearch function to find the string in the current sheet.
      if (result !== null) {
        opssname.appendRow([ipssname, worksheet, result]); // Append the found data to the output spreadsheet.
        Logger.log(`Processed ${ipssname} TRADE ${worksheet} [ ${result} USD ].`);
      } else {
        Logger.log(`No "${config.textsearch}" found in ${ipssname} TRADE ${worksheet}`); // Log the missing value but don't append to the summary sheet
      }
    });
    
    Logger.log(`Completed processing file ${ipssid}`);
  } catch (error) {
    Logger.log(`Error processing file ${ipssid}: ${error.message}`);
  }
}

/**
 * Searches for a specific string in a given sheet and returns the value in the adjacent cell.
 *
 * @param {SpreadsheetApp.Sheet} sheet - The sheet object to search.
 * @param {string} textsearch - The string to search for within the sheet.
 * @return {number|null} - Returns the numeric value of the cell adjacent to the found string, or null if not found or not numeric.
 */
function stringSearch(sheet, textsearch) {
  try {
    const values = sheet.getDataRange().getValues();  // Get all values from the sheet as a 2D array.
    for (let row of values) {
      const index = row.indexOf(textsearch);
      if (index !== -1 && index + 1 < row.length) {
        const value = row[index + 1];
        return typeof value === 'number' ? value : null; // Return the value only if it's a number
      }
    }
    return null; // Return null if the string is not found in any cell.
  } catch (error) {
    Logger.log(`Error in stringSearch for sheet "${sheet.getName()}": ${error.message}`);
    return null;
  }
}

/**
 * Applies formatting to the summary sheet.
 *
 * @param {SpreadsheetApp.Sheet} sheet - The sheet to be formatted.
 */
function formatSummarySheet(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    Logger.log(`Formatting summary sheet. Rows: ${lastRow}, Columns: ${lastColumn}`);

    if (lastRow < 7 || lastColumn < 3) {
      Logger.log(`Warning: Summary sheet has insufficient data. Rows: ${lastRow}, Columns: ${lastColumn}`);
      return;
    }

    // Set border for all data in the sheet.
    Logger.log('Applying borders...');
    sheet.getRange(1, 1, lastRow, lastColumn).setBorder(true, true, true, true, true, true);

    // Format numbers as currency, but only if the column exists
    if (lastColumn >= 3) {
      Logger.log('Formatting currency...');
      sheet.getRange(2, 3, lastRow - 1, 1).setNumberFormat('$#,##0.00');
    }
    
    // Add alternating row colors
    if (lastRow > 7) {
      Logger.log('Applying alternating colors...');
      const colorRange = sheet.getRange(7, 1, lastRow - 6, lastColumn);
      const colors = createAlternatingColors(lastRow - 6);
      Logger.log(`Created colors array with ${colors.length} rows`);
      colorRange.setBackgroundColors(colors);
    }
    
    Logger.log(`Successfully formatted summary sheet. Rows: ${lastRow}, Columns: ${lastColumn}`);
  } catch (error) {
    Logger.log(`Error in formatSummarySheet: ${error.message}`);
    Logger.log(`Error occurred at: ${error.stack}`);
    // Log the current state of the sheet for debugging
    Logger.log(`Sheet state - Name: ${sheet.getName()}, Rows: ${sheet.getLastRow()}, Columns: ${sheet.getLastColumn()}`);
    
    // Log the content of the sheet for further debugging
    const content = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
    Logger.log(`Sheet content: ${JSON.stringify(content)}`);
  }
}

/**
 * Creates an array of alternating background colors for rows.
 *
 * @param {number} numRows - The number of rows to create colors for.
 * @return {string[][]} An array of color strings for each row.
 */
function createAlternatingColors(numRows) {
  return Array(numRows).fill().map((_, i) => Array(3).fill((i % 2 === 0) ? '#ffffff' : '#f3f3f3'));
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
  createPivotTable(ssid, 'NUMPIVOT', 1, 2, SpreadsheetApp.PivotTableSummarizeFunction.COUNT);
  createSummaryPivotTable(ssid);
}

/**
 * Creates a pivot table in a new sheet within the specified spreadsheet.
 *
 * @param {string} ssid - The ID of the spreadsheet where the pivot table will be created.
 * @param {string} sheetName - The name of the new sheet that will contain the pivot table.
 * @param {number} rowGroupIndex - The index of the column to group by in rows (1-indexed).
 * @param {number} colGroupIndex - The index of the column to group by in columns (1-indexed).
 * @param {SpreadsheetApp.PivotTableSummarizeFunction} pivotFunction - The function to apply to the pivot table values.
 */
function createPivotTable(ssid, sheetName, rowGroupIndex, colGroupIndex, pivotFunction) {
  try {
    const ss = SpreadsheetApp.openById(ssid); // Open the spreadsheet by ID (with the ssid identifier).
    const sn = ss.getSheetByName('TRADE SUMMARY'); // Access the 'TRADE SUMMARY' sheet within the spreadsheet.
    const sc = 7; // Start from row 7 to exclude headers to define the range of the Pivot Table source data.
    const lr = sn.getLastRow();
    const lc = sn.getLastColumn();
    
    if (lr < sc || lc < 3) {
      Logger.log(`Insufficient data for pivot table ${sheetName}. Rows: ${lr}, Columns: ${lc}`);
      return;
    }
    
    const pr = sn.getRange(sc, 1, lr - sc + 1, lc); // Define the range of the Pivot Table source data.
    const ps = ss.insertSheet(sheetName); // Create a new sheet for the pivot table.
    const pt = ps.getRange('A1').createPivotTable(pr); // Create the pivot table in the new sheet.
    
    pt.addRowGroup(rowGroupIndex); // Configure the pivot table row group.
    if (colGroupIndex <= lc && sheetName !== 'NUMPIVOT') {
      pt.addColumnGroup(colGroupIndex); // Configure the pivot table column group.
    }
    
    const valueColIndex = sheetName === 'NUMPIVOT' ? 2 : Math.min(3, lc);
    const pv = pt.addPivotValue(valueColIndex, pivotFunction);
    
    if (sheetName !== 'NUMPIVOT') {
      pv.setDisplayName('Amount (USD)');
    }

    formatPivotTable(ps, sheetName !== 'NUMPIVOT');
    Logger.log(`Successfully created pivot table ${sheetName}`);
  } catch (error) {
    Logger.log(`Error creating pivot table ${sheetName}: ${error.message}`);
  }
}

/**
 * Creates a summary pivot table in a new sheet within the specified spreadsheet.
 *
 * @param {string} ssid - The ID of the spreadsheet where the pivot table will be created.
 */
function createSummaryPivotTable(ssid) {
  try {
    const ss = SpreadsheetApp.openById(ssid);
    const sn = ss.getSheetByName('TRADE SUMMARY');
    const sc = 7; // Start from row 7 to exclude headers
    const lastRow = sn.getLastRow();
    const lastCol = sn.getLastColumn();
    
    if (lastRow < sc || lastCol < 3) {
      Logger.log('Insufficient data for summary pivot table');
      return;
    }
    
    const pr = sn.getRange(sc, 1, lastRow - sc + 1, lastCol);
    const ps = ss.insertSheet('TICKER PERFORMANCE');
    const pt = ps.getRange('A1').createPivotTable(pr);

    pt.addRowGroup(1).showTotals(true).setDisplayName('Ticker'); // Add row group for column A.
    pt.addPivotValue(2, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA).setDisplayName('Count'); // Add count of column B.
    pt.addPivotValue(3, SpreadsheetApp.PivotTableSummarizeFunction.SUM).setDisplayName('Amount (USD)'); // Add sum of column C and sort in descending order.
    
    Logger.log('Successfully created summary pivot table');
  } catch (error) {
    Logger.log(`Error creating summary pivot table: ${error.message}`);
  }
}

/**
 * Applies formatting to a pivot table.
 *
 * @param {SpreadsheetApp.Sheet} sheet - The sheet containing the pivot table.
 * @param {boolean} formatAsCurrency - Whether to format the values as currency.
 */
function formatPivotTable(sheet, formatAsCurrency) {
  const dr = sheet.getDataRange(); // Get the data range after pivot creation.
  sheet.setFrozenRows(2); // Freeze the first two (header) rows.
  sheet.setFrozenColumns(1); // Freeze the header column.
  dr.setFontFamily('Oswald'); // Use the 'Oswald' font in the data range.
  dr.setBorder(true, true, true, true, true, true); // Apply a border to every cell in the data range.
  sheet.setHiddenGridlines(true); // Hide gridlines.

  if (formatAsCurrency) {
    dr.setNumberFormat('$#,##0.00'); // Format currencies as currency.
    sheet.getRange("1:2").setNumberFormat("0"); // Format header row numbers as numbers.
  }

  // Add alternating row colors
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(3, 1, lastRow - 2, sheet.getLastColumn());
  range.setBackgroundColors(createAlternatingColors(lastRow - 2));
}