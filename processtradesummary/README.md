# Google Sheets Trade Processing Script

## Overview
This script is designed to process and analyze trade data from a series of Google Sheets within a specified folder. It aggregates data into a summary report, creates pivot tables for data analysis, and sets up the sheets with specified formatting and layout.

## Main Features
- **Process Folder**: Iterates over Google Sheets in a specified folder, extracting and processing data.
- **Setup Summary Sheet**: Formats the summary sheet in the output spreadsheet.
- **Create Pivot Tables**: Generates various pivot tables for data analysis in the output spreadsheet.

## Functions

### processFolder
Main function to start processing. It sets up the environment, processes each spreadsheet in the specified folder, and creates pivot tables in the output spreadsheet.

- **Parameters**: None
- **Usage**: Automatically processes spreadsheets in a predefined folder.

### setupSummarySheet
Sets up the summary sheet in the output spreadsheet with headers, formulas, and formatting.

- **Parameters**:
  - `summarysheet` (SpreadsheetApp.Sheet): The sheet to be set up as the summary sheet.
- **Usage**: Called within `processFolder` to format the summary sheet.

### processInputFile
Processes each sheet in an input spreadsheet, extracting relevant data and appending it to the output spreadsheet.

- **Parameters**:
  - `ipssid` (String): ID of the input spreadsheet.
  - `opssid` (String): ID of the output spreadsheet.
- **Usage**: Called within `processFolder` for each spreadsheet in the folder.

### stringSearch
Searches for a specific string in a given sheet and returns the value in the adjacent cell.

- **Parameters**:
  - `ss` (SpreadsheetApp.Spreadsheet): The spreadsheet object.
  - `tab` (String): The name of the sheet within the spreadsheet to search.
  - `sssearch` (String): The string to search for within the sheet.
- **Returns**: Value of the adjacent cell or null if not found.

### findCellContainingString
Finds the first cell in a sheet that contains a given string.

- **Parameters**:
  - `tabObject` (SpreadsheetApp.Sheet): The sheet to search within.
  - `sssearch` (String): The string to search for within the sheet.
- **Returns**: The cell containing the string or null if not found.

### createPivotTables
Creates multiple pivot tables in the output spreadsheet for data analysis.

- **Parameters**:
  - `ssid` (String): The ID of the spreadsheet where pivot tables will be created.
- **Usage**: Called within `processFolder` after processing all input files.

### createPivotTable
Creates a single pivot table in a new sheet within the output spreadsheet.

- **Parameters**:
  - `ssid` (String): ID of the spreadsheet.
  - `sheetName` (String): Name for the new sheet containing the pivot table.
  - `rowGroupIndex` (Number): Column index for row grouping.
  - `colGroupIndex` (Number): Column index for column grouping.
  - `pivotFunction` (String): Function to apply to pivot values.
- **Usage**: Part of `createPivotTables` for creating individual pivot tables.

## How to Use
1. Set the folder identifiers (`testfolderid`, `realfolderid`, `opfolderid`) in the `processFolder` function.
2. Run the `processFolder` function to start processing the spreadsheets in the specified folder.
3. Check the output spreadsheet in the defined dump folder for the summary report and pivot tables.

## Requirements
- Google Apps Script environment.
- Spreadsheets in Google Sheets format.
- Access rights to the specified folders and spreadsheets.

