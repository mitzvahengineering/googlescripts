# Google Apps Script: Trading Day Filter and Summary

This script is designed to work with Google Sheets to filter trading data for a specific day and create a summary with a pivot table. It is particularly useful for financial analysts and traders who need to organize and analyze trading data efficiently.

## Features

- **Filter Trading Data:** Filters trading data for a specific trading day.
- **Automated Sheet Creation:** Creates new sheets for filtered data and a pivot table.
- **Dynamic Coloring:** Applies different text colors to rows based on the trading account.
- **Data Formatting:** Customizes cell formats, including currency formatting and date/time formats.
- **Pivot Table Summary:** Automatically generates a pivot table to summarize the trading data.

## Setup

1. **Open Google Sheets**: Use a spreadsheet that contains your trading data.
2. **Script Editor**: Go to `Extensions > Apps Script` in Google Sheets to open the script editor.
3. **Paste the Script**: Copy the provided script into the script editor.
4. **Save and Close**: Save the script and close the script editor.

## Usage

To run the script:

1. **Open the Spreadsheet**: Ensure you are in the spreadsheet with your trading data.
2. **Run the Script**: Go to `Extensions > Apps Script` and run the `filterTradingDay` function.
3. **View Results**: Check the newly created sheets for filtered trading data and the pivot table summary.

## Functions

- **filterTradingDay()**: Filters trading data for a predefined trading date, deletes existing sheets for that date, creates a new sheet for the filtered data, and formats the data. It ends by calling `createPivotTable` to generate a pivot table.
- **createPivotTable(sheet)**: Creates a pivot table in a new sheet based on the filtered data provided by `filterTradingDay`.

## Notes

- The script uses a fixed trading date. You may modify the `tradingdate` variable to suit your needs.
- The script assumes the trading data is in a specific format. Ensure your data matches the expected format for accurate results.
- The script dynamically assigns colors to rows based on the trading account, cycling through a predefined list of colors.

## Customization

You can customize the script by altering the trading date, adjusting the data range, changing the text colors, or modifying the data formatting as per your requirements.

---
