# Stock Analysis Script

This script is designed to analyze stock data for each quarter in an Excel workbook. It calculates the quarterly change and percentage change from the opening price to the closing price for each stock. Additionally, it identifies the stock with the greatest percentage increase, greatest percentage decrease, and greatest total volume for each quarter.

## How to Use

1. **Download the Script**: Download the provided VBA script files.
2. **Open Excel Workbook**: Open the Excel workbook containing the stock data.
3. **Enable Macros**: Enable macros in Excel to allow the script to run.
4. **Run the Script**: Run the script by executing it from the Excel Macros menu.
5. **View Results**: After execution, the scripts will generate the analysis for each quarter, including the stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume. Negative and positive increments of quarterly change will be highlighted in red and green colours separately. The analysis will be displayed on the same sheet of each sheet within the workbook.

## Script Functionality

The script performs the following tasks:
- **Loop Through Stocks**: It loops through all the stocks for each quarter in the workbook.
- **Calculate Changes**: For each stock, it calculates the quarterly change and percentage change from the opening price to the closing price.
- **Identify Extremes**: It identifies the stock with the greatest percentage increase, greatest percentage decrease, and greatest total volume for the stock.
- **Conditional Formatting**: The script applies conditional formatting to highlight positive changes in green and negative changes of quarterly changes in red for better visualization.
- **Output Analysis**: The analysis results are displayed in the same sheet of each sheet within the workbook.

## Adjustments for Multiple Worksheets

The script has been adjusted to run on every worksheet (quarter) in the workbook. It iterates through each worksheet to analyze the stock data for each quarter individually.

## Dependencies

- Microsoft Excel
- Enabled Macros

## Compatibility

This script is compatible with Microsoft Excel. Ensure that macros are enabled in your Excel settings to run the script successfully.
