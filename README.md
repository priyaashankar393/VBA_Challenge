# VBA_Challenge

VBA _Challenge Assignment2

Introduction:

In this project, I wrote VBA code to automate an Excel workbook to pull relevant data from thousands of rows in multiple worksheets and output summary columns of the stocks. This is done using a macro and VBA script. This script retrieves values, calculates simple arithmetic functions, and returns summary tables of the stock market data.

Input Files - Multiple_year_stock_data.xlsm - multiple year (2018,2019,2020) stock data
 
VBA File Name - " VBA_Challenge.cls" VBA code to run stock market macro (Also available in word)

Implementation:

* To read the VBA script: open "VBA_Challenge.cls"
* Created a "button" function to run the macro script from the spreadsheet
* Declared all the required variables with appropriate data types such as doubles, floats, string
* "lastrow" function utilizes the "xlup" function: "ws.Cells(Rows.Count, 1).End(xlUp).Row"
* Used For loop to find, aggregate, and store values
* Populated Range/cells with desired outputs 
* Looped through the worksheets within the workbook using "ws." to define the current worksheet, then looping through the next worksheet with "Next ws" function
*Finally completed the code with the End function.
