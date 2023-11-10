VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Requirements

'Create a script that loops through all the stocks for one year and outputs the following information:

     'The ticker symbol

      'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

      'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

     'The total stock volume of the stock. The result should match the following image:

      'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
      
      'Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once

'create function to loop entire stock data to produce outcome
Sub tickerloop()

'Start for loop through all the worksheets
For Each ws In Worksheets

'Declaration section for all variable used in this worksheet

        Dim tickername As String
        Dim tickervolume As Double
        tickervolume = 0
        Dim summary_ticker_row As Integer
        summary_ticker_row = 2
        Dim open_price As Double
        open_price = Cells(2, 3).Value
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim Max_volume As Double
        Dim Max_volume_change As Double
        Dim Max_volume_Change_name As String
        Dim Max_per_change  As Double
        Dim Max_per_change_name As String
        Dim Min_per_change  As Double
        Dim Min_per_change_name As String
        Max_volume_change = 0
        Max_per_change = 0
        Min_per_change = 0

'Column variable definition
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"

'VBA Script to implement the above requirements
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
              tickername = ws.Cells(i, 1).Value
              tickervolume = tickervolume + ws.Cells(i, 7).Value
              If (Max_volume_change < tickervolume) Then
                Max_volume_change = tickervolume
                Max_volume_chnage_name = tickername
                ws.Cells(4, 15).Value = tickername
                ws.Cells(4, 16).Value = tickervolume
              End If
            
              ws.Range("I" & summary_ticker_row).Value = tickername
              ws.Range("L" & summary_ticker_row).Value = tickervolume
              close_price = ws.Cells(i, 6).Value
              yearly_change = (close_price - open_price)
              ws.Range("J" & summary_ticker_row).Value = yearly_change
              
              If (open_price = 0) Then
                 percent_change = 0
              Else
                 percent_change = yearly_change / open_price
              End If
               
              If (Max_per_change < percent_change) Then
                Max_per_change = percent_change
                Max_per_change_name = tickername
                ws.Cells(2, 16).Value = percent_change
                ws.Cells(2, 16).NumberFormat = "0.00%"
                ws.Cells(2, 15).Value = tickername
              End If
             
            If (Min_per_change > percent_change) Then
                Min_per_change = percent_change
                Min_per_change_name = tickername
                ws.Cells(3, 16).Value = percent_change
                ws.Cells(3, 16).NumberFormat = "0.00%"
                ws.Cells(3, 15).Value = tickername
              End If
              ws.Range("K" & summary_ticker_row).Value = percent_change
              ws.Range("K" & summary_ticker_row).NumberFormat = "0.00%"
              summary_ticker_row = summary_ticker_row + 1
              tickervolume = 0
              open_price = ws.Cells(i + 1, 3)
           
            Else
              tickervolume = tickervolume + ws.Cells(i, 7).Value
            End If
       
        Next i
'To fill color on Yearly change based on the value
    lastrow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    For i = 2 To lastrow_summary_table
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 10
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
           
            Next i
            
Next ws 'Next worksheet
   
End Sub
