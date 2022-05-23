Sub stock_summary()

'Declare worksheet variable to start loop
Dim ws As Worksheet

'Start loop through all worksheets
For Each ws In Worksheets

'Declare all variables in Subroutine
Dim closing_value As Double
Dim opening_value As Double
Dim ticker As String
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_vol As Double
Dim summary_table_row As Integer
Dim finalrow As Long
Dim lastrow As Long
Dim greatest_ticker As String
Dim greatest_decrease As Double
Dim greatest_increase As Double

'Define variables with values
total_stock_volume = 0
summary_table_row = 2

'Add Column Headers to display Stock Info
  ws.Cells(1, "I").Value = "Ticker"
  ws.Cells(1, "J").Value = "Yearly Change"
  ws.Cells(1, "K").Value = "Percent Change"
  ws.Cells(1, "L").Value = "Total Stock Volume"

'Find last row used in worksheets
    finalrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
  
'Begin i loop using 2 (skip headers) and lastrow
  For i = 2 To finalrow
  
'Where opening and closing values will start in loop
    closing_value = ws.Cells(i, "F").Value
    opening_value = ws.Cells(i, "C").Value

'Find how to check if ticker value is different (class notes)
    If ws.Cells(i + 1, "A").Value <> ws.Cells(i, "A").Value Then

'Define ticker name
      ticker = ws.Cells(i, "A").Value

'Print ticker name in summary "Ticker" column
      ws.Range("I" & summary_table_row).Value = ticker
      
'Add stock volumes of ticker together
      total_stock_volume = total_stock_volume + ws.Cells(i, "G").Value

'Print total stock volume for ticker in summary "Total Stock Volume" column
      ws.Range("L" & summary_table_row).Value = total_stock_volume

'Calculate yearly change using opening and closing values
      yearly_change = (closing_value - opening_value)

'Print Yearly change in summary "Yearly Change" column
      ws.Range("J" & summary_table_row).Value = yearly_change

'Check for undefined numbers in percent change
        If opening_value = 0 Then
            
            percent_change = yearly_change

'Calculate the percent change of ticker
        Else
            
            percent_change = yearly_change / opening_value
        
        End If
        
'Print Percent change in summary "Percent Change" column
      ws.Range("K" & summary_table_row).Value = percent_change
      
'Go down the summary table rows by 1 for next ticker
      summary_table_row = summary_table_row + 1

'Reset variables for next ticker count
      total_stock_volume = 0
      
'Add to total stock volume when ticker is the same value
    Else
'Adds to the total stock volume if ticker is the same
      total_stock_volume = total_stock_volume + ws.Cells(i, "G").Value
    
    End If
     
'Color code the percent change (Green for positive)
        If ws.Cells(i, "K") >= 0 Then
        
        ws.Cells(i, "K").Interior.ColorIndex = 10
          
'Color code the percent change (Red for less than 0)
        ElseIf ws.Cells(i, "K") < 0 Then
        
        ws.Cells(i, "K").Interior.ColorIndex = 3
          
     End If
        
        ws.Cells(i, "K").NumberFormat = "0.00%"
        
    Next i
 
'Add cell header values for bonus part
  ws.Range("N2").Value = "Greatest % Increase"
  ws.Range("N3").Value = "Greatest % Decrease"
  ws.Range("N4").Value = "Greatest Total Volume"
  ws.Range("O1").Value = "Ticker"
  ws.Range("P1").Value = "Value"

'Find the last row in summary table to create new for loop
  lastrow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row

'Create new for loop using variable for rows
  For i = 2 To lastrow
    
'Calculate greatest % increase with ticker
    If ws.Cells(i, "K").Value > ws.Cells(2, "P").Value Then
    
    ws.Cells(2, "P").Value = ws.Cells(i, "K").Value
    
    ws.Cells(2, "O").Value = ws.Cells(i, "I").Value
    
    End If
    
'Calculate greatest % decrease with ticker
    If ws.Cells(i, "K").Value < ws.Cells(3, "P").Value Then
    
    ws.Cells(3, "P").Value = ws.Cells(i, "K").Value
    
    ws.Cells(3, "O").Value = ws.Cells(i, "I").Value
    
    End If
    
'Calculate greatest total volume with ticker
    If ws.Cells(i, "L").Value > ws.Cells(4, "P").Value Then
    
    ws.Cells(4, "P").Value = ws.Cells(i, "L").Value
    
    ws.Cells(4, "O").Value = ws.Cells(i, "I").Value
    
    End If
    
  Next i
    
'Cleaning up worksheet info and formatting bonus data w/functions
  ws.Columns("I:P").AutoFit
  ws.Range("P2").NumberFormat = "0.00%"
  ws.Range("P3").NumberFormat = "0.00%"
  
Next ws

End Sub
