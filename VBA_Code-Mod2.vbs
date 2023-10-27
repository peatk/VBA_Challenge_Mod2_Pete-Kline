Sub StockData():

'Set variable to hold ticker
Dim ticker As String

'Set variable to hold formulas - year change, % change, and stock volume
Dim year_change As Double
year_change = 0

Dim percent_change As Double
percent_change = 0

Dim total_stock_vol As Double
total_stock_vol = 0

'Track location for each ticker calculation
Dim summary_table_row As Integer    
summary_table_row = 2

'Set values for required columns to analyze
Dim greatest_increase As Double
Dim greatest_decrease As Double
Dim greattes_total_vol As Double

'help from ed
Dim start_value As Double
Dim end_value As Double

'Set num1 value - Used in count function for year_change
Dim num1 As Double

'For each loop to sift through stock data
For Each ws In Worksheets
    
    'set year starting value
    start_value = ws.Cells(2, 3).Value
    
    'Set last row
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'For loop through the data
    For i = 2 To lastRow
       
    'If cell ticker does not match do this..if not do that
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
    'Set values
    ticker = ws.Cells(i, 1).Value
    end_value = ws.Cells(i, 6).Value
    
    'Add to year change
    year_change = end_value - start_value
    
    'Add last line of total stock vol
    total_stock_vol = total_stock_vol + ws.Cells(i, 7).Value
    
   'Print ALL headers
    ws.Range("I1,O1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
   
    ws.Range("P1").Value = "Value"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest % Total Volume"
    
    'Print values in columns
    ws.Range("I" & summary_table_row).Value = ticker
    ws.Range("J" & summary_table_row).Value = year_change
    ws.Range("K" & summary_table_row).Value = year_change / start_value
    ws.Range("L" & summary_table_row).Value = total_stock_vol
    
    'Add to summary table row, and push down by 1 cell
    summary_table_row = summary_table_row + 1

    'Reset all counting values
    year_gchange = 0
    num1 = 0
    total_stock_vol = 0
    greatest_increase = 0
    greatest_decrease = 0
    greatest_total_vol = 0
    start_value = ws.Cells(i + 1, 3)
    
    Else
    total_stock_vol = total_stock_vol + ws.Cells(i, 7).Value
    
    End If
         
Next i

'moved greatest values outside of loop
'Add greatest increase, decrease and vol
greatest_increase = WorksheetFunction.Max(ws.Range("K1:K3500"))
greatest_decrease = WorksheetFunction.Min(ws.Range("K1:K3500"))
greatest_total_vol = WorksheetFunction.Max(ws.Range("L1:L3500"))

'Print greatest values
ws.Range("P2").Value = greatest_increase
ws.Range("P3").Value = greatest_decrease
ws.Range("P4").Value = greatest_total_vol

'Lookup ticker value based up on P2-P4 data. enter into respective cells
ws.Range("O2").Value = WorksheetFunction.Index(ws.Range("I1:I3500"), WorksheetFunction.Match(ws.Range("P2").Value, ws.Range("K1:K3500"), 0))
ws.range("O3").Value = WorksheetFunction.Index(ws.Range("I1:I3500"), WorksheetFunction.Match(ws.range("P3").Value, ws.Range("K1:K3500"), 0))
ws.range("O4").Value = WorksheetFunction.Index(ws.Range("I1:I3500"), WorksheetFunction.Match(ws.range("P4").Value, ws.Range("L1:L3500"), 0))
    
    'For loop - formatting red (-) & green (+)
    lastrow2 = ws.Cells(Rows.Count, 10).End(xlUp).Row
    For c = 2 To lastrow2
        
        'If cells are < 0 color red if not, color green
        If ws.Cells(c, 10).Value < 0 Then
        ws.Cells(c, 10).Interior.ColorIndex = 3
        
        Else
        ws.Cells(c, 10).Interior.ColorIndex = 4
        
        End If
    Next c

'Reset summary table for each sheet
summary_table_row = 2

'Format cells
ws.Range("K1:K3500").NumberFormat = "0.00%"
ws.Range("P2:P3").NumberFormat = "0.00%"
ws.Cells.EntireColumn.AutoFit

Next ws

End Sub






