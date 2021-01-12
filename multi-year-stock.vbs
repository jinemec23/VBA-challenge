Sub TestDataLoop()
'Set variables
Dim ws As Worksheet
Dim ticker As String
Dim price_change As Double
price_change = 0
Dim percent_change As Double
percent_change = 0
Dim vol As Double
vol = 0
'BONUS VARIABLES
    Dim Max_Percent_name As String
    Max_Percent_name = " "
    Dim Min_Percent_name As String
    Min_Percent_name = " "
    Dim MAX_PERCENT As Double
    MAX_PERCENT = 0
    Dim MIN_PERCENT As Double
    MIN_PERCENT = 0
    Dim max_volume_name As String
    max_volume_name = " "
    Dim MAX_VOLUME As Double
    MAX_VOLUME = 0




'Loop through all sheets
For Each ws In Worksheets

'Set Headers
ws.Range("i1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest total volume"

'Set loop for summary
Dim summary_table_row As Long
summary_table_row = 2

Dim lastrow As Long
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


'Loop through all tickers
For i = 2 To lastrow

    'Check for duplicates to summarize
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        'set values
        ticker = ws.Cells(i, 1).Value
        
        'Set calculations
        open_price = ws.Cells(i, 3).Value
        close_price = ws.Cells(i, 6).Value
        price_change = close_price - open_price
        vol = vol + ws.Cells(i, 7).Value
        
       If open_price <> 0 Then
       percent_change = (price_change / open_price) * 100
       End If

      
        'Print data in summary row table
        ws.Range("I" & summary_table_row).Value = ticker
        ws.Range("J" & summary_table_row).Value = price_change
        ws.Range("K" & summary_table_row).Value = (CStr(percent_change) & "%")
        ws.Range("L" & summary_table_row).Value = vol
        
     'Format Cells
         If (price_change > 0) Then
             ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
             
        ElseIf (price_change <= 0) Then
             ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
             
         End If
        
        'add 1 to summary row
        summary_table_row = summary_table_row + 1
        
  
    


    'BONUS CALCULATIONS
      If (percent_change > MAX_PERCENT) Then
                    MAX_PERCENT = percent_change
                    Max_Percent_name = ticker
                    
                ElseIf (percent_change < MIN_PERCENT) Then
                    MIN_PERCENT = percent_change
                    Min_Percent_name = ticker
                End If
                       
                If (vol > MAX_VOLUME) Then
                    MAX_VOLUME = vol
                    max_volume_name = ticker
                End If
                
        'RESET VALUES
        percent_change = 0
        vol = 0
    
    Else
      vol = vol + ws.Cells(i, 7).Value
      
      End If
      
      Next i
      
    'BONUS PRINT
    ws.Range("Q2").Value = (CStr(MAX_PERCENT) & "%")
    ws.Range("Q3").Value = (CStr(MIN_PERCENT) & "%")
    ws.Range("Q4").Value = MAX_VOLUME
    ws.Range("P2").Value = Max_Percent_name
    ws.Range("P3").Value = Min_Percent_name
    ws.Range("P4").Value = max_volume_name
    



Next

End Sub