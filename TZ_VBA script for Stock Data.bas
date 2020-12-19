Attribute VB_Name = "Module1"
Sub Stock_Market(ws)
    'set variables
    Dim ticker As String
    Dim total_volume As LongLong
    Dim summary_table_row As Integer
    total_volume = 0
    
    Dim year_open As Double
    Dim year_close As Double
    Dim yearly_change As Double
    Dim precent_change As Double
    year_open = "0.00"
    year_close = "0.00"

    
    Dim max_value As Double
    max_value = 0
    Dim min_value As Double
    min_value = 0
    Dim max_ticker As String
    Dim min_ticker As String
    Dim max_volume As Double
    max_volume = 0
    Dim max_volume_ticker As String

    'set location for each ticket in the summary table
    summary_table_row = 2
    
     'set the last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Loop through data
    For i = 2 To LastRow
    
        'Check if we are within the same ticker data, if not then..
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            
            'set the open cell
            year_open = ws.Cells(i, 3).Value
        
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'set the ticker cell
            ticker = ws.Cells(i, 1).Value
            
            'set the close cell
            year_close = ws.Cells(i, 6).Value
            
            'calculate the change
            yearly_change = year_close - year_open
            
            
            
            If year_open <> 0 Then
                
                'Get the precent of change
                precent_change = (yearly_change / year_open)
                
            End If
            
            
            
            'add the total
            total_volume = total_volume + ws.Cells(i, 7).Value
            
            'Print the ticker in the summary table
            ws.Range("I" & summary_table_row).Value = ticker
            
            'print the total in the summary table
            ws.Range("L" & summary_table_row).Value = total_volume
            
            'print the yearly change in the summary table
            ws.Range("J" & summary_table_row).Value = yearly_change
            
            'print the precent change in the summary table and format
            ws.Range("K" & summary_table_row).Value = precent_change
            ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
            
            
            
            'Conditional formatting for plus/negative change
            If (yearly_change > 0) Then
        
                ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
            
            ElseIf (yearly_change <= 0) Then
            
                ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
            
            End If
            
            
            
            'Bonus: calc. for min and max
            If (precent_change > max_value) Then
                
                max_value = precent_change
                max_ticker = ticker
            
            ElseIf (precent_change < min_value) Then
                
                min_value = precent_change
                min_ticker = ticker
            
            End If
            
            'Bonus: print the min/max values and ticker names
            ws.Range("P2").Value = max_ticker
            ws.Range("Q2").Value = max_value
            ws.Range("P3").Value = min_ticker
            ws.Range("Q3").Value = min_value
            ws.Range("Q2:Q3").NumberFormat = "0.00%"
            
            
            'Bonus: calc. for total
            If (total_volume > max_volume) Then
            
                max_volume = total_volume
                max_volume_ticker = ticker
            
            End If
            
            'Bonus: print the max total volume and ticker name
            ws.Range("Q4").Value = max_volume
            ws.Range("P4").Value = max_volume_ticker
            
            
            
            
            
            'add one to the summary table row
            summary_table_row = summary_table_row + 1
            
            'reset the total
            total_volume = 0
            yearly_change = 0
            precent_change = 0
            
        'if the cell immediately following a row is the same ticker...
        Else
        
            total_volume = total_volume + ws.Cells(i, 7).Value
        
        End If
        
    Next i

End Sub

Sub Copy_Year()

 For Each ws In Worksheets
    
    Stock_Market ws
        
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    
    
    ws.Columns("I:L").AutoFit
        
    Next ws

End Sub
