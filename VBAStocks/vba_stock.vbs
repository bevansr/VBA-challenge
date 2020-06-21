Sub yearly_stock_summary()
    
    ' Loop through all sheets in the workbook
    For Each ws In Worksheets
        
        ' Create summary header row
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Volume"
        
        ' Populate title cells in min max section
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ' Last row variable used as we loop through each ws
         last_row = ws.Cells(Rows.Count, "A").End(xlUp).Row
         
        ' last_row single sheet
        ' last_row = Cells(Rows.Count, "A").End(xlUp).Row
         
        ' Declare variables for Stock Summary
         Dim yearly_change As Double
         Dim stock_open As Double
         Dim stock_close As Double
         Dim percent_change As Double
         Dim total_volume As Double
         Dim summary_row As Integer
         
         'Set initial values for stock_open and summary row
         stock_open = ws.Cells(2, 3).Value
         total_volume = 0
         summary_row = 2
         
         ' Loop through all stocks in the sheet
         For i = 2 To last_row

            ' Check for ticker change
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Add ticker to summary
                ws.Range("I" & summary_row).Value = ws.Cells(i, 1).Value
                
                ' Lock in stock_close and calculate yearly_change and percent_change
                stock_close = ws.Cells(i, 6).Value
                'MsgBox (stock_open)
                'MsgBox (stock_close)
                yearly_change = stock_close - stock_open
                'MsgBox (yearly_change)
               If stock_open = 0 And yearly_change = 0 Then
                    percent_change = 0
                
               ElseIf stock_open = 0 Then
                    percent_change = Str("Not Applicable")
               
               Else
                    percent_change = yearly_change / stock_open * 100
                
               End If
                
                ' Add yearly_change and percent_change to summary
                ws.Range("J" & summary_row).Value = yearly_change
                ws.Range("K" & summary_row).Value = percent_change & "%"
                
                ' Conditional formatting for yearly_change summary value
                If yearly_change > 0 Then
                   ws.Range("J" & summary_row).Interior.ColorIndex = 4
                
                ElseIf yearly_change < 0 Then
                   ws.Range("J" & summary_row).Interior.ColorIndex = 3
                
                Else
                    ws.Range("J" & summary_row).Interior.ColorIndex = 6
                
                End If
                
                ' Update total_volume and add to summary
                total_volume = total_volume + ws.Cells(i, 7).Value
                ws.Range("L" & summary_row).Value = total_volume
                
                ' Reset variables for next stock and increment summary_row
                stock_open = ws.Cells(i + 1, 3).Value
                total_volume = 0
                summary_row = summary_row + 1
                
            Else
                total_volume = total_volume + ws.Cells(i, 7).Value
                
                    If stock_open = 0 Then
                        stock_open = ws.Cells(i + 1, 3).Value
                    End If
            End If
         Next i
         
        'Declare variables to store ticker for max values and store the first ticker by default
        max_increase_ticker = ws.Range("I2").Value
        max_decrease_ticker = ws.Range("I2").Value
        max_volume_ticker = ws.Range("I2").Value
    
        max_increase = ws.Cells(2, 11).Value
        max_decrease = ws.Cells(2, 11).Value
        max_volume = ws.Cells(2, 12).Value
    
        'MsgBox (summary_row - 2)
        ' Iterate through all of the values in the summary until the 2nd to last row. The last row represents j + 1 in condition.
        For j = 2 To summary_row - 2
    
            ' Determine greatest % increase
            If ws.Cells(j + 1, 11).Value > max_increase Then
                max_increase = ws.Cells(j + 1, 11).Value
                max_increase_ticker = ws.Cells(j + 1, 9).Value
            End If
        
            ' Determine greatest % decrease
            If ws.Cells(j + 1, 11).Value < max_decrease Then
                max_decrease = ws.Cells(j + 1, 11).Value
                max_decrease_ticker = ws.Cells(j + 1, 9).Value
            End If
        
            ' Determine greatest total volume
            If ws.Cells(j + 1, 12).Value > max_volume Then
                max_volume = ws.Cells(j + 1, 12).Value
                max_volume_ticker = ws.Cells(j + 1, 9).Value
            End If
    
        Next j
       
       'Convert to percentages
       max_increase = max_increase * 100
       max_decrease = max_decrease * 100
       
        'Print max/min summary for each sheet
        ws.Range("P2").Value = max_increase_ticker
        ws.Range("P3").Value = max_decrease_ticker
        ws.Range("P4").Value = max_volume_ticker
        ws.Range("Q2").Value = max_increase & "%"
        ws.Range("Q3").Value = max_decrease & "%"
        ws.Range("Q4").Value = max_volume
    Next ws
    
End Sub