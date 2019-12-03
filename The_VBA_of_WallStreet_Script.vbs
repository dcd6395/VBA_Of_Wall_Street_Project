Attribute VB_Name = "Module1"
Sub Wallstreet()

    ' loop through each worksheet
    For Each WS In Worksheets
        
        
        ' find the last row of the dataset
        Dim LastRow As Long
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' create variables pt 1
        Dim ticker As String, yearly_change, percent_change, yearly_open, yearly_close, row_open, total_volume As Double
        total_volume = 0
        row_open = 2
        
        ' create variables pt 2
        Dim LastRow_PercentChange, LastRow_PercentChange2 As Long
        
        
    
    
       ' record new columns
        Dim Print_Summary As Integer
        Print_Summary = 2

        ' Add column headers
                WS.Cells(1, 9).Value = "Ticker"
                
                WS.Cells(1, 10).Value = "Yearly Change"
                
                WS.Cells(1, 11).Value = "Percent Change"
                
                WS.Cells(1, 12).Value = "Total Stock Volume"
        
       

        ' loop through dataset
        For i = 2 To LastRow
        
            total_volume = total_volume + WS.Cells(i, 7).Value
            
            ' if statement for unique ticker names
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                
                ' values of variables
                ticker = WS.Cells(i, 1).Value
                yearly_open = WS.Range("C" & row_open)
                yearly_close = WS.Range("F" & i)
                ' calculating yearly change
                yearly_change = (yearly_close - yearly_open)
                
                ' if/else statement to calculate percent change
                If yearly_open > 0 Then
                percent_change = yearly_change / yearly_open
                
                Else
                    
                    
                End If
                
                 'format percentages
                WS.Range("K" & Print_Summary).NumberFormat = "0.00%"
                
                ' print values into columns
                WS.Range("I" & Print_Summary).Value = ticker
                WS.Range("J" & Print_Summary).Value = yearly_change
                WS.Range("K" & Print_Summary).Value = percent_change
                WS.Range("L" & Print_Summary).Value = total_volume
                Print_Summary = Print_Summary + 1

        
            End If
            
            ' conditional formatting for yearly change
            ' this works in my text excel file but not the stock file
            If WS.Range("J" & Print_Summary).Value < 0 Then
            WS.Range("J" & Print_Summary).Interior.ColorIndex = 3
            
            Else
            WS.Range("J" & Print_Summary).Interior.ColorIndex = 4
        
            End If
        
        Next i
        
        ' print variables in cells
        WS.Range("O2").Value = "Greatest % Increase"
        WS.Range("O3").Value = "Greatest % Decrease"
        WS.Range("O4").Value = "Greatest Total Volume"
        WS.Range("P1").Value = "Ticker"
        WS.Range("Q1").Value = "Value"
        
        'format percentages
        WS.Range("Q2:Q3").NumberFormat = "0.00%"
        
        
        'find last row of percent changes
        LastRow_PercentChange = WS.Cells(Rows.Count, 11).End(xlUp).Row
        LastRow_PercentChange2 = WS.Cells(Rows.Count, 12).End(xlUp).Row
        
        'loop
         For i = 2 To LastRow_PercentChange
            
            ' statement to find the greatest increase and print the ticker and value in the correct cells
            If WS.Cells(i, 11) = WorksheetFunction.Max(WS.Range("K2:K" & LastRow_PercentChange)) Then
            WS.Cells(2, 16).Value = WS.Cells(i, 9).Value
            WS.Cells(2, 17).Value = WS.Cells(i, 11).Value
            
            ' statement to find the greatest decrease and print the ticker and value in the correct cells
            ElseIf WS.Cells(i, 11) = WorksheetFunction.Min(WS.Range("K2:K" & LastRow_PercentChange)) Then
            WS.Cells(3, 16).Value = WS.Cells(i, 9).Value
            WS.Cells(3, 17).Value = WS.Cells(i, 11).Value
            
            ' statement to find the greatest stock volume
            ElseIf WS.Cells(i, 12) = WorksheetFunction.Max(WS.Range("L2:L" & LastRow_PercentChange2)) Then
            WS.Cells(4, 16).Value = WS.Cells(i, 9).Value
            WS.Cells(4, 17).Value = WS.Cells(i, 12).Value
            
            End If
            
        Next i
        
            

    
    
    Next WS

    

End Sub

