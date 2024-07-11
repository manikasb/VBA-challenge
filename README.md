Sub multiplestocks()


For Each ws In Worksheets
    
        Dim WorksheetName As String
        Dim i As Long
        Dim j As Long
        Dim TickCount As Long
        Dim LastRowA As Long
        Dim LastRowI As Long
        Dim PerChange As Double
        Dim GreatIncr As Double
        Dim GreatDecr As Double
        Dim GreatVol As Double
        
        'Get the WorksheetName
        WorksheetName = ws.Name
        
        'Create column headers
        ws.Cells(1, 9).value = "Ticker"
        ws.Cells(1, 10).value = "Yearly Change"
        ws.Cells(1, 11).value = "Percent Change"
        ws.Cells(1, 12).value = "Total Stock Volume"
        ws.Cells(1, 16).value = "Ticker"
        ws.Cells(1, 17).value = "Value"
        ws.Cells(2, 15).value = "Greatest % Increase"
        ws.Cells(3, 15).value = "Greatest % Decrease"
        ws.Cells(4, 15).value = "Greatest Total Volume"
        
        'Set Ticker Counter to first row
        TickCount = 2
        
        'Set start row to 2
        j = 2
        
        'Find the last non-blank cell in column A
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            'Loop through all rows
            For i = 2 To LastRowA
            
                'Check if ticker name changed
                If ws.Cells(i + 1, 1).value <> ws.Cells(i, 1).value Then
                
                'Write ticker in column I (#9)
                ws.Cells(TickCount, 9).value = ws.Cells(i, 1).value
                
                'Calculate and write Yearly Change in column J (#10)
                ws.Cells(TickCount, 10).value = ws.Cells(i, 6).value - ws.Cells(j, 3).value
                
                    'Conditional formating
                    If ws.Cells(TickCount, 10).value < 0 Then
                    
                    'Set cell background color to red
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 3
                
                    Else
                
                    'Set cell background color to green
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    'Calculate and write percent change in column K (#11)
                    If ws.Cells(j, 3).value <> 0 Then
                    PerChange = ((ws.Cells(i, 6).value - ws.Cells(j, 3).value) / ws.Cells(j, 3).value)
                    
                    'Percent formating
                    ws.Cells(TickCount, 11).value = Format(PerChange, "Percent")
        
        Else
                    
                    ws.Cells(TickCount, 11).value = Format(0, "Percent")
                    
                    End If
                    
                'Calculate and write total volume in column L (#12)
                ws.Cells(TickCount, 12).value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                'Increase TickCount by 1
                TickCount = TickCount + 1
                
                'Set new start row of the ticker block
                j = i + 1
                
                End If
                
     Next i
            
        'Find last non-blank cell in column I
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        'MsgBox ("Last row in column I is " & LastRowI)
        
        'Prepare for summary
        GreatVol = ws.Cells(2, 12).value
        GreatIncr = ws.Cells(2, 11).value
        GreatDecr = ws.Cells(2, 11).value
        
            'Loop for summary
            For i = 2 To LastRowI
            
                'For greatest total volume--check if next value is larger--if yes take over a new value and populate ws.Cells
                If ws.Cells(i, 12).value > GreatVol Then
                GreatVol = ws.Cells(i, 12).value
                ws.Cells(4, 16).value = ws.Cells(i, 9).value
                
                Else
                
                GreatVol = GreatVol
                
                End If
                
                'For greatest increase--check if next value is larger--if yes take over a new value and populate ws.Cells
                If ws.Cells(i, 11).value > GreatIncr Then
                GreatIncr = ws.Cells(i, 11).value
                ws.Cells(2, 16).value = ws.Cells(i, 9).value
                
                Else
                
                GreatIncr = GreatIncr
                
                End If
                
            'Write summary results in ws.Cells
            ws.Cells(2, 17).value = Format(GreatIncr, "Percent")
            ws.Cells(3, 17).value = Format(GreatDecr, "Percent")
            ws.Cells(4, 17).value = Format(GreatVol, "Scientific")
            
            Next i
            
        'Adjust column width automatically
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
            
    Next ws
        

End Sub

