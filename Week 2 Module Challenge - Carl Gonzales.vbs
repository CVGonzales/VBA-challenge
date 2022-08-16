Sub stock_stats():

For Each ws In Worksheets

    Dim WorksheetName As String
    Dim i As Long
    Dim j As Long
    Dim Summary_Table_Row As Long
    Dim LastRowA As Long
    Dim LastRowI As Long
    Dim percent_change As Double
    
    WorksheetName = ws.Name
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
Summary_Table_Row = 2
j = 2

LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To LastRowA
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        ws.Cells(Summary_Table_Row, 9).Value = ws.Cells(i, 1).Value
        
        
        ws.Cells(Summary_Table_Row, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
        
        If ws.Cells(Summary_Table_Row, 10).Value >= 0 Then
        
        ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
        
        Else
        
        ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
        
        End If
        
        If ws.Cells(j, 3).Value <> 0 Then
        
        percent_change = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
            
        ws.Cells(Summary_Table_Row, 11) = Format(percent_change, "Percent")
        
        Else
        
        ws.Cells(Summary_Table_Row, 11).Value = Format(0, "Percent")
        
        End If
        
        ws.Cells(Summary_Table_Row, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
        
        Summary_Table_Row = Summary_Table_Row + 1
        
        j = i + 1
        
        End If
        
Next i

'BONUS
LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row

    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
great_increase = ws.Cells(2, 11).Value
great_decrease = ws.Cells(2, 11).Value
great_volume = ws.Cells(2, 12).Value

    For i = 2 To LastRowI
    
      If ws.Cells(i, 11).Value > great_increase Then
      great_increase = ws.Cells(i, 11).Value
      ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
      
      Else
      
      great_increase = great_increase
      ws.Cells(2, 17).Value = Format(great_increase, "Percent")
      
      End If

      If ws.Cells(i, 11).Value < great_decrease Then
      great_decrease = ws.Cells(i, 11).Value
      ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
      
      Else
      
      great_decrease = great_decrease
      ws.Cells(3, 17).Value = Format(great_decrease, "Percent")
      
      End If
      
      If ws.Cells(i, 12).Value > great_volume Then
      great_volume = ws.Cells(i, 12).Value
      ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
      
      Else
      
      great_volume = great_volume
      ws.Cells(4, 17).Value = great_volume
      
      End If
      
Next i

Next ws
        
End Sub