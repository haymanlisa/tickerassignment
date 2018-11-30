Sub LoopThroughWorksheets()

    Dim WorkSheet_Count As Integer
    Dim I As Integer

    ' Number of worksheets in the project workbook
    WorkSheet_Count = ActiveWorkbook.Worksheets.Count

    ' Iterate through the worksheets
    For I = 1 To WorkSheet_Count

       ' Activate worksheet and Analyze stock data
       ActiveWorkbook.Worksheets(I).Activate
       Call stock_market

    Next I
End Sub


Sub stock_market()
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    item = 2
    totalsum = 0
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Total Volume"
    
    For I = 2 To LastRow
    
        totalsum = totalsum + Cells(I, 7).Value
    
    
        If Cells(I, 1).Value <> Cells(I + 1, 1).Value Then
        
            Cells(item, 9).Value = Cells(I, 1)
            
            Cells(item, 10).Value = totalsum
            
            item = item + 1
            
            totalsum = 0
            
        End If
        
    Next I
  
    
End Sub


