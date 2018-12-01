Sub Easy()
    
    Dim SummaryRow As Integer
    SummaryRow = 2
    Dim CurrentStock As String
    Dim StockVolume As Double
    StockVolume = 0
    
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Ticker"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Total Stock Volume"
    Range("J2").Select
    Columns("J:J").ColumnWidth = 15.83
    
    
    Dim LastRow As Long
    With ActiveSheet
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    
    For i = 2 To LastRow

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            CurrentStock = Cells(i, 1).Value
            StockVolume = StockVolume + Cells(i, 7).Value
            Cells(SummaryRow, 9) = CurrentStock
            Cells(SummaryRow, 10) = StockVolume
            SummaryRow = SummaryRow + 1
            StockVolume = 0
        
        Else
            StockVolume = StockVolume + Cells(i, 7).Value
        End If
    Next i

End Sub
••••ˇˇˇˇ