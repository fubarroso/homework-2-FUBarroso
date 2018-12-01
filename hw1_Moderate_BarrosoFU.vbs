Sub Moderate()
    
    'Set variables
    Dim SummaryRow As Integer
    Dim CurrentStock As String
    Dim OpenStockValue As Double
    Dim CloseStockValue As Double
    Dim StockVolume As Double
    Dim YearlyChange As Double
    
    SummaryRow = 2
    StockVolume = 0
    
    'Print column titles and adjust width
    Cells(1, 9) = "Ticker"
    Columns("I:I").ColumnWidth = 12
    Cells(1, 10) = "Yearly Change"
    Columns("J:J").ColumnWidth = 16
    Cells(1, 11) = "Percentage Change"
    Columns("K:K").ColumnWidth = 16
    Cells(1, 12) = "Total Stock Volume"
    Columns("L:L").ColumnWidth = 16
    
    'find total number of rows
    Dim LastRow As Long
    With ActiveSheet
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    
    'First run print Stok IDs and calculate and print total stock Volume
    For i = 2 To LastRow

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            CurrentStock = Cells(i, 1).Value
            StockVolume = StockVolume + Cells(i, 7).Value
            Cells(SummaryRow, 9) = CurrentStock
            Cells(SummaryRow, 12) = StockVolume
            SummaryRow = SummaryRow + 1
            StockVolume = 0
        
        Else
            StockVolume = StockVolume + Cells(i, 7).Value
        End If
    Next i
    
    'reset print row
    SummaryRow = 2
    
    'Second run to calculate and print Yearly change and percentage change
    For i = 2 To LastRow
        
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            OpenStockValue = Cells(i, 3).Value
           
        Else
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                CloseStockValue = Cells(i, 6).Value
                              
                YearlyChange = CloseStockValue - OpenStockValue
                Cells(SummaryRow, 10) = YearlyChange
                'Print yearly change with conditianal for positive or negative
                If (CloseStockValue - OpenStockValue) >= 0 Then
                    Cells(SummaryRow, 10).Select
                    Selection.Style = "Good"
               
                Else
                    Cells(SummaryRow, 10).Select
                    Selection.Style = "Bad"
                
                End If
                'Calculate and print percentage change (protecting against no value for missing opening value
                If OpenStockValue <> 0 Then
                Cells(SummaryRow, 11) = YearlyChange / OpenStockValue
                Else
                    Cells(SummaryRow, 11) = "N/A"
                End If
                SummaryRow = SummaryRow + 1
            End If
        End If
        
    Next i
    

End Sub
 

 
