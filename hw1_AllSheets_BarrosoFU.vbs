Sub Dosomething()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call Hard
    Next
    Application.ScreenUpdating = True
End Sub

Sub Hard()
    
    'Set variables
    Dim SummaryRow As Integer
    Dim CurrentStock As String
    Dim OpenStockValue As Double
    Dim CloseStockValue As Double
    Dim PercentageMax As Double
    Dim PercentageMin As Double
    Dim CurrentStockPercentage As Variant
    Dim CurrentStockVolume As Double
    Dim MaxStockVolume As Double
    Dim YearlyChange As Double
    Dim MaxVol As Double
    
    SummaryRow = 2
    StockVolume = 0
    
    'Print column titles and format columns as needed
    Cells(1, 9) = "Ticker"
    Columns("I:I").ColumnWidth = 12
    Cells(1, 10) = "Yearly Change"
    Columns("J:J").ColumnWidth = 16
    Cells(1, 11) = "Percentage Change"
    Columns("K:K").ColumnWidth = 16
    Cells(1, 12) = "Total Stock Volume"
    Columns("L:L").ColumnWidth = 16
    Columns("K:K").Select
    Selection.NumberFormat = "0.00%"
    Columns("J:J").Select
    Selection.NumberFormat = "0.000000000"
    Columns("O:O").ColumnWidth = 20
    Cells(2, 15) = "Greatest % Increase"
    Cells(3, 15) = "Greatest % Decrease"
    Cells(4, 15) = "Greatest Total Volume"
    Cells(1, 16) = "Ticker"
    Cells(1, 17) = "Value"
    Range("Q2:Q3").Select
    Selection.NumberFormat = "0.00%"
    Columns("Q:Q").ColumnWidth = 20
    
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
    
    'count rows in first output
    With ActiveSheet
        LastRow = .Cells(.Rows.Count, "I").End(xlUp).Row
    End With
    
   'reset print row
    SummaryRow = 2
    PercentageMax = 0
    PercentageMin = 0
    MaxStockVolume = 0
    
    'Find Max/Min % change and Greatest Total Volume
    
    For i = 2 To LastRow
    
        CurrentStock = Cells(i, 9).Value
        If Cells(i, 11).Value <> "N/A" Then
        CurrentStockPercentage = Cells(i, 11).Value
        Else
        CurrentStockPercentage = 0
        End If
        CurrentStockVolume = Cells(i, 12).Value
        
        If CurrentStockPercentage > PercentageMax Then
            PercentageMax = CurrentStockPercentage
            Cells(2, 16).Value = CurrentStock
            Cells(2, 17).Value = PercentageMax
        Else
        End If
        
         If CurrentStockPercentage < PercentageMin Then
            PercentageMin = CurrentStockPercentage
            Cells(3, 16).Value = CurrentStock
            Cells(3, 17).Value = PercentageMin
        Else
        End If
        
        If CurrentStockVolume > MaxStockVolume Then
            MaxStockVolume = CurrentStockVolume
            Cells(4, 16).Value = CurrentStock
            Cells(4, 17).Value = CurrentStockVolume
        Else
        End If
        
    Next i

End Sub




