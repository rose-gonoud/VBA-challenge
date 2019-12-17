Sub yearlyStockData()

    Dim RowTicker As String
    RowTicker = " "
    Dim CurrentTicker As String
    CurrentTicker = " "
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim TickerVolumeTotal As Double
    TickerVolumeTotal = 0
    Dim TickerYearlyChanges As Double
    TickerYearlyChanges = 0
    Dim TickerPercentChanges As Double
    TickerPercentChanges = 0

    Dim SummaryTableRow As Long
    SummaryTableRow = 2

    numRows = Range("A1", Range("A1").End(xlDown)).Rows.Count

    For i = 2 To numRows

        RowTicker = Cells(i, 1).Value
        TickerVolumeTotal = TickerVolumeTotal + Cells(i, 7).Value

            ' this denotes current stock type
            If CurrentTicker <> RowTicker Then
    
                CurrentTicker = RowTicker
                OpeningPrice = Cells(i, 3).Value
        
            End If

            ' this denotes new stock type
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
                ClosingPrice = Cells(i, 6).Value
                TickerYearlyChanges = (ClosingPrice - OpeningPrice)
                
                    If OpeningPrice <> 0 Then
                        TickerPercentChanges = ((ClosingPrice - OpeningPrice) / OpeningPrice)
                    Else
                        TickerPercentChanges = 0
                    End If
                    
                    If TickerYearlyChanges < 0 Then
                        Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                    Else
                        Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                    End If
            
                Range("I" & SummaryTableRow).Value = RowTicker
                Range("J" & SummaryTableRow).Value = TickerYearlyChanges
                Range("K" & SummaryTableRow).Value = TickerPercentChanges
                Range("L" & SummaryTableRow).Value = TickerVolumeTotal
            
                TickerVolumeTotal = 0
                TickerYearlyChanges = 0
                TickerPercentChanges = 0
            
                SummaryTableRow = SummaryTableRow + 1
        
            End If

    Next i


End Sub
