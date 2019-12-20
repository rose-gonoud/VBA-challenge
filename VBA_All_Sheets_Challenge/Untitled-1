Sub Challenge1_summaryData()

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Changes"
    Cells(1, 11).Value = "Percent Changes"
    Cells(1, 12).Value = "Total Stock Volume"
    
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"

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
    Dim greatestPercInc As Double
    greatestPercInc = 0
    Dim greatestPercDec As Double
    greatestPercDec = 0
   Dim greatestTickerVolume As Double
   greatestTickerVolume = 0

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
            
                    If TickerPercentChanges > greatestPercInc Then
                        greatestPercInc = TickerPercentChanges
                        Cells(2, 15).Value = RowTicker
                        Cells(2, 16).Value = greatestPercInc
                    End If
                    
                     If TickerPercentChanges < greatestPercDec Then
                        greatestPercDec = TickerPercentChanges
                        Cells(3, 15).Value = RowTicker
                        Cells(3, 16).Value = greatestPercDec
                    End If
                    
                     If TickerVolumeTotal > greatestTickerVolume Then
                        greatestTickerVolume = TickerVolumeTotal
                        Cells(4, 15).Value = RowTicker
                        Cells(4, 16).Value = greatestTickerVolume
                    End If
            
                Range("I" & SummaryTableRow).Value = RowTicker
                Range("J" & SummaryTableRow).Value = TickerYearlyChanges
                Range("K" & SummaryTableRow).Value = TickerPercentChanges
                Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                Range("L" & SummaryTableRow).Value = TickerVolumeTotal
                
                Cells(2, 16).NumberFormat = "0.00%"
                Cells(3, 16).NumberFormat = "0.00%"
            
                TickerVolumeTotal = 0
                TickerYearlyChanges = 0
                TickerPercentChanges = 0
            
                SummaryTableRow = SummaryTableRow + 1
        
            End If

    Next i
     
Cells().EntireColumn.AutoFit
    
End Sub

Sub Challenge2_allSheets()

Dim ws_Count As Integer
Dim i As Integer

         ws_Count = ActiveWorkbook.Worksheets.Count

         For i = 1 To ws_Count
         
            Worksheets(i).Select
            Challenge1_summaryData

         Next i

End Sub
