Attribute VB_Name = "Module1"
Sub StockTickerLoop():

    For Each ws In Worksheets

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        Dim TickerName As String
        Dim TotalTickerVolume As Double
        TotalTickerVolume = 0
        Dim Opening As Double
        Dim Closing As Double
        Dim Change As Double
        
        Dim PreviousAmount As Long
        PreviousAmount = 2
        Dim PercentChange As Double
        Dim GreatestIncrease As Double
        GreatestIncrease = 0
        Dim GreatestDecrease As Double
        GreatestDecrease = 0
        Dim GreatestTotalVolume As Double
        GreatestTotalVolume = 0
                
        Dim SummaryTableRow As Long
        SummaryTableRow = 2
        
        Dim LastRow As Long
        Dim LastRowValue As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow

            TotalTickerVolume = TotalTickerVolume + ws.Cells(i, 7).Value
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                TickerName = ws.Cells(i, 1).Value
                    ws.Range("I" & SummaryTableRow).Value = TickerName
                    ws.Range("L" & SummaryTableRow).Value = TotalTickerVolume
                TotalTickerVolume = 0

                Opening = ws.Range("C" & PreviousAmount)
                Closing = ws.Range("F" & i)
                Change = Closing - Opening
                    ws.Range("J" & SummaryTableRow).Value = Change

                    If Opening = 0 Then
                        PercentChange = 0
                    Else
                        Opening = ws.Range("C" & PreviousAmount)
                        PercentChange = Change / Opening
                    End If
                        ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                        ws.Range("K" & SummaryTableRow).Value = PercentChange

                    If ws.Range("J" & SummaryTableRow).Value >= 0 Then
                        ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                    Else
                        ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                    End If

                    SummaryTableRow = SummaryTableRow + 1
                    PreviousAmount = i + 1
                    End If
                    
                Next i

            LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
                For i = 2 To LastRow
                    If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                        ws.Range("Q2").Value = ws.Range("K" & i).Value
                        ws.Range("P2").Value = ws.Range("I" & i).Value
                    End If

                    If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                        ws.Range("Q3").Value = ws.Range("K" & i).Value
                        ws.Range("P3").Value = ws.Range("I" & i).Value
                    End If

                    If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                        ws.Range("Q4").Value = ws.Range("L" & i).Value
                        ws.Range("P4").Value = ws.Range("I" & i).Value
                    End If

                Next i
            
                    ws.Range("Q2").NumberFormat = "0.00%"
                    ws.Range("Q3").NumberFormat = "0.00%"
            
                ws.Columns("I:Q").AutoFit

        Next ws

End Sub
