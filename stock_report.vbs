Sub stock_report()
    Dim rowcount As Double
    Dim totalVolume As Double
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim outputRow As Integer
    Dim delta As Double
    Dim deltaPct As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim largestVolume As Double
    Dim greatestIncreasStock As String
    Dim greatestDecreasStock As String
    Dim largestVolumStock As String

    For Each ws In Worksheets
        ' initialize variables
        greatestIncrease = 0
        greatestDecrease = 0
        largestVolume = 0
        rowcount = ws.Cells(Rows.Count, 1).End(xlUp).Row
        totalVolume = 0
        outputRow = 2


        ' output header row and labels
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"

        For i = 2 To rowcount
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            If ws.Cells(i, 1) <> ws.Cells(i - 1, 1) Then
                ' row i is the first line of the ticker symbol
                openingPrice = ws.Range("C" & i).Value
            End If

            If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
                ' row i is the last line of the ticker symbol
                closingPrice = ws.Range("F" & i).Value

                ' output stock information
                delta = closingPrice - openingPrice
                ws.Range("I" & outputRow).Value = ws.Cells(i, 1)
                ws.Range("J" & outputRow).Value = delta
                If (delta < 0) Then
                    ws.Range("J" & outputRow).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & outputRow).Interior.ColorIndex = 4
                End If
                deltaPct = delta / openingPrice
                ws.Range("K" & outputRow).Value = deltaPct
                ws.Range("L" & outputRow).Value = totalVolume

                ' update variables
                outputRow = outputRow + 1
                If totalVolume > largestVolume Then
                    largestVolume = totalVolume
                    largestVolumeStock = ws.Cells(i, 1)
                End If
                totalVolume = 0
                If deltaPct < greatestDecrease Then
                    greatestDecrease = deltaPct
                    greatestDecreaseStock = ws.Cells(i, 1)
                End If
                If deltaPct > greatestIncrease Then
                    greatestIncrease = deltaPct
                    greatestIncreaseStock = ws.Cells(i, 1)
                End If
            End If
        ' go on to the next row in this sheet
        Next i

        'Output summary information
        ws.Range("O2").Value = greatestIncreaseStock
        ws.Range("P2").Value = greatestIncrease
        ws.Range("O3").Value = greatestDecreaseStock
        ws.Range("P3").Value = greatestDecrease
        ws.Range("O4").Value = largestVolumeStock
        ws.Range("P4").Value = largestVolume

        ' Cleanup formatting
        ws.Range("K2:K" & rowcount).NumberFormat = "0.00%"
        ws.Range("P2").NumberFormat = "0.00%"
        ws.Range("P3").NumberFormat = "0.00%"
        ws.Columns("A:L").AutoFit
        ws.Columns("N:P").AutoFit

        ' go on to the next worksheet
    Next
End Sub
