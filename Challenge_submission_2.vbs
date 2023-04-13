Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data()
    Dim ws As Worksheet, wsOutput As Worksheet
    Dim lastRow As Long, startRow As Long, endRow As Long
    Dim Ticker As String, prevTicker As String
    Dim openPrice As Double, closePrice As Double
    Dim totalVolume As Double, yearlyChange As Double, percentChange As Double
    Dim i As Long, outputRow As Long
    Dim greatestPctIncrease As Double, greatestPctDecrease As Double, greatestTotalVolume As Double
    Dim tickerGreatestPctIncrease As String, tickerGreatestPctDecrease As String, tickerGreatestTotalVolume As String

    For Each ws In ThisWorkbook.Worksheets
        ' Set the output sheet
        Set wsOutput = ws

        ' Find the last row in the worksheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        ' Set variables
        prevTicker = ws.Cells(2, 1).Value
        startRow = 2
        outputRow = 2
        greatestPctIncrease = 0
        greatestPctDecrease = 0
        greatestTotalVolume = 0

        ' Add headers
        wsOutput.Cells(1, 9).Value = "Ticker"
        wsOutput.Cells(1, 10).Value = "Yearly Change"
        wsOutput.Cells(1, 11).Value = "Percentage Change"
        wsOutput.Cells(1, 12).Value = "Total Volume"

        ' Add headers for greatest values
        wsOutput.Cells(2, 14).Value = "Greatest % Increase"
        wsOutput.Cells(3, 14).Value = "Greatest % Decrease"
        wsOutput.Cells(4, 14).Value = "Greatest Total Volume"

        ' Loop through all the rows
        For i = 2 To lastRow + 1
            Ticker = ws.Cells(i, 1).Value

            ' Check if the stock symbol has changed or reached the end
            If Ticker <> prevTicker Or i = lastRow + 1 Then
                endRow = i - 1

                ' Calculate the yearly change
                openPrice = ws.Cells(startRow, 3).Value
                closePrice = ws.Cells(endRow, 4).Value
                yearlyChange = closePrice - openPrice

                ' Calculate the percentage change
                If openPrice <> 0 Then
                    percentChange = (yearlyChange) / openPrice * 100
                Else
                    percentChange = 0
                End If

                ' Calculate the total volume
                totalVolume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(startRow, 5), ws.Cells(endRow, 5)))

                ' Update the greatest percentage increase and decrease, and greatest total volume
                If totalVolume > greatestTotalVolume Then
                    greatestTotalVolume = totalVolume
                    tickerGreatestTotalVolume = prevTicker
                End If

                ' Output the results to the output sheet
                wsOutput.Cells(outputRow, 9).Value = prevTicker
                wsOutput.Cells(outputRow, 10).Value = yearlyChange
                wsOutput.Cells(outputRow, 11).Value = Format(percentChange) & "%"
                wsOutput.Cells(outputRow, 12).Value = totalVolume

                ' Update outputRow for the next stock
                outputRow = outputRow + 1

                ' Update variables for the next stock
                prevTicker = Ticker
                startRow = i
            End If
            
            ' Output results
                wsOutput.Cells(outputRow, 9).Value = prevTicker
                wsOutput.Cells(outputRow, 10).Value = yearlyChange

                'Cell color based on the yearly change
                If yearlyChange < 0 Then
                    wsOutput.Cells(outputRow, 10).Interior.Color = RGB(255, 0, 0) ' Red
                Else
                    wsOutput.Cells(outputRow, 10).Interior.Color = RGB(0, 255, 0) ' Green
                End If

                wsOutput.Cells(outputRow, 11).Value = Format(percentChange) & "%"
                wsOutput.Cells(outputRow, 12).Value = totalVolume
        Next i

        ' Output the greatest percent increase, decrease and total volume
        wsOutput.Cells(2, 15).Value = tickerGreatestPctIncrease & " (" & Format(greatestPctIncrease, "0.00") & "%)"
        wsOutput.Cells(3, 15).Value = tickerGreatestPctDecrease & " (" & Format(greatestPctDecrease, "0.00") & "%)"
        wsOutput.Cells(4, 15).Value = tickerGreatestTotalVolume & " (" & Format(greatestTotalVolume, "#,##0") & ")"

    Next ws
End Sub


