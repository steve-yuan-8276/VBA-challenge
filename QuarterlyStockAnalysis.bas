Attribute VB_Name = "Module1"
Sub QuarterlyStockAnalysis()
    ' Define variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim dateValue As Date
    Dim openPrice As Double
    Dim closePrice As Double
    Dim volume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double ' Change to Double to handle large volumes
    Dim startRow As Long
    Dim endRow As Long
    Dim resultRow As Long
    Dim i As Long
    Dim j As Long
    Dim dateString As String

    ' Variables to store greatest values
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim tickerGreatestIncrease As String
    Dim tickerGreatestDecrease As String
    Dim tickerGreatestVolume As String

    ' Initialize variables
    greatestIncrease = -999999
    greatestDecrease = 999999
    greatestVolume = 0

    ' Loop through each worksheet (Q1, Q2, Q3, Q4)
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "Q*" Then
            ' Get the last row with data in column A
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

            ' Set result table headers
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Quarterly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"

            resultRow = 2

            ' Loop through all data
            For i = 2 To lastRow
                ticker = ws.Cells(i, 1).Value
                ' Ensure dateValue is properly set
                dateString = ws.Cells(i, 2).Value
                
                If IsDate(dateString) Then
                    dateValue = CDate(dateString)
                Else
                    MsgBox "Date format error in row " & i & " of sheet " & ws.Name & ": " & dateString, vbExclamation
                    Exit Sub
                End If
                
                ' Calculate the quarter
                Dim quarter As String
                quarter = Year(dateValue) & " Q" & Application.WorksheetFunction.RoundUp(Month(dateValue) / 3, 0)

                ' Find the first and last row of each quarter
                startRow = i
                For j = i To lastRow
                    If ws.Cells(j, 1).Value = ticker And (Year(CDate(ws.Cells(j, 2).Value)) & " Q" & Application.WorksheetFunction.RoundUp(Month(CDate(ws.Cells(j, 2).Value)) / 3, 0)) = quarter Then
                        endRow = j
                    Else
                        Exit For
                    End If
                Next j

                ' Calculate the total volume
                openPrice = ws.Cells(startRow, 3).Value
                closePrice = ws.Cells(endRow, 6).Value
                totalVolume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(startRow, 7), ws.Cells(endRow, 7)))

                ' Calculate the quarterly change and percent change
                quarterlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = (quarterlyChange / openPrice)
                Else
                    percentChange = 0
                End If

                ' Output results
                ws.Cells(resultRow, 9).Value = ticker
                ws.Cells(resultRow, 10).Value = quarterlyChange
                ws.Cells(resultRow, 11).Value = percentChange
                ws.Cells(resultRow, 11).NumberFormat = "0.00%" ' Format as percentage
                ws.Cells(resultRow, 12).Value = totalVolume

                ' Check for greatest values
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    tickerGreatestIncrease = ticker
                End If
                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    tickerGreatestDecrease = ticker
                End If
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    tickerGreatestVolume = ticker
                End If

                resultRow = resultRow + 1
                i = endRow
            Next i
            
            ' Apply conditional formatting to "Quarterly Change" column (J)
            Dim rng As Range
            Set rng = ws.Range(ws.Cells(2, 10), ws.Cells(resultRow - 1, 10))
            With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
                .Interior.ColorIndex = 4 ' Green
            End With
            With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
                .Interior.ColorIndex = 3 ' Red
            End With
            
            ' Apply conditional formatting to "Percent Change" column (K)
            Set rng = ws.Range(ws.Cells(2, 11), ws.Cells(resultRow - 1, 11))
            With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
                .Interior.ColorIndex = 4 ' Green
            End With
            With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
                .Interior.ColorIndex = 3 ' Red
            End With

            ' Output greatest values in crosstab format
            ws.Cells(1, 14).Value = "Category"
            ws.Cells(1, 15).Value = "Ticker"
            ws.Cells(1, 16).Value = "Value"

            ws.Cells(2, 14).Value = "Greatest % Increase"
            ws.Cells(2, 15).Value = tickerGreatestIncrease
            ws.Cells(2, 16).Value = greatestIncrease
            ws.Cells(2, 16).NumberFormat = "0.00%"

            ws.Cells(3, 14).Value = "Greatest % Decrease"
            ws.Cells(3, 15).Value = tickerGreatestDecrease
            ws.Cells(3, 16).Value = greatestDecrease
            ws.Cells(3, 16).NumberFormat = "0.00%"

            ws.Cells(4, 14).Value = "Greatest Total Volume"
            ws.Cells(4, 15).Value = tickerGreatestVolume
            ws.Cells(4, 16).Value = greatestVolume
        End If
    Next ws

    MsgBox ("Stock analysis complete.")
End Sub

