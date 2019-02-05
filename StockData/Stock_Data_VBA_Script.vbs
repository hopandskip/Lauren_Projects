Sub StockData()

For Each ws In Worksheets

'Get the LastRow of Data
Dim LastRow As Long
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Sort Data by Ticker
ws.Range("A2:G" & LastRow).Sort Key1:=ws.Range("A1"), Order1:=xlAscending

'Add Headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'Get the last row for results dataset
Dim LastRowResults As Long
LastRowResults = ws.Cells(Rows.Count, 9).End(xlUp).Row

'Get the distinct list of tickers and put in Ticker Column
For i = 2 To LastRow
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        LastRowResults = LastRowResults + 1
        ws.Cells(LastRowResults, 9).Value = ws.Cells(i, 1).Value
    End If
Next i

'Get the Yearly Change
Dim EarliestDate As Long
Dim EndDate As Long
Dim EndYearStockPrice As Double
Dim BegyearStockprice As Double
Dim TotalVolume As Double

'Create variable to assist with the for loop
LastMatch = 2

For i = 2 To LastRowResults
    TotalVolume = 0
    For j = LastMatch To LastRow
        If ws.Cells(j, 1).Value <> ws.Cells(i, 9).Value Then
        Exit For
        ElseIf ws.Cells(j, 1).Value = ws.Cells(i, 9).Value Then
            'To Get A Date for Comparison
            If ws.Cells(j - 1, 1).Value <> ws.Cells(j, 1).Value Then
                EarliestDate = ws.Cells(j, 2).Value
                EndDate = ws.Cells(j, 2).Value
                EndYearStockPrice = ws.Cells(j, 6).Value
                BegyearStockprice = ws.Cells(j, 3).Value
            End If
            If ws.Cells(j, 2).Value < EarliestDate Then
                EarliestDate = ws.Cells(j, 2).Value
                BegyearStockprice = ws.Cells(j, 3).Value
            End If
            If ws.Cells(j, 2).Value > EndDate Then
                EndDate = ws.Cells(j, 2).Value
                EndYearStockPrice = ws.Cells(j, 6).Value
            End If
            TotalVolume = TotalVolume + ws.Cells(j, 7).Value
            LastMatch = j + 1
        End If
    Next j
    ws.Cells(i, 12).Value = TotalVolume
    ws.Cells(i, 10).Value = (EndYearStockPrice - BegyearStockprice)
    
    If (BegyearStockprice = 0) Or (ws.Cells(i, 10).Value = 0) Then
        ws.Cells(i, 11).Value = 0
    Else
        ws.Cells(i, 10).Value = (EndYearStockPrice - BegyearStockprice)
        ws.Cells(i, 11).Value = (EndYearStockPrice - BegyearStockprice) / BegyearStockprice
    End If
    
    If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
    ElseIf ws.Cells(i, 10).Value < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 3
    End If
Next i

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

Dim GreatestDecrease As Double
Dim GreatestIncrease As Double
Dim GreatestVolume As Double
Dim GreatestIncreaseTicker As String
Dim GreatestDecreaseTicker As String
Dim GreatestVolumeTicker As String

GreatestDecrease = 0
GreatestIncrease = 0
GreatestVolume = 0
For i = 3 To LastRowResults
    If ws.Cells(i, 11).Value < GreatestDecrease Then
        GreatestDecrease = ws.Cells(i, 11).Value
        GreatestDecreaseTicker = ws.Cells(i, 9).Value
    ElseIf ws.Cells(i, 11).Value > GreatestIncrease Then
        GreatestIncrease = ws.Cells(i, 11).Value
        GreatestIncreaseTicker = ws.Cells(i, 9).Value
    End If
    If ws.Cells(i, 12).Value > GreatestVolume Then
        GreatestVolume = ws.Cells(i, 12).Value
        GreatestVolumeTicker = ws.Cells(i, 9).Value
    End If
Next i

ws.Range("P2").Value = GreatestIncreaseTicker
ws.Range("P3").Value = GreatestDecreaseTicker
ws.Range("Q2").Value = GreatestIncrease
ws.Range("Q3").Value = GreatestDecrease
ws.Range("P4").Value = GreatestVolumeTicker
ws.Range("Q4").Value = GreatestVolume

ws.Columns("A:Q").AutoFit
ws.Columns("K:K").NumberFormat = "0.00%"
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").NumberFormat = "0.00%"

Next ws

End Sub
