Attribute VB_Name = "Module1"
Sub StockTicker()
    Dim ws As Worksheet
    For Each ws In Worksheets
        Dim WorksheetName As String
        WorksheetName = ws.Name



ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"


ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest Decreae"
ws.Range("O4").Value = "Greatest Total Volume"


Dim Ticker As Long
Ticker = 2
Dim Start As Long
Start = 2

Dim LastRowA As Long
LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim yearlyChange As Double
    Dim percentChange As Double
Dim i As Long
Dim totalVolume As Double
For i = 2 To LastRowA
    totalVolume = totalVolume + ws.Cells(i, 7).Value
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ws.Cells(Ticker, 9).Value = ws.Cells(i, 1).Value
    
  
        yearlyChange = ws.Cells(i, 6) - ws.Cells(Start, 3)
        ws.Cells(Ticker, 10).Value = yearlyChange
    

        If ws.Cells(Ticker, 3).Value = 0 Then
            percentChange = 0
        Else
            percentChange = yearlyChange / ws.Cells(Start, 3).Value
        End If
        ws.Cells(Ticker, 11).Value = Format(percentChange, "0.00%")
    
        ws.Cells(Ticker, 12).Value = totalVolume

        totalVolume = 0

        Start = i + 1
        If yearlyChange < 0 Then
        ws.Cells(Ticker, 10).Interior.ColorIndex = 3
        ElseIf yearlyChange > 0 Then
        ws.Cells(Ticker, 10).Interior.ColorIndex = 4
        End If
        Ticker = Ticker + 1
    End If
Next i


 MinValue = WorksheetFunction.Min(ws.Range("K2:K" & LastRowA))
 ws.Range("Q3") = MinValue * 100
 Ticker_Pos = WorksheetFunction.Match(MinValue, ws.Range("K2:K" & LastRowA), 0)
 ws.Range("P3") = ws.Cells(Ticker_Pos, 9)
 
 MaxValue = WorksheetFunction.Max(ws.Range("K2:K" & LastRowA))
 ws.Range("Q2") = MaxValue * 100
 Ticker_Pos = WorksheetFunction.Match(MaxValue, ws.Range("K2:K" & LastRowA), 0)
 ws.Range("P2") = ws.Cells(Ticker_Pos, 9)
 
 MaxVolume = WorksheetFunction.Max(ws.Range("L2:L" & LastRowA))
 ws.Range("Q4") = MaxVolume
 Ticker_Pos = WorksheetFunction.Match(MaxVolume, ws.Range("L2:L" & LastRowA), 0)
 ws.Range("P4") = ws.Cells(Ticker_Pos, 9)
 
 Next ws
 
 
 End Sub
