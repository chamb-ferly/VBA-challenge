Attribute VB_Name = "Module3"
Sub VBAchallenge_homework2_chelby()
'A tab

Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate

Dim LastRow As Double
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row - 1
    
'Moderate
'Name New Headers
    WS.Range("I1").Value = "Ticker"
    WS.Range("J1").Value = "Yearly Change"
    WS.Range("K1").Value = "Percent Change"
    WS.Range("L1").Value = "Total Stock Volume"
    
'Name Headers as Dim
    Dim Tickers As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStock As Double
    Dim OpenPrice As Double
'Keep track of the location for each ticker in summary table
    Dim Summary_Row As Integer
    Summary_Row = 2
    
'Define i
    Dim i As Long
    For i = 2 To LastRow
    
'Check if we are within the same ticker, if/then
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    'Tickers
    Tickers = Cells(i, 1).Value
    OpenPrice = Cells(2, 3).Value
    WS.Range("I" & Summary_Row).Value = Tickers
    
    'YearlyChange
    YearlyChange = Cells(i, 6).Value - OpenPrice
    WS.Range("J" & Summary_Row).Value = YearlyChange
    
    'Color Coding
    If Range("J" & Summary_Row).Value > 0 Then
    WS.Range("J" & Summary_Row).Interior.ColorIndex = 4
    Else
    WS.Range("J" & Summary_Row).Interior.ColorIndex = 3
    End If
    
    'PercentChange
    PercentChange = (Cells(i, 6).Value - OpenPrice) / OpenPrice
    WS.Range("K" & Summary_Row).Value = PercentChange
    WS.Range("K1").EntireColumn.NumberFormat = "0.00%"
    
    
    'TotalStock
    TotalStock = TotalStock + Cells(i, 7).Value
    Range("L" & Summary_Row).Value = TotalStock
    
'Add one to the summary table row
    Summary_Row = Summary_Row + 1

'Reset Totals
    YearlyChange = 0
    PercentChange = 0
    TotalStock = 0


'If the cell immediately following a row is the same brand
Else


OpenPrice = Cells(i + 1, 3).Value
YearlyChange = Cells(i, 6) - OpenPrice
PercentChange = (Cells(i, 6).Value - OpenPrice) / OpenPrice
TotalStock = TotalStock + Cells(i, 7).Value


End If
Next i
    
'Challenges

'Create new headers
Range("N2").Value = "Greatest % Increase"
Range("N3").Value = "Greatst % Decrease"
Range("N4").Value = "Greatest Total Volume"
Range("O1").Value = "Ticker"
Range("P1").Value = "Value"

For i = 2 To LastRow
Dim MaxValue As Integer
Dim MinValue As Integer
Dim HighTotal As Integer



Next i
Next WS
End Sub

    
