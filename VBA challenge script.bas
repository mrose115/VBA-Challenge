Attribute VB_Name = "Module2"
Sub StockData2()

For Each ws In Worksheets

Dim WorksheetName As String

WorksheetName = ws.Name

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

Dim Ticker As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim StockVolume As Double
StockVolume = 0

Dim SummaryTable As Integer
SummaryTable = 2

Dim StockOpen As Double
Dim StockClose As Double
StockOpen = 0
StockClose = 0

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow

  If StockOpen = 0 Then
        StockOpen = ws.Cells(i, 3).Value
        
End If

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        ws.Range("I" & SummaryTable).Value = Ticker
        
        StockClose = ws.Cells(i, 6).Value
        
        YearlyChange = StockClose - StockOpen
        ws.Range("J" & SummaryTable).Value = YearlyChange
            
            If StockClose = 0 Then
            ws.Range("K" & SummaryTable).Value = Null
            
            Else: PercentChange = (StockClose - StockOpen) / StockOpen
            ws.Range("K" & SummaryTable).Value = PercentChange
            
            End If
        
        StockVolume = StockVolume + ws.Cells(i, 7).Value
        ws.Range("L" & SummaryTable).Value = StockVolume
    
    SummaryTable = SummaryTable + 1
    
    StockVolume = 0
    
    StockOpen = 0
    
    StockClose = 0
    
Else
    StockVolume = StockVolume + ws.Cells(i, 7).Value

End If

Next i

For i = 2 To LastRow

ws.Cells(i, 11).NumberFormat = "0.00%"

If ws.Cells(i, 10).Value > 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 4

Else
ws.Cells(i, 10).Interior.ColorIndex = 3

End If

Next i

Next ws



End Sub

