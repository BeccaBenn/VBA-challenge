Attribute VB_Name = "Module1"

'Public Sub StockChallenge()

'End Sub
Sub StockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    
    For Each ws In Worksheets
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        openingPrice = ws.Cells(2, 3).Value
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        
        For i = 2 To lastRow
            ticker = ws.Cells(i, 1).Value
            closingPrice = ws.Cells(i, 6).Value
            yearlyChange = closingPrice - openingPrice
            percentChange = (yearlyChange / openingPrice) * 100
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' Output information
            ws.Cells(i, 9).Value = ticker
            ws.Cells(i, 10).Value = yearlyChange
            ws.Cells(i, 11).Value = percentChange
            ws.Cells(i, 12).Value = totalVolume
            
            ' Conditional formatting
            If yearlyChange > 0 Then
                ws.Cells(i, 10).Interior.Color = RGB(0, 255, 0) ' Green
            ElseIf yearlyChange < 0 Then
                ws.Cells(i, 10).Interior.Color = RGB(255, 0, 0) ' Red
            End If
            
            ' Check for greatest increase, decrease, and volume
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                greatestIncreaseTicker = ticker
            End If
            
            If percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                greatestDecreaseTicker = ticker
            End If
            
            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                greatestVolumeTicker = ticker
            End If
            
            ' Move to the next ticker
            openingPrice = ws.Cells(i + 1, 3).Value
        Next i
        
        ' Output greatest increase, decrease, and volume
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ws.Cells(2, 16).Value = greatestIncreaseTicker
        ws.Cells(2, 17).Value = greatestIncrease & "%"
        
        ws.Cells(3, 16).Value = greatestDecreaseTicker
        ws.Cells(3, 17).Value = greatestDecrease & "%"
        
        ws.Cells(4, 16).Value = greatestVolumeTicker
        ws.Cells(4, 17).Value = greatestVolume
        
        ' Apply formatting for greatest increase, decrease, and volume
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ' Reset totalVolume for the next worksheet
        totalVolume = 0
    Next ws
End Sub

