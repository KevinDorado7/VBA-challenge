Attribute VB_Name = "Module1"
Sub StockData_AllSheets()
    
    Dim ws As Worksheet
    
        For Each ws In Worksheets
        
        StockData ws
        
    Next ws
End Sub

Sub StockData(ws As Worksheet)
    
    Dim lastRow As Long
    
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
    
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim volumeTotal As Double
    
    Dim startRow As Long
        startRow = 2

    Dim i As Long
    Dim outputRow As Long
        outputRow = 2
    
    Dim greatestIncrease As Double
        greatestIncrease = 0
    
    Dim greatestDecrease As Double
        greatestDecrease = 0
    
    Dim greatestVolume As Double
        greatestVolume = 0
    
    Dim tickerGreatestIncrease As String
    Dim tickerGreatestDecrease As String
    Dim tickerGreatestVolume As String
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Yearly Value"

    For i = 2 To lastRow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Or i = lastRow Then
            
            ticker = ws.Cells(i, 1).Value
                
                openingPrice = ws.Cells(startRow, 3).Value
                    closingPrice = ws.Cells(i, 6).Value
                volumeTotal = volumeTotal + ws.Cells(i, 7).Value
            
                yearlyChange = closingPrice - openingPrice

            If openingPrice <> 0 Then
            
                percentChange = (yearlyChange / openingPrice)
            
            Else
            
                percentChange = 0
            End If

            If percentChange > greatestIncrease Then
                
                greatestIncrease = percentChange
                
                tickerGreatestIncrease = ticker
            End If
            
            If percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                tickerGreatestDecrease = ticker
            End If
            
            If volumeTotal > greatestVolume Then
                greatestVolume = volumeTotal
                tickerGreatestVolume = ticker
            
            End If

            ws.Cells(outputRow, 9).Value = ticker
            ws.Cells(outputRow, 10).Value = yearlyChange
            ws.Cells(outputRow, 11).Value = percentChange
            ws.Cells(outputRow, 12).Value = volumeTotal

            volumeTotal = 0
            startRow = i + 1
            outputRow = outputRow + 1
        Else
            volumeTotal = volumeTotal + ws.Cells(i, 7).Value
        
        End If
    Next i

    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"

    ws.Cells(2, 16).Value = tickerGreatestIncrease
    ws.Cells(3, 16).Value = tickerGreatestDecrease
    ws.Cells(4, 16).Value = tickerGreatestVolume

    ws.Cells(2, 17).Value = greatestIncrease
    ws.Cells(3, 17).Value = greatestDecrease
    ws.Cells(4, 17).Value = greatestVolume

    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Columns("J:K").NumberFormat = "0.00%"
    ws.Columns("L").NumberFormat = "0"
    ws.Columns("K:K").NumberFormat = "0.00%"
    
    Dim lastRow2 As Long
    lastRow2 = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    
    For j = 2 To lastRow2
        If ws.Cells(j, 10).Value < 0 Then
        
            ws.Cells(j, 10).Interior.ColorIndex = 3
            
        ElseIf ws.Cells(j, 10).Value > 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 4
            
        End If
        
    Next j
End Sub

