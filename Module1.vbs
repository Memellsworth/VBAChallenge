Attribute VB_Name = "Module1"
Sub vbachallenge()
For Each ws In Worksheets

Dim worksheetName As String

Dim ticker As String
Dim yearlychange As Double
Dim startingprice As Double
startingprice = ws.Cells(2, 3).Value
Dim endprice As Double
Dim percentchange As Double
Dim tickcount As Double
tickcount = 2
Dim Greatestincrease As Double
Greatestincrease = ws.Cells(2, 11).Value
Dim Greatestdecrease As Double
Greatestdecrease = ws.Cells(2, 11).Value
Dim Greatestvolume As Double
Greatestvolume = ws.Cells(2, 12).Value



worksheetName = ws.Name

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Grand Total Volume"
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
For i = 2 To lastrow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ws.Cells(tickcount, 9).Value = ws.Cells(i, 1).Value
        endprice = ws.Cells(i, 6).Value
        yearlychange = endprice - startingprice
        ws.Cells(tickcount, 10).Value = yearlychange
        percentchange = ((endprice - startingprice) / startingprice)
        ws.Cells(tickcount, 11).Value = Format(percentchange, "Percent")
        totalstock = ws.Cells(i, 7).Value + totalstock
        ws.Cells(tickcount, 12).Value = totalstock
        startingprice = ws.Cells(i + 1, 3).Value
        tickcount = tickcount + 1
        totalstock = 0
        
        Else
        totalstock = ws.Cells(i, 7).Value + totalstock

        End If
    Next i
      lastrowb = ws.Cells(Rows.Count, 9).End(xlUp).Row
      
For i = 2 To lastrowb

        If ws.Cells(i, 10).Value < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 3
        
        Else
        ws.Cells(i, 10).Interior.ColorIndex = 4
        
        
    
    End If
    
    
Next i


For i = 2 To lastrowb
 If ws.Cells(i, 11).Value > Greatestincrease Then
    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
    Greatestincrease = ws.Cells(i, 11).Value
    ws.Cells(2, 17).Value = Format(Greatestincrease, "Percent")
    
    Else
    Greatestincrease = Greatestincrease
    
    End If
    
    
If ws.Cells(i, 11).Value < Greatestdecrease Then
    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    Greatestdecrease = ws.Cells(i, 11).Value
    ws.Cells(3, 17).Value = Format(Greatestdecrease, "Percent")
    
    Else
    Greatestdecrease = Greatestdecrease
    
    End If
    
If ws.Cells(i, 12).Value > Greatestvolume Then
    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    Greatestvolume = ws.Cells(i, 12).Value
    ws.Cells(4, 17).Value = Greatestvolume
    
    Else
    
    Greatestvolume = Greatestvolume
    
    End If
Next i
Worksheets(worksheetName).Columns("A:Z").AutoFit

Next ws
End Sub
