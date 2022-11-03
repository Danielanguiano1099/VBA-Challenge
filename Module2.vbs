Sub Stocks()
    Dim Total As Double
    Dim Blank As Long
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Price_Change As Double
    Dim PCT_Change As Double
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Total = 0
    Blank = 2
    
    
    
    For i = 1 To 759001
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Open_Price = Cells(i + 1, 3).Value
        ElseIf Cells(i + 1, 1).Value <> Cells(i + 2, 1).Value Then
            Close_Price = Cells(i + 1, 6).Value
            Price_Change = Close_Price - Open_Price
            Cells(Blank, 10).Value = Price_Change
            PCT_Change = Price_Change / Open_Price
            Cells(Blank, 11).Value = PCT_Change
            Open_Price = 0
            Close_Price = 0
            Blank = Blank + 1
        End If
        
    Next i
    Blank = 2
    For i = 2 To 759001
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Total = Total + Cells(i, 7).Value
            Cells(Blank, 9).Value = Cells(i, 1).Value
            Cells(Blank, 12).Value = Total
            Total = 0
            Blank = Blank + 1
        Else
            Total = Total + Cells(i, 7).Value
        End If
        
    Next i
    
    For i = 2 To 3001
        Cells(i, 11) = Format(Cells(i, 11), "0.00%")
        
    Next i
    
    For i = 2 To 3001
        If Cells(i, 10) > 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
        ElseIf Cells(i, 10) < 0 Then
            Cells(i, 10).Interior.ColorIndex = 3
        Else
            Cells(i, 10).Interior.ColorIndex = 2
            
        End If
        
    Next i
    
    
End Sub
